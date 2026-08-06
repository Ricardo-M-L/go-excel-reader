package excel_test

import (
	"bufio"
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"math"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"
	"time"

	excel "github.com/jloor/go-excel-reader"
)

// This file pins the reader's behaviour against ExcelDataReader's own test
// corpus -- 303 real spreadsheets the upstream maintainers collected because
// something about each of them broke a reader.
//
// The corpus itself is not vendored; it lives in the ExcelDataReader repository
// and is several megabytes of binaries with their own provenance. Point
// EXCEL_CORPUS at a checkout to run these tests:
//
//	git clone --depth 1 https://github.com/ExcelDataReader/ExcelDataReader
//	EXCEL_CORPUS=ExcelDataReader/src/TestData go test -run TestCorpus ./...
//
// What is committed is testdata/corpus_baseline.tsv: for every corpus file, the
// verdict from comparing our output with ExcelDataReader's, plus a hash of *our*
// output. Hashing our own output is what lets the check run without .NET
// installed -- the hashes were derived from a run where the comparison was
// actually performed, so pinning them pins the parity result too.

const corpusBaseline = "testdata/corpus_baseline.tsv"

type baselineRow struct {
	verdict string
	file    string
	sha     string
}

func loadBaseline(t *testing.T) []baselineRow {
	t.Helper()
	f, err := os.Open(corpusBaseline)
	if err != nil {
		t.Fatalf("reading baseline: %v", err)
	}
	defer f.Close()

	var rows []baselineRow
	sc := bufio.NewScanner(f)
	sc.Buffer(make([]byte, 1<<20), 1<<20)
	for sc.Scan() {
		line := sc.Text()
		if strings.HasPrefix(line, "#") || strings.TrimSpace(line) == "" {
			continue
		}
		parts := strings.Split(line, "\t")
		if len(parts) != 3 {
			t.Fatalf("malformed baseline line: %q", line)
		}
		rows = append(rows, baselineRow{verdict: parts[0], file: parts[1], sha: parts[2]})
	}
	if err := sc.Err(); err != nil {
		t.Fatal(err)
	}
	return rows
}

// TestCorpusBaselineShape guards the baseline itself. It needs no corpus.
//
// The number that matters is how many files we decode identically to the
// reference implementation. Writing it down means a change that quietly trades
// agreement for something else has to say so.
func TestCorpusBaselineShape(t *testing.T) {
	rows := loadBaseline(t)

	counts := map[string]int{}
	for _, r := range rows {
		counts[r.verdict]++
	}

	want := map[string]int{
		"IDENTICAL":      214, // byte-identical to ExcelDataReader
		"DIFFER":         29,  // both read it, output disagrees
		"BOTH_REJECT":    33,  // both refuse: xlsb, encrypted, corrupt
		"ONLY_EDR_READS": 27,  // our gaps: BIFF5 and earlier, and open bugs
	}
	for verdict, n := range want {
		if counts[verdict] != n {
			t.Errorf("baseline has %d %s rows, expected %d", counts[verdict], verdict, n)
		}
	}

	// Nothing may read in this library that the reference implementation
	// refuses. That direction would mean we are inventing data.
	if n := counts["ONLY_GO_READS"]; n != 0 {
		t.Errorf("%d files parse here but not in ExcelDataReader; "+
			"reading a file the reference rejects means inventing data", n)
	}

	if len(rows) != 303 {
		t.Errorf("baseline covers %d files, expected 303", len(rows))
	}
}

// TestCorpusParity re-derives every dump and compares it with the baseline.
func TestCorpusParity(t *testing.T) {
	dir := os.Getenv("EXCEL_CORPUS")
	if dir == "" {
		t.Skip("set EXCEL_CORPUS to an ExcelDataReader src/TestData checkout")
	}
	if testing.Short() {
		t.Skip("skipping corpus test in short mode")
	}

	rows := loadBaseline(t)
	var changed []string

	// Regeneration goes through this same function rather than a separate
	// script, so the recorded hashes cannot drift from the ones checked here
	// for reasons that have nothing to do with the reader.
	regen := os.Getenv("EXCEL_CORPUS_REGEN") == "1"

	for i, r := range rows {
		path := filepath.Join(dir, r.file)
		if _, err := os.Stat(path); err != nil {
			t.Logf("missing from corpus, skipping: %s", r.file)
			continue
		}

		got := dumpFile(path)
		if regen {
			rows[i].sha = got
			continue
		}
		if got != r.sha {
			changed = append(changed, fmt.Sprintf("%s (%s): %s -> %s",
				r.file, r.verdict, short(r.sha), short(got)))
		}
	}

	if regen {
		writeBaseline(t, rows)
		t.Log("baseline regenerated; re-verify against ExcelDataReader before committing")
		return
	}

	sort.Strings(changed)
	for _, c := range changed {
		t.Errorf("output changed: %s", c)
	}
	if len(changed) > 0 {
		t.Logf("%d of %d corpus files changed. If deliberate, regenerate the "+
			"baseline with journey/corpus_parity.sh and re-verify against "+
			"ExcelDataReader before committing.", len(changed), len(rows))
	}
}

// writeBaseline rewrites the committed baseline, preserving the verdict column.
//
// The verdicts come from an actual comparison against ExcelDataReader and
// cannot be recomputed without .NET, so regeneration only ever refreshes the
// hashes. A change in what we agree on has to be re-measured deliberately.
func writeBaseline(t *testing.T, rows []baselineRow) {
	t.Helper()

	var b strings.Builder
	b.WriteString("# Parity baseline against the ExcelDataReader test corpus.\n")
	b.WriteString("# Regenerate: EXCEL_CORPUS=<TestData> EXCEL_CORPUS_REGEN=1 go test -run TestCorpusParity .\n")
	b.WriteString("#\n")
	b.WriteString("# verdict = result of comparing our canonical dump with ExcelDataReader 3.9.0's.\n")
	b.WriteString("#           Refreshing hashes does not re-derive it; that needs the .NET harness.\n")
	b.WriteString("# sha256  = hash of OUR dump, so this check runs without .NET installed.\n")
	b.WriteString("#\n")
	b.WriteString("# verdict\tfile\tsha256\n")
	for _, r := range rows {
		fmt.Fprintf(&b, "%s\t%s\t%s\n", r.verdict, r.file, r.sha)
	}

	if err := os.WriteFile(corpusBaseline, []byte(b.String()), 0o644); err != nil {
		t.Fatalf("writing baseline: %v", err)
	}
}

func short(sha string) string {
	if len(sha) > 12 {
		return sha[:12]
	}
	return sha
}

// dumpFile renders a spreadsheet to the canonical form the parity harness uses,
// and returns its hash. A file we refuse hashes to "REJECTED", so that gaining
// or losing the ability to read something shows up as a change.
//
// Numbers are emitted as raw IEEE 754 bit patterns rather than as text. C# and
// Go disagree about how to render a float64, and that disagreement would appear
// as a difference on every non-integral value, burying the real ones.
func dumpFile(path string) string {
	r, err := excel.OpenFile(path)
	if err != nil {
		return "REJECTED"
	}
	defer r.Close()

	h := sha256.New()
	sheetIndex := 0
	for r.NextSheet() {
		name := escapeDump(r.SheetName())
		rowIndex := 0
		for r.Read() {
			for col := 0; col < r.FieldCount(); col++ {
				v := r.GetValue(col)
				if v == nil {
					continue
				}
				tag, text := renderDump(v)
				fmt.Fprintf(h, "%d\t%s\t%d\t%d\t%s\t%s\n",
					sheetIndex, name, rowIndex, col, tag, text)
			}
			rowIndex++
		}
		sheetIndex++
	}
	if r.Err() != nil {
		return "REJECTED"
	}
	return hex.EncodeToString(h.Sum(nil))
}

func renderDump(v any) (tag, text string) {
	switch x := v.(type) {
	case string:
		return "S", escapeDump(x)
	case bool:
		if x {
			return "B", "true"
		}
		return "B", "false"
	case time.Time:
		return "D", x.UTC().Format("2006-01-02 15:04:05.000")
	case float64:
		return "N", fmt.Sprintf("%016X", math.Float64bits(x))
	case int64:
		return "N", fmt.Sprintf("%016X", math.Float64bits(float64(x)))
	default:
		return "S", escapeDump(fmt.Sprint(x))
	}
}

var dumpEscaper = strings.NewReplacer("\\", "\\\\", "\t", "\\t", "\r", "\\r", "\n", "\\n")

func escapeDump(s string) string { return dumpEscaper.Replace(s) }
