package excel_test

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"runtime"
	"strings"
	"testing"
	"time"

	"golang.org/x/text/encoding/charmap"

	excel "github.com/jloor/go-excel-reader"
)

func fixture(name string) string { return filepath.Join("testdata", name) }

// sheet is a decoded sheet: its name and its rows of normalised values.
type sheet struct {
	name string
	rows [][]any
}

// normalise makes values comparable with reflect.DeepEqual across formats.
// time.Time carries a monotonic clock and a location that DeepEqual sees but
// callers do not care about, so times collapse to their UTC rendering.
func normalise(v any) any {
	if t, ok := v.(time.Time); ok {
		return "time:" + t.UTC().Format("2006-01-02T15:04:05.999999999")
	}
	return v
}

func at(y int, mo time.Month, d, h, mi, s int) any {
	return normalise(time.Date(y, mo, d, h, mi, s, 0, time.UTC))
}

// drain reads every sheet and row from a reader.
func drain(t *testing.T, r excel.Reader) []sheet {
	t.Helper()
	var out []sheet
	for r.NextSheet() {
		s := sheet{name: r.SheetName()}
		for r.Read() {
			row := make([]any, r.FieldCount())
			for i := range row {
				row[i] = normalise(r.GetValue(i))
			}
			s.rows = append(s.rows, row)
		}
		out = append(out, s)
	}
	if err := r.Err(); err != nil {
		t.Fatalf("streaming error: %v", err)
	}
	return out
}

func openFixture(t *testing.T, name string) excel.Reader {
	t.Helper()
	r, err := excel.OpenFile(fixture(name))
	if err != nil {
		t.Fatalf("OpenFile(%s): %v", name, err)
	}
	t.Cleanup(func() { r.Close() })
	return r
}

// The table shared by simple.xlsx, simple.xls and simple.csv.
var simpleRows = [][]any{
	{"Name", "Qty", "Price"},
	{"Widget", int64(10), 9.99},
	{"Gadget", int64(3), 19.5},
	{"Doohickey", int64(25), 0.5},
}

func TestDetect(t *testing.T) {
	cases := []struct {
		file string
		want excel.Format
	}{
		{"simple.xlsx", excel.FormatXLSX},
		{"simple.xls", excel.FormatXLS},
		{"simple.csv", excel.FormatCSV},
		{"tab.tsv", excel.FormatCSV},
		{"bom.csv", excel.FormatCSV},
	}
	for _, c := range cases {
		f, err := os.Open(fixture(c.file))
		if err != nil {
			t.Fatal(err)
		}
		got, err := excel.Detect(f)
		f.Close()
		if err != nil {
			t.Errorf("Detect(%s): %v", c.file, err)
			continue
		}
		if got != c.want {
			t.Errorf("Detect(%s) = %v, want %v", c.file, got, c.want)
		}
	}
}

// TestDetectRejectsBinary checks that the CSV fallback does not swallow binary
// data. Detection ends with "otherwise it is text", so the guard on that branch
// is the difference between an error and pages of mojibake.
func TestDetectRejectsBinary(t *testing.T) {
	junk := []byte{0x1F, 0x8B, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x03, 0xFF}
	if _, err := excel.Detect(bytes.NewReader(junk)); err == nil {
		t.Error("Detect accepted binary data as text")
	}

	withNUL := append([]byte("name,value\n"), 0x00, 'x')
	if _, err := excel.Detect(bytes.NewReader(withNUL)); err == nil {
		t.Error("Detect accepted data containing a NUL byte as text")
	}
}

func TestXLSXSimple(t *testing.T) {
	got := drain(t, openFixture(t, "simple.xlsx"))
	if len(got) != 1 {
		t.Fatalf("got %d sheets, want 1", len(got))
	}
	if got[0].name != "Sheet1" {
		t.Errorf("sheet name = %q, want Sheet1", got[0].name)
	}
	if !reflect.DeepEqual(got[0].rows, simpleRows) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, simpleRows)
	}
}

// TestXLSXStringStorage covers both ways an xlsx may store text.
//
// openpyxl writes strings inline; xlsxwriter puts them in the shared string
// table and stores an index. The two files hold identical data, so any
// difference in the decoded result is a bug in one of the two paths. Testing
// only one writer leaves the other path unexercised, which is how the shared
// string table went untested here at first.
func TestXLSXStringStorage(t *testing.T) {
	inline := drain(t, openFixture(t, "simple.xlsx"))
	shared := drain(t, openFixture(t, "sst.xlsx"))

	if !reflect.DeepEqual(inline[0].rows, simpleRows) {
		t.Errorf("inline strings: %#v", inline[0].rows)
	}
	if !reflect.DeepEqual(shared[0].rows, simpleRows) {
		t.Errorf("shared strings: %#v", shared[0].rows)
	}
}

// TestFixturesCoverBothStringForms guards the property the tests above rely on.
// If a regenerated fixture quietly stopped using the shared string table, the
// string-table reader would go untested while every test still passed.
func TestFixturesCoverBothStringForms(t *testing.T) {
	// A missing fixture must not skip the assertions that do not need it, so
	// availability is reported rather than turned into a skip of the whole test.
	hasPart := func(file, part string) (present, known bool) {
		zr, err := zip.OpenReader(fixture(file))
		if err != nil {
			return false, false
		}
		defer zr.Close()
		for _, e := range zr.File {
			if e.Name == part {
				return true, true
			}
		}
		return false, true
	}

	for _, f := range []string{"sst.xlsx", "large.xlsx"} {
		switch present, known := hasPart(f, "xl/sharedStrings.xml"); {
		case !known:
			t.Logf("fixture %s unavailable; skipping its check", f)
		case !present:
			t.Errorf("%s has no xl/sharedStrings.xml; the shared string reader is untested", f)
		}
	}
	if present, known := hasPart("simple.xlsx", "xl/sharedStrings.xml"); known && present {
		t.Error("simple.xlsx now uses shared strings; the inline string reader is untested")
	}
}

func TestXLSXTypes(t *testing.T) {
	got := drain(t, openFixture(t, "types.xlsx"))
	want := [][]any{
		{"str", "int", "float", "date", "bool", "empty"},
		{"hello", int64(42), 3.14159, at(2023, 6, 15, 0, 0, 0), true},
		{"world", int64(-7), -0.5, at(1999, 12, 31, 23, 59, 59), false},
	}
	if !reflect.DeepEqual(got[0].rows, want) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, want)
	}
}

func TestXLSXMultiSheet(t *testing.T) {
	r := openFixture(t, "multisheet.xlsx")
	if n := r.SheetCount(); n != 3 {
		t.Fatalf("SheetCount = %d, want 3", n)
	}
	got := drain(t, r)

	wantNames := []string{"Alpha", "Beta", "Gamma With Spaces"}
	for i, w := range wantNames {
		if got[i].name != w {
			t.Errorf("sheet %d name = %q, want %q", i, got[i].name, w)
		}
	}
	if !reflect.DeepEqual(got[0].rows, [][]any{{"a1", "a2"}, {int64(1), int64(2)}}) {
		t.Errorf("Alpha rows = %#v", got[0].rows)
	}
	if !reflect.DeepEqual(got[1].rows, [][]any{{"b1"}, {int64(100)}}) {
		t.Errorf("Beta rows = %#v", got[1].rows)
	}
	if !reflect.DeepEqual(got[2].rows, [][]any{{"g1", "g2", "g3"}}) {
		t.Errorf("Gamma rows = %#v", got[2].rows)
	}
}

// TestXLSXGaps pins down index handling for a sparse sheet: skipped rows must
// still be counted, and column AA must land at index 26 rather than 27.
func TestXLSXGaps(t *testing.T) {
	r := openFixture(t, "gaps.xlsx")
	if !r.NextSheet() {
		t.Fatal("no sheet")
	}

	type snapshot struct {
		index  int
		fields int
		last   any
	}
	var got []snapshot
	for r.Read() {
		s := snapshot{index: r.RowIndex(), fields: r.FieldCount()}
		if s.fields > 0 {
			s.last = r.GetValue(s.fields - 1)
		}
		got = append(got, s)
	}
	if err := r.Err(); err != nil {
		t.Fatal(err)
	}

	want := []snapshot{
		{index: 0, fields: 1, last: "first"},
		{index: 1, fields: 3, last: "third"},       // B2 is a hole inside the row
		{index: 2, fields: 0, last: nil},           // row 3 is missing from the file
		{index: 3, fields: 27, last: "twenty-six"}, // AA is column index 26
		{index: 4, fields: 0, last: nil},
		{index: 5, fields: 2, last: "sixth"},
	}
	if !reflect.DeepEqual(got, want) {
		t.Errorf("row shape mismatch:\n got %#v\nwant %#v", got, want)
	}
}

func TestXLSXDates(t *testing.T) {
	got := drain(t, openFixture(t, "dates.xlsx"))
	want := [][]any{
		{"format", "value"},
		{"m/d/yy", at(2023, 1, 2, 0, 0, 0)},
		{"d-mmm-yy", at(2023, 3, 14, 0, 0, 0)},
		{"yyyy-mm-dd", at(2024, 12, 25, 0, 0, 0)},
		{"h:mm AM/PM", at(2022, 5, 4, 13, 30, 0)},
		{"yyyy-mm-dd hh:mm:ss", at(2020, 2, 29, 8, 5, 3)},
		{"[$-409]mmmm d, yyyy", at(2021, 7, 4, 0, 0, 0)},
		// The format's quoted literal contains "day"; a reader that scans for
		// date letters without honouring quotes turns this number into a date.
		{"not-a-date", 1234.5},
	}
	if !reflect.DeepEqual(got[0].rows, want) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, want)
	}
}

func TestXLSSimple(t *testing.T) {
	got := drain(t, openFixture(t, "simple.xls"))
	if len(got) != 1 {
		t.Fatalf("got %d sheets, want 1", len(got))
	}
	if got[0].name != "Sheet1" {
		t.Errorf("sheet name = %q, want Sheet1", got[0].name)
	}
	if !reflect.DeepEqual(got[0].rows, simpleRows) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, simpleRows)
	}
}

func TestXLSTypes(t *testing.T) {
	got := drain(t, openFixture(t, "types.xls"))
	want := [][]any{
		{"str", "int", "float", "date", "bool", "empty"},
		{"hello", int64(42), 3.14159, at(2023, 6, 15, 0, 0, 0), true},
		{"world", int64(-7), -0.5, at(1999, 12, 31, 23, 59, 59), false},
	}
	if !reflect.DeepEqual(got[0].rows, want) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, want)
	}
}

func TestXLSMultiSheet(t *testing.T) {
	r := openFixture(t, "multisheet.xls")
	if n := r.SheetCount(); n != 3 {
		t.Fatalf("SheetCount = %d, want 3", n)
	}
	got := drain(t, r)
	wantNames := []string{"Alpha", "Beta", "Gamma With Spaces"}
	for i, w := range wantNames {
		if got[i].name != w {
			t.Errorf("sheet %d name = %q, want %q", i, got[i].name, w)
		}
	}
	if !reflect.DeepEqual(got[0].rows, [][]any{{"a1", "a2"}, {int64(1), int64(2)}}) {
		t.Errorf("Alpha rows = %#v", got[0].rows)
	}
}

// TestXLSVsXLSX is the cross-format check that matters most: the same table,
// written by two unrelated producers into two unrelated binary formats, must
// decode to byte-identical values.
func TestXLSVsXLSX(t *testing.T) {
	fromXLSX := drain(t, openFixture(t, "simple.xlsx"))
	fromXLS := drain(t, openFixture(t, "simple.xls"))

	if !reflect.DeepEqual(fromXLSX[0].rows, fromXLS[0].rows) {
		t.Errorf("xlsx and xls disagree:\nxlsx %#v\n xls %#v",
			fromXLSX[0].rows, fromXLS[0].rows)
	}
}

func TestCrossFormatIdentical(t *testing.T) {
	x := drain(t, openFixture(t, "simple.xlsx"))
	s := drain(t, openFixture(t, "simple.xls"))

	f, err := os.Open(fixture("simple.csv"))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	// A CSV carries no types at all, so matching the binary formats requires
	// inference to be switched on explicitly.
	c, err := excel.OpenCSV(f, excel.WithCSVTypeInference(true))
	if err != nil {
		t.Fatal(err)
	}
	defer c.Close()
	csvSheets := drain(t, c)

	if !reflect.DeepEqual(x[0].rows, simpleRows) {
		t.Errorf("xlsx: %#v", x[0].rows)
	}
	if !reflect.DeepEqual(s[0].rows, simpleRows) {
		t.Errorf("xls: %#v", s[0].rows)
	}
	if !reflect.DeepEqual(csvSheets[0].rows, simpleRows) {
		t.Errorf("csv: %#v", csvSheets[0].rows)
	}
}

func TestCSVBasic(t *testing.T) {
	got := drain(t, openFixture(t, "simple.csv"))
	// Without inference every field is text, which is the documented default.
	want := [][]any{
		{"Name", "Qty", "Price"},
		{"Widget", "10", "9.99"},
		{"Gadget", "3", "19.5"},
		{"Doohickey", "25", "0.5"},
	}
	if !reflect.DeepEqual(got[0].rows, want) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, want)
	}
}

func TestCSVTabDelimited(t *testing.T) {
	f, err := os.Open(fixture("tab.tsv"))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	r, err := excel.OpenCSV(f, excel.WithCSVDelimiter('\t'), excel.WithCSVTypeInference(true))
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()

	got := drain(t, r)
	if !reflect.DeepEqual(got[0].rows, simpleRows) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, simpleRows)
	}
}

func TestCSVWithBOM(t *testing.T) {
	got := drain(t, openFixture(t, "bom.csv"))
	if first := got[0].rows[0][0]; first != "Name" {
		t.Errorf("first header = %q, want %q; the BOM was not stripped", first, "Name")
	}
}

func TestCSVQuoted(t *testing.T) {
	got := drain(t, openFixture(t, "quoted.csv"))
	want := [][]any{
		{"a", "b", "c"},
		{"has,comma", `has "quote"`, "has\nnewline"},
		{"plain", nil, "trailing"},
	}
	if !reflect.DeepEqual(got[0].rows, want) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, want)
	}
}

func TestCSVEncoding(t *testing.T) {
	f, err := os.Open(fixture("latin1.csv"))
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()

	r, err := excel.OpenCSV(f, excel.WithCSVEncoding(charmap.ISO8859_1))
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()

	got := drain(t, r)
	want := [][]any{
		{"name", "city"},
		{"José", "Málaga"},
		{"René", "Köln"},
	}
	if !reflect.DeepEqual(got[0].rows, want) {
		t.Errorf("rows mismatch:\n got %#v\nwant %#v", got[0].rows, want)
	}
}

func TestTypedGetters(t *testing.T) {
	r := openFixture(t, "types.xlsx")
	if !r.NextSheet() {
		t.Fatal("no sheet")
	}
	r.Read() // header
	r.Read() // first data row

	if s, ok := r.GetString(0); !ok || s != "hello" {
		t.Errorf("GetString(0) = %q, %v", s, ok)
	}
	if n, ok := r.GetInt(1); !ok || n != 42 {
		t.Errorf("GetInt(1) = %d, %v", n, ok)
	}
	// An integer cell must also read as a float; which of the two a file
	// produces is an accident of how it was written.
	if f, ok := r.GetFloat(1); !ok || f != 42 {
		t.Errorf("GetFloat(1) = %v, %v", f, ok)
	}
	if f, ok := r.GetFloat(2); !ok || f != 3.14159 {
		t.Errorf("GetFloat(2) = %v, %v", f, ok)
	}
	// A fractional value must not silently truncate to an integer.
	if _, ok := r.GetInt(2); ok {
		t.Error("GetInt on 3.14159 succeeded; it must report false")
	}
	if tm, ok := r.GetTime(3); !ok || !tm.Equal(time.Date(2023, 6, 15, 0, 0, 0, 0, time.UTC)) {
		t.Errorf("GetTime(3) = %v, %v", tm, ok)
	}
	if b, ok := r.GetBool(4); !ok || !b {
		t.Errorf("GetBool(4) = %v, %v", b, ok)
	}
	// Wrong-type access reports false rather than coercing.
	if _, ok := r.GetString(1); ok {
		t.Error("GetString on a number succeeded; it must report false")
	}
	if _, ok := r.GetFloat(0); ok {
		t.Error("GetFloat on a string succeeded; it must report false")
	}
	// Out of range and empty behave the same way to a caller.
	if !r.IsEmpty(99) {
		t.Error("IsEmpty(99) = false for an out-of-range column")
	}
	if _, ok := r.Cell(99); ok {
		t.Error("Cell(99) reported ok for an out-of-range column")
	}
}

func TestReadAll(t *testing.T) {
	r := openFixture(t, "multisheet.xlsx")
	sheets, err := excel.ReadAll(r)
	if err != nil {
		t.Fatal(err)
	}
	if len(sheets) != 3 {
		t.Fatalf("got %d sheets, want 3", len(sheets))
	}
	if len(sheets[0]) != 2 {
		t.Fatalf("Alpha has %d rows, want 2", len(sheets[0]))
	}
	if v := sheets[0][1][0].Value; v != int64(1) {
		t.Errorf("Alpha[1][0].Value = %#v, want int64(1)", v)
	}
	// Raw must survive alongside the typed value.
	if raw := sheets[0][1][0].Raw; raw != "1" {
		t.Errorf("Alpha[1][0].Raw = %q, want %q", raw, "1")
	}
}

func TestOpenFileErrors(t *testing.T) {
	if _, err := excel.OpenFile(fixture("does-not-exist.xlsx")); err == nil {
		t.Error("opening a missing file succeeded")
	}

	tmp := filepath.Join(t.TempDir(), "junk.xlsx")
	if err := os.WriteFile(tmp, []byte{0x00, 0x01, 0x02, 0x03}, 0o600); err != nil {
		t.Fatal(err)
	}
	if _, err := excel.OpenFile(tmp); err == nil {
		t.Error("opening binary junk succeeded")
	}
}

// TestXLSXLargeMemory is the streaming guarantee, stated as a test.
//
// It reads a multi-megabyte workbook with a 200,000-entry string table and
// asserts that peak heap stays far below the file size. A reader that
// materialises the sheet, or that stores the string table as a []string, fails
// here rather than in production.
func TestXLSXLargeMemory(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping memory test in short mode")
	}
	path := fixture("large.xlsx")
	info, err := os.Stat(path)
	if err != nil {
		t.Skipf("large fixture unavailable: %v", err)
	}

	runtime.GC()
	var before runtime.MemStats
	runtime.ReadMemStats(&before)

	r, err := excel.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()

	var (
		rows int
		peak uint64
	)
	for r.NextSheet() {
		for r.Read() {
			rows++
			// Touch the values so nothing can be optimised away.
			for i := 0; i < r.FieldCount(); i++ {
				_ = r.GetValue(i)
			}
			if rows%10000 == 0 {
				var m runtime.MemStats
				runtime.ReadMemStats(&m)
				if m.HeapAlloc > peak {
					peak = m.HeapAlloc
				}
			}
		}
	}
	if err := r.Err(); err != nil {
		t.Fatal(err)
	}
	if rows == 0 {
		t.Fatal("read no rows")
	}

	const limit = 50 << 20
	t.Logf("file %.1f MB, %d rows, peak heap %.1f MB",
		float64(info.Size())/(1<<20), rows, float64(peak)/(1<<20))
	if peak > limit {
		t.Errorf("peak heap %.1f MB exceeds the %d MB budget",
			float64(peak)/(1<<20), limit>>20)
	}
}

// TestXLSXMemoryScalesWithStrings pins the reader's actual memory behaviour.
//
// repeat.xlsx and large.xlsx hold identical row and cell counts and differ only
// in how many distinct strings they contain. Rows stream, so the row count
// costs nothing; the shared string table is resident, so it costs everything.
// Asserting the baseline separately from the table is what keeps the README's
// claim honest: if row streaming ever regressed into accumulation, the repeat
// case would blow past this budget while the large case still looked fine.
func TestXLSXMemoryScalesWithStrings(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping memory test in short mode")
	}

	measure := func(t *testing.T, name string) (rows int, peak uint64) {
		t.Helper()
		path := fixture(name)
		if _, err := os.Stat(path); err != nil {
			t.Skipf("fixture %s unavailable: %v", name, err)
		}

		runtime.GC()
		r, err := excel.OpenFile(path)
		if err != nil {
			t.Fatal(err)
		}
		defer r.Close()

		for r.NextSheet() {
			for r.Read() {
				rows++
				for i := 0; i < r.FieldCount(); i++ {
					_ = r.GetValue(i)
				}
				if rows%10000 == 0 {
					var m runtime.MemStats
					runtime.ReadMemStats(&m)
					if m.HeapAlloc > peak {
						peak = m.HeapAlloc
					}
				}
			}
		}
		if err := r.Err(); err != nil {
			t.Fatal(err)
		}
		return rows, peak
	}

	rows, peak := measure(t, "repeat.xlsx")
	t.Logf("repeat.xlsx: %d rows, 100 distinct strings, peak heap %.1f MB",
		rows, float64(peak)/(1<<20))

	// A million cells with a hundred distinct strings must stay near the
	// streaming baseline. Accumulating rows would put this in the hundreds
	// of megabytes.
	const baselineBudget = 20 << 20
	if peak > baselineBudget {
		t.Errorf("peak heap %.1f MB exceeds the %d MB streaming baseline; "+
			"rows are no longer streaming",
			float64(peak)/(1<<20), baselineBudget>>20)
	}
}

// TestXLSLargeSharedStrings exercises the BIFF shared string table across
// CONTINUE record boundaries, where a split string restates its compression
// flag. The fixture uses randomly sized strings so that the split lands mid
// string; without the restated flag every string after the first boundary
// comes back garbled.
func TestXLSLargeSharedStrings(t *testing.T) {
	if testing.Short() {
		t.Skip("skipping large fixture test in short mode")
	}
	path := fixture("large.xls")
	if _, err := os.Stat(path); err != nil {
		t.Skipf("large fixture unavailable: %v", err)
	}

	r, err := excel.OpenFile(path)
	if err != nil {
		t.Fatal(err)
	}
	defer r.Close()

	if !r.NextSheet() {
		t.Fatal("no sheet")
	}

	rows := 0
	for r.Read() {
		idx := r.RowIndex()
		got, ok := r.GetString(0)
		if !ok {
			t.Fatalf("row %d: column 0 is not a string: %#v", idx, r.GetValue(0))
		}
		// Every string was written as "s<row>-xxxx"; the prefix proves the
		// table stayed aligned across every CONTINUE boundary.
		want := fmt.Sprintf("s%d-", idx)
		if !strings.HasPrefix(got, want) {
			t.Fatalf("row %d: string table desynchronised: got %q, want prefix %q",
				idx, got, want)
		}
		if n, ok := r.GetInt(1); !ok || n != int64(idx) {
			t.Fatalf("row %d: column 1 = %v (ok=%v), want %d", idx, n, ok, idx)
		}
		rows++
	}
	if err := r.Err(); err != nil {
		t.Fatal(err)
	}
	if rows < 1000 {
		t.Fatalf("read only %d rows; the fixture should have tens of thousands", rows)
	}
	t.Logf("verified %d rows of BIFF shared strings", rows)
}

func BenchmarkXLSXLarge(b *testing.B) {
	path := fixture("large.xlsx")
	if _, err := os.Stat(path); err != nil {
		b.Skipf("large fixture unavailable: %v", err)
	}

	b.ReportAllocs()
	for b.Loop() {
		r, err := excel.OpenFile(path)
		if err != nil {
			b.Fatal(err)
		}
		for r.NextSheet() {
			for r.Read() {
				for i := 0; i < r.FieldCount(); i++ {
					_ = r.GetValue(i)
				}
			}
		}
		if err := r.Err(); err != nil {
			b.Fatal(err)
		}
		r.Close()
	}
}
