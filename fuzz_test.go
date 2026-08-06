package excel_test

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"

	excel "github.com/jloor/go-excel-reader"
)

// Fuzz targets for the container parsers.
//
// Both binary formats do a great deal of arithmetic on lengths and offsets read
// straight out of the file: sector numbers, FAT chains, record payload sizes,
// shared string counts, column indices. Every one of those is attacker-supplied
// when a service accepts spreadsheet uploads, and the failure mode for a Go
// library is a panic that takes down the caller rather than an error it can
// handle.
//
// The property asserted here is narrow and absolute: for any input at all,
// reading either returns an error or produces rows, and never panics, hangs, or
// allocates without bound. Correctness on garbage is not the goal; refusing it
// safely is.
//
//	go test -fuzz=FuzzOpenReader -fuzztime=5m .

// seedFromFixtures adds the committed fixtures as fuzzing seeds.
//
// Seeding with real files matters more here than in most targets. A random byte
// string is rejected by magic-byte detection in microseconds and exercises
// nothing; a *mutated* valid workbook still reaches the ZIP directory, the CFB
// FAT and the BIFF record loop, which is where the arithmetic lives.
func seedFromFixtures(f *testing.F) {
	f.Helper()
	matches, err := filepath.Glob(filepath.Join("testdata", "*"))
	if err != nil {
		return
	}
	for _, path := range matches {
		switch filepath.Ext(path) {
		case ".xlsx", ".xls", ".csv", ".tsv":
		default:
			continue
		}
		info, err := os.Stat(path)
		if err != nil || info.Size() > 1<<20 {
			continue // keep the corpus small enough to mutate quickly
		}
		if b, err := os.ReadFile(path); err == nil {
			f.Add(b)
		}
	}
}

// FuzzOpenReader drives the full public API over arbitrary bytes.
func FuzzOpenReader(f *testing.F) {
	seedFromFixtures(f)
	f.Add([]byte("PK\x03\x04"))
	f.Add([]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1})
	f.Add([]byte("a,b,c\n1,2,3\n"))

	f.Fuzz(func(t *testing.T, data []byte) {
		r, err := excel.OpenReader(bytes.NewReader(data))
		if err != nil {
			return
		}
		defer r.Close()

		// Bound the walk. A corrupt file can legitimately describe an enormous
		// sheet, and the point of this target is to find panics, not to wait
		// for a hostile row count to finish.
		const maxSheets, maxRows, maxCols = 8, 20000, 4096

		sheets := 0
		for r.NextSheet() && sheets < maxSheets {
			sheets++
			_ = r.SheetName()

			rows := 0
			for r.Read() && rows < maxRows {
				rows++
				n := r.FieldCount()
				if n > maxCols {
					n = maxCols
				}
				for i := range n {
					_ = r.GetValue(i)
					_, _ = r.GetString(i)
					_, _ = r.GetFloat(i)
					_, _ = r.GetInt(i)
					_, _ = r.GetBool(i)
					_, _ = r.GetTime(i)
					_ = r.IsEmpty(i)
					_, _ = r.Cell(i)
				}
				// Out-of-range access must be an absence, never a panic.
				_ = r.GetValue(-1)
				_ = r.GetValue(n + 1_000_000)
				if !r.IsEmpty(-1) {
					t.Fatal("IsEmpty(-1) reported a value")
				}
			}
		}
		_ = r.Err()
	})
}

// FuzzDetect drives format detection alone.
//
// Detection runs before any parser and decides which one runs, so a panic here
// is reachable by any input at all -- including inputs the readers would have
// refused.
func FuzzDetect(f *testing.F) {
	seedFromFixtures(f)
	f.Add([]byte{})
	f.Add([]byte{0x00})
	f.Add([]byte("PK\x03\x04garbage"))

	f.Fuzz(func(t *testing.T, data []byte) {
		format, err := excel.Detect(bytes.NewReader(data))
		if err != nil {
			return
		}
		switch format {
		case excel.FormatXLSX, excel.FormatXLS, excel.FormatCSV:
		default:
			t.Fatalf("Detect returned %v with no error", format)
		}
	})
}

// FuzzOpenXLS targets the compound file and BIFF layers directly.
//
// Going through OpenXLS skips detection, so mutations that break the CFB magic
// bytes still reach the header parser instead of being turned away.
func FuzzOpenXLS(f *testing.F) {
	for _, name := range []string{"simple.xls", "types.xls", "multisheet.xls"} {
		if b, err := os.ReadFile(filepath.Join("testdata", name)); err == nil {
			f.Add(b)
		}
	}

	f.Fuzz(func(t *testing.T, data []byte) {
		r, err := excel.OpenXLS(bytes.NewReader(data))
		if err != nil {
			return
		}
		defer r.Close()

		sheets := 0
		for r.NextSheet() && sheets < 8 {
			sheets++
			rows := 0
			for r.Read() && rows < 20000 {
				rows++
				n := r.FieldCount()
				if n > 4096 {
					n = 4096
				}
				for i := range n {
					_ = r.GetValue(i)
				}
			}
		}
		_ = r.Err()
	})
}
