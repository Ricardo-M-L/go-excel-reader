# go-excel-reader

A streaming reader for `.xlsx`, `.xls` (BIFF8) and `.csv` files. Pure Go, no CGo,
one dependency (`golang.org/x/text`, used only for CSV character encodings).

It is a port of the core of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)
(C#, MIT) and keeps that library's cursor model: advance a sheet, advance a row,
read cells by column index. Nothing larger than the current row is held in
memory.

```bash
go get github.com/jloor/go-excel-reader
```

## Quick start

```go
package main

import (
	"fmt"
	"log"

	excel "github.com/jloor/go-excel-reader"
)

func main() {
	r, err := excel.OpenFile("report.xlsx") // format detected from content
	if err != nil {
		log.Fatal(err)
	}
	defer r.Close()

	for r.NextSheet() {
		fmt.Println("sheet:", r.SheetName())
		for r.Read() {
			for i := 0; i < r.FieldCount(); i++ {
				fmt.Printf("  [%d] %v\n", i, r.GetValue(i))
			}
		}
	}
	if err := r.Err(); err != nil { // errors surface here, not from Read
		log.Fatal(err)
	}
}
```

## Supported formats

| Format | Container | Detected by | Status |
|---|---|---|---|
| `.xlsx` | ZIP of XML | `PK\x03\x04` + `xl/workbook.xml` | Supported |
| `.xls` | CFB / OLE2 | `D0CF11E0A1B11AE1` | Supported (BIFF8 only) |
| `.csv` / `.tsv` | plain text | fallback, after a binary check | Supported |
| `.xlsb` | ZIP of binary | `xl/workbook.bin` | Rejected with a clear error |

Detection reads the content, not the file extension, because the extension is
the least reliable thing about a spreadsheet.

## API

```go
func OpenFile(path string) (Reader, error)
func OpenReader(r io.ReadSeeker) (Reader, error)
func OpenXLSX(r io.ReadSeeker) (Reader, error)
func OpenXLS(r io.ReadSeeker) (Reader, error)
func OpenCSV(r io.Reader, opts ...CSVOption) (Reader, error)
func ReadAll(r Reader) ([][]Row, error)
func Detect(rs io.ReadSeeker) (Format, error)
```

| Method | Returns |
|---|---|
| `NextSheet() bool` | advance a sheet; false when done |
| `SheetName() string` | current sheet name |
| `SheetCount() int` | number of sheets |
| `Read() bool` | advance a row; false at end of sheet |
| `FieldCount() int` | cells in the current row |
| `GetValue(col) any` | `string`, `int64`, `float64`, `bool`, `time.Time` or `nil` |
| `GetString/GetFloat/GetInt/GetBool/GetTime(col)` | typed value plus an ok flag |
| `IsEmpty(col) bool` | whether the cell holds no value |
| `Cell(col) (Cell, bool)` | typed value **and** the raw text it came from |
| `RowIndex() int` | zero-based row index, counting rows the file omitted |
| `Err() error` | first error encountered while streaming |
| `Close() error` | release resources |

CSV options: `WithCSVDelimiter(rune)`, `WithCSVEncoding(encoding.Encoding)`,
`WithCSVTypeInference(bool)`.

### Type mapping

| Cell | Go type |
|---|---|
| text | `string` |
| whole number | `int64` |
| number with a fractional part | `float64` |
| date- or time-formatted number | `time.Time` (always UTC) |
| boolean | `bool` |
| formula error (`#DIV/0!`) | `string` |
| empty | `nil` |

Typed getters are strict: `GetString` on a number reports `false` rather than
formatting it. Only `int64` and `float64` interconvert, because whether a number
lands as one or the other depends on how the file happened to be written.

Dates are returned in UTC. Excel serials carry no timezone, so attaching a local
zone would make the same file decode differently on different machines.

## Memory

The design goal is that peak memory tracks the widest row, not the file size.
Two things make that work: worksheets are parsed as a token stream and never
materialised, and the shared string table is stored as one contiguous byte blob
with an offset index instead of a `[]string`. A `[]string` of 400,000 entries is
400,000 separate allocations the garbage collector must trace on every cycle; a
`[]byte` is one.

Measured on `testdata/large.xlsx` — 9.2 MB on disk, 200,000 rows, 1,000,000
cells, with a 15.4 MB shared string table (peak RSS, best of two runs):

| Reader | Time | Peak RSS |
|---|---|---|
| **go-excel-reader** | **3.2 s** | **40 MB** |
| ExcelDataReader 3.9.0 (.NET 9) | 3.3 s | 95 MB |
| excelize v2.11.0 | 4.0 s | 221 MB |

The excelize figure is not quite like for like: its `Rows()` cursor returns
strings, so it does less type resolution than the other two. The memory column
is the meaningful comparison.

## Correctness

The reader is checked against the original C# implementation rather than only
against its own expectations. Both libraries dump each fixture to a canonical
form — numbers as raw IEEE 754 bit patterns, so that C# and Go string formatting
cannot introduce false differences — and the dumps are compared byte for byte.

All fixtures match exactly, including the 1,000,000-cell workbook and an
8.9 MB `.xls` whose shared string table spans hundreds of `CONTINUE` records.

Test fixtures are produced by three independent writers (openpyxl, xlsxwriter
and xlwt) so that the tests exercise the format rather than agreeing with a
writer of our own. openpyxl and xlsxwriter store strings differently — inline
versus shared — and a test asserts that both representations remain covered.

## Not supported in v0.1.0

- `.xlsb` (OpenXML Binary) — detected and rejected with a clear error
- Password-protected or encrypted workbooks
- BIFF5 and earlier `.xls` — rejected with a clear error rather than decoded
  into mojibake
- Fraction number formats (`# ?/?`) in `numfmt.FormatValue`
- Cell styling, merged cells, formulas as source text (cached results are read)

## Development

```bash
go test ./...          # small fixtures are committed
go test -short ./...   # skip the large-fixture tests
```

The two large fixtures (~18 MB) are generated rather than committed, so that
`go get` does not download them. Tests that need them skip when absent:

```bash
python3 -m venv .venv && .venv/bin/pip install openpyxl xlsxwriter xlwt
.venv/bin/python testdata/gen_fixtures.py testdata/
```

## License

MIT. ExcelDataReader is also MIT; its copyright notice is retained in `LICENSE`.
