# go-excel-reader

A streaming reader for `.xlsx`, `.xls` (BIFF8) and `.csv` files. Pure Go, no CGo,
one dependency (`golang.org/x/text`, used only for CSV character encodings).

It is a port of the core of [ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader)
(C#, MIT) and keeps that library's cursor model: advance a sheet, advance a row,
read cells by column index. Rows stream, so row count does not affect memory.

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
func OpenFile(path string, opts ...Option) (Reader, error)
func OpenReader(r io.ReadSeeker, opts ...Option) (Reader, error)
func OpenXLSX(r io.ReadSeeker, opts ...Option) (Reader, error)
func OpenXLS(r io.ReadSeeker, opts ...Option) (Reader, error)
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

Options: `WithErrorValues(bool)` on any format. For delimited text:
`WithCSVDelimiter(rune)`, `WithCSVEncoding(encoding.Encoding)`,
`WithCSVTypeInference(bool)`.

### Type mapping

| Cell | Go type |
|---|---|
| text | `string` |
| whole number | `int64` |
| number with a fractional part | `float64` |
| date- or time-formatted number | `time.Time` (always UTC) |
| boolean | `bool` |
| formula error (`#DIV/0!`) | `nil`; `WithErrorValues(true)` makes it a `string` |
| empty | `nil` |

Typed getters are strict: `GetString` on a number reports `false` rather than
formatting it. Only `int64` and `float64` interconvert, because whether a number
lands as one or the other depends on how the file happened to be written.

Dates are returned in UTC. Excel serials carry no timezone, so attaching a local
zone would make the same file decode differently on different machines.

## Memory

Rows stream. Worksheets are parsed as an XML token stream and never
accumulated, so row count does not affect peak memory. The shared string table
is the one structure held fully in memory, and it is the only part that grows
with the file.

The difference is measurable. These two workbooks have **identical** row and
cell counts and differ only in how many distinct strings they contain:

| Workbook | Rows | Cells | Distinct strings | Peak Go heap | Peak RSS |
|---|---|---|---|---|---|
| `repeat.xlsx` (5.3 MB) | 200,000 | 1,000,000 | 100 | **3.5 MB** | 9.1 MB |
| `large.xlsx` (9.2 MB) | 200,000 | 1,000,000 | 400,000 | **26.9 MB** | 39.3 MB |

A million cells decode in 3.5 MB of heap when the values repeat. Everything
above that baseline is the string table. Since most real spreadsheets repeat
their values heavily, memory usually stays near the baseline no matter how long
the file is — and the row count never affects it at all.

A test asserts the baseline, so a regression from streaming rows to accumulating
them fails the build rather than quietly showing up as a bigger number.

That table is stored as one contiguous byte blob with an offset index rather
than a `[]string`. A `[]string` of 400,000 entries is 400,000 separate
allocations the garbage collector must trace on every cycle; a `[]byte` is one.
That is most of the difference below.

Measured on `testdata/large.xlsx` — the unfavourable case above, with a 15.4 MB
shared string table (peak RSS, best of two runs):

| Reader | Time | Peak RSS |
|---|---|---|
| **go-excel-reader** | **3.2 s** | **40 MB** |
| ExcelDataReader 3.9.0 (.NET 9) | 3.3 s | 95 MB |
| excelize v2.11.0 | 4.0 s | 221 MB |

The excelize figure is not quite like for like: its `Rows()` cursor returns
strings, so it does less type resolution than the other two. The memory column
is the meaningful comparison.

One caveat worth stating plainly: this is not a constant-memory reader. A
workbook with a very large number of *distinct* strings will hold that table in
memory, and no amount of streaming elsewhere avoids it. What is avoided is
paying three times over for it.

## Correctness

The reader is checked against the original C# implementation rather than only
against its own expectations. Both libraries dump a file to a canonical form —
numbers as raw IEEE 754 bit patterns, so that C# and Go string formatting cannot
introduce false differences — and the dumps are compared byte for byte.

That comparison runs over **ExcelDataReader's own 303-file test corpus**: real
spreadsheets its maintainers collected because each one broke something.

| Result | Files |
|---|---|
| Byte-identical to ExcelDataReader | **214** |
| Both readers refuse (xlsb, encrypted, corrupt) | 33 |
| Both read it, output differs | 29 |
| Only ExcelDataReader reads it | 27 |
| Only this library reads it | **0** |

88% of the files both readers accept decode identically. The last row is the one
that would worry me: nothing here parses a file the reference implementation
rejects, which would mean inventing data.

The 27 gaps are mostly BIFF5-and-earlier `.xls` and encrypted workbooks, both
refused deliberately. The 29 differences are tracked in
`testdata/corpus_baseline.tsv`, and a test pins the counts so that trading
agreement for something else has to be a deliberate act.

Running it needs a corpus checkout; the committed baseline hashes let the
regression check run without .NET:

```bash
git clone --depth 1 https://github.com/ExcelDataReader/ExcelDataReader /tmp/edr
EXCEL_CORPUS=/tmp/edr/src/TestData go test -run TestCorpus ./...
```

The project's own fixtures are produced by three independent writers (openpyxl,
xlsxwriter and xlwt) so that the tests exercise the format rather than agreeing
with a writer of our own. openpyxl and xlsxwriter store strings differently —
inline versus shared — and a test asserts that both representations stay covered.

### Known intentional differences

| Case | ExcelDataReader | Here |
|---|---|---|
| Formula error cell (`#DIV/0!`) | no value | no value; `WithErrorValues(true)` returns the text |
| Raw cell text | not exposed | always available via `Cell(col).Raw` |
| Mid-stream failure | throws | reported by `Err()` |

## Not supported in v0.1.0

- `.xlsb` (OpenXML Binary) — detected and rejected with a clear error
- Password-protected or encrypted workbooks
- BIFF5 and earlier `.xls` — rejected with a clear error rather than decoded
  into mojibake. This accounts for most of the 27 corpus files ExcelDataReader
  reads and this library does not.
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
