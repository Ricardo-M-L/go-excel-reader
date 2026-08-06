# Engineering log

Notes taken while porting ExcelDataReader to Go, kept for a write-up. These are
the things that were not obvious from the spec and that cost real time or would
have shipped as silent bugs.

## The bug that hid in the test fixtures, not the code

The shared string table is the memory-critical part of an xlsx reader, and the
one the port spends the most design effort on. Every test passed. The peak-heap
test reported 3.6 MB on a 4.6 MB file, which looked like a win.

It was measuring nothing. openpyxl, which generated the fixtures, writes strings
**inline** (`<c t="inlineStr"><is><t>`) rather than into `xl/sharedStrings.xml`.
The fixtures had no string table at all. `shared_strings.go` had never executed.

The tell was the number being *too good*: a 200,000-entry table should have cost
several megabytes, and 3.6 MB was not consistent with having loaded one. Checking
the zip listing took thirty seconds and showed no `sharedStrings.xml` in any
fixture.

Two changes followed. xlsxwriter was added as a second xlsx generator, since it
uses the shared table by default, so both string representations are now covered
by fixtures written by independent tools. And a test now asserts the property the
other tests depend on:

```go
// If a regenerated fixture quietly stopped using the shared string table, the
// string-table reader would go untested while every test still passed.
for _, f := range []string{"sst.xlsx", "large.xlsx"} {
        if !hasPart(f, "xl/sharedStrings.xml") { t.Errorf(...) }
}
if hasPart("simple.xlsx", "xl/sharedStrings.xml") { t.Error(...) }
```

With a real 400,000-entry table the same test reports 26.9 MB — still inside the
50 MB budget, but a number that actually means something.

The general lesson: a test fixture is an input you did not write and did not
read. "All tests pass" says nothing until you know the fixture exercises the
path. Suspiciously good numbers deserve the same scrutiny as bad ones.

## The shared string table: bytes, not strings

The obvious representation of the string table is `[]string`. For 400,000
entries that is 400,000 heap objects, each carrying a 16-byte header, and — the
part that actually hurts — 400,000 pointers the garbage collector must trace on
every cycle.

Instead the table is one contiguous `[]byte` blob plus an `[]int64` offset index.
Two allocations, zero pointers, nothing for the GC to scan. Strings come back by
slicing the blob, which is safe because the blob is never mutated after parsing.

This is most of the difference behind the 40 MB vs 95 MB vs 221 MB figures in the
README.

Left on the table: `At()` still copies when it converts a byte slice to a string.
`unsafe.String` over the immutable blob would make lookups allocation-free and
remove roughly 400,000 allocations per pass. It was not done, because the memory
budget is already met with margin and adding `unsafe` to an otherwise plain Go
library is not worth a benchmark number nobody asked for. Worth revisiting only
if profiling on a real workload says so.

## BIFF8's most delicate detail: the restated flags byte

A BIFF record payload cannot exceed 8224 bytes, so the shared string table
arrives as one `SST` record followed by a run of `CONTINUE` records. Concatenate
them and parse the result and almost everything works.

Almost. When a string's *characters* are split across a `CONTINUE` boundary, the
first byte of the continuation is not character data — it restates the
compression flag, because the second half of the string may be stored
differently from the first. Miss that byte and it is read as text, shifting every
subsequent string by one byte. The table does not fail; it silently desynchronises,
and every string after the first split comes back garbled.

So the record reader keeps the boundary offsets rather than just concatenating:

```go
// contStarts holds the offsets within data at which a CONTINUE record's
// payload begins.
if c.atContinuation() {
        flags, _ := c.u8()
        high = flags&0x01 != 0
}
```

The check sits inside the character loop, which is exactly right: a boundary that
falls *between* strings carries no flags byte, and only a boundary that resumes a
character array does.

This is hard to hit by accident and easy to miss in review, so the fixture
generator writes strings of random length specifically to force the case, and a
unit test includes a negative control — parsing the same bytes *without* the
boundary information and asserting the result comes out wrong. Without that
control, a test asserting `"abcd" == "abcd"` can pass for the wrong reason.

## Dates are the whole game

Nothing in a spreadsheet cell says "date". A date is a `float64` whose *style*
happens to carry a date number format. Resolving it means three hops — cell
`s=` → `cellXfs` index → `numFmtId` → format string — and any broken link
returns `45217` where the user expects a date.

Two traps found while writing it:

**`cellXfs` is not the only `<xf>` list.** `styles.xml` also contains
`cellStyleXfs`, holding named-style definitions. The elements are identical and
appear first. Reading `<xf>` without tracking which parent list you are in binds
every cell to the wrong format.

**Scanning for date letters is not enough.** A format like `"Paid on day "0.0`
contains `d`, `a` and `y`, all inside a quoted literal. A reader that does
`strings.ContainsAny(format, "dmyhs")` turns 1234.5 into a date in the year 2093.
Quoted literals, `\`-escapes, `_` width placeholders, `*` fill characters and
`[Red]`/`[$-409]` brackets all have to be skipped — while `[h]`, `[mm]` and `[ss]`
*are* date tokens, since they mean elapsed time. That trap is a fixture.

## Excel's leap year that never happened

Excel believes 1900 was a leap year, reproducing a Lotus 1-2-3 bug deliberately
for compatibility. Serial 60 is a 29 February that does not exist, and every
serial from 61 onward is one greater than a real day count would give.

Rather than add a correction term, the conversion switches epoch at the boundary:
1899-12-31 below serial 60, 1899-12-30 at or above it. Both branches then check
by hand — serial 1 is 1900-01-01, serial 61 is 1900-03-01 — which a `+1` buried
in an expression does not.

Serial 60 has no correct answer; it maps to 1900-02-28 so the value stays inside
the real calendar instead of rolling into March.

A related fixture trap: the first draft used a `h:mm AM/PM` fixture dated
1899-12-30, right on the boundary, where writers disagree about what serial to
emit. The test would have been asserting openpyxl's opinion rather than a fact.
The fixture moved to a modern date and the boundary is tested directly in
`numfmt`, where the expectation can be stated exactly.

## Verifying against the original, not against myself

The strongest check in the project is not a unit test. Both this library and the
real C# ExcelDataReader dump each fixture to a canonical text form, and the dumps
are compared byte for byte.

The one design decision that made it work: emit numbers as raw IEEE 754 bit
patterns (`%016X`) rather than as text. C# and Go disagree about how to render a
`double` as a string, and that disagreement would otherwise appear as a false
difference on every non-integral value — noise that buries real differences.

All fixtures match exactly, including a 1,000,000-cell workbook and an 8.9 MB
`.xls`. That is a much stronger claim than "my tests pass", and it took about
twenty minutes to set up.

## Two deliberate additions to the specified API

The target interface was specified precisely, and two methods were added anyway:

- **`Err() error`.** A `Read() bool` cursor has nowhere to report a failure
  mid-stream, so a truncated or corrupt file is indistinguishable from a short
  one. `ReadAll` is specified to return an `error`; without `Err` on the
  interface that error could only ever be `nil`.
- **`Cell(col) (Cell, bool)`.** `ReadAll` is specified to return cells carrying
  a `Raw` field. Nothing else on the interface exposes raw text, so it was not
  implementable as written.

Both are additions the specification's own requirements imply. Worth stating
explicitly rather than quietly widening an interface someone specified on
purpose.

## Numbers

| | `.xlsx` 9.2 MB / 1M cells | `.xls` 8.9 MB / 262k cells |
|---|---|---|
| go-excel-reader | 3.2 s, 40 MB RSS | 0.98 s, 34 MB RSS |
| ExcelDataReader | 3.3 s, 95 MB RSS | 1.15 s, 93 MB RSS |
| excelize v2.11.0 | 4.0 s, 221 MB RSS | not supported |

Peak heap measured from inside Go on the same xlsx: 26.9 MB against a 50 MB
budget.
