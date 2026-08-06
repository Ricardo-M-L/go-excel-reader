# Chapter 2 — Streaming xlsx, and the table everything turns on

*2026-08-06. The xlsx reader. Worked on the first test run, which turned out to
mean less than it sounded like.*

An `.xlsx` is a ZIP of XML. The parts that matter:

```
xl/workbook.xml          sheet names and relationship ids
xl/_rels/workbook.xml.rels   ids → part paths
xl/styles.xml            number formats (chapter 1)
xl/sharedStrings.xml     the string table
xl/worksheets/sheet1.xml the actual cells
```

Two of those are unbounded: the worksheet and the string table. Everything about
the memory behaviour of this library is decided by how those two are handled.

## The worksheet: never build a tree

The worksheet is parsed with `xml.Decoder` token by token, one row at a time,
reusing the row buffer between rows. It is never materialised.

This part is not clever, it is just discipline. `encoding/xml`'s `Unmarshal` on a
26 MB sheet would allocate the whole document; `Decoder.Token()` allocates a
token. The cost is that you write a small state machine by hand instead of
declaring a struct.

Details that cost time:

- **Column letters are bijective base-26.** There is no zero digit: `A`–`Z` are
  1–26 and `AA` is 27, which maps to zero-based index 26. Ordinary base-26 puts
  `AA` at 27 and shifts every column past `Z` by one. `XFD`, Excel's last column,
  is index 16383 and is a test case.
- **`r` attributes are optional.** Streaming producers omit them, so both row and
  column positions need an implicit fallback. One corpus file later turned out to
  be entirely composed of this case.
- **`<rPh>` contains `<t>`.** Those are furigana pronunciation hints for East
  Asian text. Concatenating every `<t>` under an `<is>` splices pronunciation
  guides into the middle of the value.
- **Gaps are real.** Cells omitted from the file are holes; rows omitted from the
  file are missing rows. Both are filled so that column index and `RowIndex()`
  match what the user sees in Excel.

## The string table: bytes, not strings

Almost every string in a workbook lives in one shared table, and each cell stores
an index. A workbook with a million country names stores each distinct name once.

Good for file size. Dangerous for a reader, because the obvious representation is
`[]string`, and for 400,000 entries that is:

- 400,000 separate heap allocations
- 400,000 string headers, 16 bytes each, before any actual text
- 400,000 pointers the garbage collector traces on **every** cycle

Instead the table is one contiguous `[]byte` blob plus an `[]int64` offset index.
Two allocations. Zero pointers. The GC never traces into a `[]byte`, so a
400,000-entry table stops being 400,000 objects it has to walk.

```go
func (s *SharedStrings) At(i int) (string, bool) {
        if s == nil || i < 0 || i >= s.Len() { return "", false }
        return string(s.blob[s.offs[i]:s.offs[i+1]]), true
}
```

The spec suggested indexing byte offsets and lazily re-reading very large tables
from disk. That is not actually available: the table lives inside a *compressed*
zip member, so seeking backwards means re-inflating from the start — O(n²) across
a full pass. The blob gets most of the benefit with none of that.

This is the single decision most responsible for 40 MB against 95 MB and 221 MB.

### What was left on the table

`At` still copies. `unsafe.String` over the immutable blob would make lookups
allocation-free and remove ~400,000 allocations per pass.

Not done. The budget is met with margin, and adding `unsafe` to an otherwise plain
Go library costs something real for anyone who audits dependencies, in exchange
for a number nobody asked for. Recorded rather than silently skipped, so the
option survives.

## Lazy by default

The string table is loaded on first use, not at open. A caller reading only
numeric columns never pays for it. Three lines of bookkeeping.

## And then it all passed

`TestXLSXSimple`, `TestXLSXTypes`, `TestXLSXMultiSheet`, `TestXLSXGaps`,
`TestXLSXDates`, and the CSV suite — green on the first run. The memory benchmark
reported 3.6 MB on a 4.6 MB file, comfortably inside a 50 MB budget.

Every word of this chapter about the string table was, at that moment,
unverified. See chapter 3.
