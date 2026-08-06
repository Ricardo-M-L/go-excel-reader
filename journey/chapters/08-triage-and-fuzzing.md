# Chapter 8 — Triaging 29 disagreements, and letting a machine write the inputs

*2026-08-06. The corpus said 214 of 303 matched. This is what the other 89 were.*

Chapter 7 ended with a number and a shrug: 214 identical, 29 differing, 27
refused. "Tracked in a baseline, not triaged one by one" is an honest thing to
write and a slightly cowardly place to stop.

So: what are the 29?

## Sorting them before reading them

The first useful move was not to open a single file. It was to classify all 29
mechanically — do the two dumps name the *same cells* with different values, or
different cells entirely?

```
SAME_CELLS  14   typing bugs: right cell, wrong value
DIFF_CELLS  15   structural: cells present in one and not the other
```

Fourteen files where the disagreement is about what a value *is*, fifteen where
it is about what exists. Two completely different investigations, and knowing
which is which before reading any of them saved reading most of them.

## The harness was wrong first

Before blaming the reader, the diffs looked like this:

```
< 0  Sheet1  0  5  S  05/07/2009 11:01:02 -04:00
> 0  Sheet1  0  5  D  2009-05-07 11:01:02.000
```

`S` versus `D`. My C# dumper handled `string`, `bool`, `DateTime`, `double` and
`int`, and sent everything else through a `default:` branch that called
`Convert.ToString`. ExcelDataReader also returns `DateTimeOffset` and `TimeSpan`.

So the comparison was rendering two of the reference's types through a path that
guaranteed a mismatch. Some fraction of "differences" were my measuring
instrument.

Fixed — `DateTimeOffset` renders as UTC, `TimeSpan` as total seconds under its
own tag — and re-ran. **Zero of the 29 resolved.** The harness was wrong *and*
the differences were real. Worth knowing separately.

## The finding I did not expect

With clean tags, a pattern appeared:

```
< 0  Sheet1  0  4  D  2009-01-01 16:00:00.000
> 0  Sheet1  0  4  D  2009-01-01 11:00:00.000
```

Five hours. Which is suspiciously like a timezone.

```
$ for tz in America/New_York UTC Asia/Tokyo; do TZ=$tz dotnet edrdump strict/Open.xlsx | sed -n 5p; done
  TZ=America/New_York   0  Sheet1  0  4  D  2009-01-01 16:00:00.000
  TZ=UTC                0  Sheet1  0  4  D  2009-01-01 11:00:00.000
  TZ=Asia/Tokyo         0  Sheet1  0  4  D  2009-01-01 02:00:00.000
```

**ExcelDataReader's output depends on the machine's timezone.** The same bytes
decode to three different instants. For OOXML strict files, which store dates as
ISO 8601 text, it parses them and converts through the local zone.

This library returns 11:00 in all three, because chapter 1 decided that Excel
serials carry no timezone and attaching one makes a file decode differently on
different machines. That decision was made on principle, months of reasoning
compressed into a paragraph, and it is nice to see it hold up against a reference
implementation that went the other way.

Note the direction of the evidence: under `TZ=UTC`, ExcelDataReader agrees with
us exactly. We are not disagreeing about the value. We are disagreeing about
whether the answer should depend on the host.

The parity harness now pins `TZ=UTC` and says why.

## The real bugs

**Locale date formats.** `Issue541` carried `numFmtId="55"` and no format string
to inspect — id 55 is a built-in. My table covered ids 0–49.

Ids 27–36 and 50–58 are reserved for locale-specific formats and are *entirely*
dates: the Japanese, Chinese, Korean and Thai calendar layouts. ECMA-376 leaves
their format strings undefined, so a file using one supplies nothing to
pattern-match. A reader that only knows 0–49 silently returns a raw serial for
every date in a CJK-locale workbook. Three files.

**Impossible serials.** `Issue14_InvalidOADate` holds `1.0E12` with a genuine
date format. I converted it and produced the year 2,737,908,906. Excel's
calendar tops out at 9999; anything outside that range is a formatting accident,
not a date, and the number should be returned unchanged. Two files.

**Records that arrive out of order.** This was the big one, and I had the model
wrong.

My xls reader assumed cell records arrive grouped by row, in increasing order. It
emitted a row whenever the row number changed. On most files that is true.
`OldIssue11545_NoIndex.xls` walks like this:

```
read#0   RowIndex=0    cells=20  nonEmpty=0
read#1   RowIndex=1    cells=20  nonEmpty=0
...
read#75  RowIndex=0    cells=20  nonEmpty=20  "Group#"
```

Seventy-five blank rows, and then row index *restarts at zero* with the data.
The file emits MULBLANK records covering rows 0–74 and only afterwards the cells
that fill them. Every such sheet was reported twice: once empty, once populated.

Reading the reference implementation settled it — `XlsWorksheet` buffers a block
of rows and places cells by absolute index, precisely because BIFF makes no
ordering guarantee within a row block.

The fix is a bounded window: rows accumulate in a map keyed by absolute index,
and only rows strictly below the highest one seen are complete enough to emit.

```go
complete := r.maxBuffered
if r.sheetOpen {
        complete-- // the top row may still be receiving cells
}
```

That `complete--` is the whole correctness argument, and I got it wrong on the
first attempt — emitting the top row early truncated it, which a test caught at
row 1023, exactly where the buffer window ended.

The window caps memory at 1024 rows so a hostile file cannot make the reader
buffer a whole sheet. **Nine files.**

## Result

| | Before | After |
|---|---|---|
| Byte-identical | 214 | **229** |
| Both read it, output differs | 29 | **14** |
| Both refuse | 33 | 33 |
| Only ExcelDataReader reads | 27 | 27 |
| Only this library reads | 0 | 0 |

94% of files both readers accept. Sixteen fixed, one regressed — a single extra
empty cell in `Issue158.xls`, which I attempted to fix, broke a different file
with the guess, and reverted. It is still open, and one file is a better place
to stop than a second speculative change.

Of the remaining 14: six are elapsed-time cells where the reference returns
`TimeSpan`, three are the strict ISO timezone case above, one is whitespace
(the reference trims shared strings that carry `xml:space="preserve"`, which I
believe is wrong of it), and four are untriaged.

## Then: inputs nobody wrote

The corpus is 303 files that exist because a human hit a bug. Fuzzing is the
same idea with the human removed.

Three targets, one property: **for any input, reading either returns an error or
produces rows — never a panic, a hang, or an unbounded allocation.** Correctness
on garbage is not the goal. Refusing it safely is.

The important detail is seeding. Random bytes are rejected by magic-byte
detection in microseconds and exercise nothing. Seeding with the real fixtures
means mutations still reach the ZIP directory, the CFB allocation table and the
BIFF record loop, which is where the arithmetic lives.

**Twenty-five seconds in:**

```
runtime: out of memory: cannot allocate 17179869184-byte block
fatal error: out of memory

runtime.makeslice(...)
github.com/jloor/go-excel-reader/xls.(*cfb).readDIFAT(...)
	/home/bullwinkle/excel-go/xls/cfb.go:166
```

Seventeen gigabytes, from a two-kilobyte file.

```go
sectors := make([]uint32, 0, numFATSectors)
```

`numFATSectors` is a 32-bit field read straight out of the compound file header.
Set it to `0xFFFFFFFF` and the reader dutifully asks the allocator for
4,294,967,295 entries.

This is worse than a panic. A panic is recoverable; a Go fatal out-of-memory is
not. Any service accepting `.xls` uploads could be killed by a file small enough
to fit in a tweet.

The fix is to bound the claim by physical reality — a file cannot contain more
FAT sectors than it contains sectors:

```go
maxSectors := c.size/int64(c.sectorSize) + 1
if int64(numFATSectors) > maxSectors {
        return nil, fmt.Errorf("xls: header claims %d FAT sectors but the file holds at most %d",
                numFATSectors, maxSectors)
}
```

Auditing for the same shape turned up a second one: `columnIndex` accumulates a
cell reference's letters with no bound, and the row buffer is grown to reach
whatever column it names. An eleven-letter reference would ask for billions of
cells. Now bounded by the format's own limit of 16,384 columns.

After the fixes: **2.1 million executions across three targets, no further
failures.** The crasher is committed under `testdata/fuzz/` and runs as an
ordinary test forever.

## What the two techniques are each good for

They found completely different bugs, and neither would have found the other's.

The corpus found **semantic** errors — right structure, wrong meaning. A locale
date id I had never heard of. A serial out of range. Records in an order I did
not know was legal. Every one required a real file produced by real software
doing something reasonable, and no fuzzer would have generated a valid workbook
with `numFmtId="55"`.

The fuzzer found a **safety** error — a number used without asking whether it
could be true. No human would ever write a file claiming four billion FAT
sectors, so it will never appear in any corpus, and it is the one that gets your
service killed.

One asks "does it do the right thing with real input?" The other asks "can it be
made to do something catastrophic with input nobody would write?" Both took
about twenty minutes to set up. Neither is a substitute for the other.
