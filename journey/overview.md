I ported the core of a C# spreadsheet library to Go in a day. Every test passed. The
memory benchmark reported 3.6 MB against a 50 MB budget.

The component those tests existed to exercise had never once executed.

## What this is

[ExcelDataReader](https://github.com/ExcelDataReader/ExcelDataReader) is the C# library
most .NET codebases reach for when they have to read a spreadsheet. It handles `.xlsx`,
the legacy `.xls` binary format, and CSV, through a streaming cursor rather than by
loading a workbook into memory.

[go-excel-reader](https://github.com/jloor/go-excel-reader) is a port of its core to Go.
Pure Go, no CGo, one dependency. It reads all three formats through the same cursor model.

The finished thing works, and it is verified in a way I think is worth more than the code:
it and the original C# library both dump a file to a canonical form, and the dumps are
compared byte for byte — not on files I wrote, but on **ExcelDataReader's own 303-file
test corpus**, collected over a decade because each file broke something.

229 match exactly. 27 it refuses on purpose. 14 disagree and are written down. None are
read here that the original refuses.

That last sentence is the one worth having. The rest of this is how those numbers moved.

I am not a spreadsheet-format expert. I had not written a line of BIFF or CFB parsing
before this.

## The plan was good. The tests were the problem.

Yesterday's project was about a plan that was confidently wrong about the codebase it
described. This time the plan was detailed and largely right — it named the record types,
the magic bytes, the four RK encodings, and a gotchas table that was accurate.

So the failure moved somewhere else.

The shared string table is the reason an xlsx reader is hard to write efficiently. Almost
every string in a spreadsheet lives in one shared table and each cell stores an index into
it. It is the largest allocation a reader makes, and the whole memory design of this port
is built around it.

I wrote it, benchmarked it, and got 3.6 MB on a 4.6 MB file. Comfortably inside budget.

The fixtures had no shared string table in them. openpyxl, which generated them, writes
strings *inline* — `<c t="inlineStr">` — rather than into `xl/sharedStrings.xml`. There
was no such file in any fixture. The code had never run.

The tell was that the number was **too good**. A 200,000-entry table should cost several
megabytes; 3.6 MB was not consistent with having loaded one. Checking took one command.

With a real 400,000-entry table the honest figure is 26.9 MB. Still inside budget, but now
it means something.

## Then I did it again

The README said peak memory "tracks the widest row, not the file size."

That sentence describes what I designed, not what I measured. Rows do stream. But the
string table is held fully in memory, and on a string-heavy workbook it *is* the memory
profile — about 30 of the 39 MB. I shipped that claim in v0.1.0.

The fix was to stop asserting and measure. Two workbooks, identical row and cell counts,
differing only in how many *distinct* strings they hold:

| Workbook | Cells | Distinct strings | Peak heap |
|---|---|---|---|
| `repeat.xlsx` | 1,000,000 | 100 | **3.5 MB** |
| `large.xlsx` | 1,000,000 | 400,000 | **26.9 MB** |

That is a better result than the claim I couldn't defend, and unlike the claim it can be
falsified. A million cells in 3.5 MB shows rows are not accumulating. The gap between the
two numbers is the string table and nothing else. Row count appears in neither.

Both are now tests. The baseline one matters more: if row streaming ever regressed into
accumulation, the string-heavy fixture would still look plausible while the repeating one
blew past its budget.

## Then I stopped grading my own work

At this point everything passed and I had checked my twelve fixtures against the
original C# library, byte for byte. That felt like strong evidence. It was
evidence about twelve files I chose.

ExcelDataReader ships **379 test files** of its own — real spreadsheets its
maintainers collected over a decade because each one broke something. Names like
`DateFormatButNotDate.xls`, `ClipboardBiff8.xls`, `Issue467_EmptyContinueLeftoverbytes.xls`.

So I pointed the same harness at all 303 of the `.xls` and `.xlsx` files:

| Result | Before fixes | After |
|---|---|---|
| Byte-identical to ExcelDataReader | 207 | **229** |
| Both readers refuse (xlsb, encrypted, corrupt) | 33 | 33 |
| Both read it, output differs | 37 | **14** |
| Only ExcelDataReader reads it | 26 | 27 |
| **Only mine reads it** | **0** | **0** |

That last row is the one I care about most. Nothing here parses a file the
reference implementation refuses, which would mean inventing data.

It found three defects, each a different kind:

**A silent one.** `Issue411.xls` returned zero rows, no error, exit code 0. A
caller gets an empty spreadsheet and no indication anything went wrong.

The cause is the best thing in this project. A compound file is a *tree* of
storages, and a `.xls` containing an embedded object has nested storages with
streams of their own. I was finding streams by scanning all directory entries
linearly, so I resolved a `Workbook` belonging to the embedded object rather
than the real one.

I had deleted the sibling and child pointers as dead fields, and written a
comment explaining why that was fine:

> *"walking the tree would be asymptotically better on a directory far larger
> than any spreadsheet has"*

Confidently wrong. Those pointers aren't a performance detail — they are what
gives an entry its **scope**. I removed them, rationalised it in a comment, and
the rationalisation read as informed. Fixed, that file now says plainly that it
is BIFF5 and unsupported. Silent data loss became an honest error.

**A subtle one.** ExcelDataReader distinguishes an empty *inline* string from an
empty entry in the *shared* string table — the first is an absent value, the
second is a value that happens to be empty. I treated both as text.

My first fix applied the rule everywhere. It repaired four files and broke
eight. The corpus caught that immediately, which is the entire point: a
one-line change that looks obviously right, measured against 303 files, is a
question with an answer instead of an opinion.

**A deliberate one.** Error cells. EDR reports `#DIV/0!` as having no value; I
returned the text. Mine is arguably more useful, but "arguably more useful" is
not "identical," and my README said identical. It is now an option, off by
default, and there is a table of known intentional differences.

## The failure mode, again

Three times now, something reported success while measuring less than I thought.
Same shape as the four that did it yesterday:

**A passing test proves the assertion held. It does not prove the code ran.**

Yesterday's version was a check that could not distinguish "the thing is fine" from "I
failed to look." This one could not distinguish "the code works" from "the code was never
called." A test whose fixture doesn't reach the code under test passes for free, and it
passes *faster*, and the speed reads as efficiency.

The third instance was the twelve fixtures themselves. They agreed with the reference
implementation perfectly — and they were twelve files I chose, which is a sample that
cannot surprise me. The corpus could, and did, eight times.

What helps is the same thing that helped yesterday: compare against something that can
actually differ. So the fixtures are now written by three independent tools, two of which
store strings differently on purpose, and a test asserts that property directly —

```go
// If a regenerated fixture quietly stopped using the shared string table, the
// string-table reader would go untested while every test still passed.
for _, f := range []string{"sst.xlsx", "large.xlsx"} { ... }
if hasPart("simple.xlsx", "xl/sharedStrings.xml") { t.Error(...) }
```

The same reasoning produced a negative control in the BIFF string test: it parses the same
bytes *without* the continuation boundaries and asserts the result comes out **wrong**.
Without it, an assertion that `"abcd" == "abcd"` can pass for reasons that have nothing to
do with the code being right.

## What the port actually cost

The honest scorecard, because a post that only lists wins isn't worth reading:

| | |
|---|---|
| Wrote it and it worked first run | `numfmt`, the xlsx reader, the CSV reader, the BIFF8 reader |
| Cost real time | Fixtures that tested nothing; a memory claim I couldn't defend |
| Found only by someone else's tests | A silent zero-row read; an empty-string rule; a parity claim that was narrower than it sounded |
| Deliberately not ported | xlsb, encryption, BIFF5 and earlier, `AsDataSet()` |
| Deliberately not optimised | `unsafe.String` on the string blob — meets budget without it |

Four subsystems landed green on the first test run. That is not a boast; it is the setup
for the point. **The two things that went wrong were both in the layer that was supposed to
tell me whether the code was right.** Being good at the hard part bought me nothing,
because the part that failed was the verification.

## The two details that were genuinely hard

**A date is not a type.** A spreadsheet stores a date as a bare `float64`. Nothing in the
cell says "date" — the only signal is a number format reached through the style table,
three hops away. Break any hop and a delivery date becomes the number 45217. Worse,
scanning the format string for `d`/`m`/`y` is not enough: a format like `"Paid on day "0.0`
contains all of them inside a quoted literal. That one is a fixture now.

**BIFF8's restated flags byte.** A record payload can't exceed 8224 bytes, so a large
string table arrives as one record plus a run of `CONTINUE` records. Concatenate them and
almost everything works. But when a *string* is split across a boundary, the first byte of
the continuation is not text — it restates the compression flag. Miss it and the table
doesn't fail; it silently desynchronises, and every string after that point comes back
garbled. The fixture generator writes random-length strings specifically to force the case.

## Where it beats the original, and where the comparison is unfair

On the unfavourable input — 400,000 distinct strings, the worst case for this design:

| Reader | Time | Peak RSS |
|---|---|---|
| **go-excel-reader** | **3.2 s** | **40 MB** |
| ExcelDataReader 3.9.0 (.NET 9) | 3.3 s | 95 MB |
| excelize v2.11.0 | 4.0 s | 221 MB |

The excelize row is not like for like on time: its cursor returns strings and does less
type resolution than the other two. Memory is the meaningful column, and the README says
so rather than letting the time column imply something it shouldn't.

The mechanism is one idea: the string table is stored as a single byte blob with an offset
index instead of a `[]string`. For 400,000 entries that is two allocations and zero
pointers instead of 400,000 of each — and the garbage collector never traces into a
`[]byte`.

## Deliberately not claimed

- **229 of 303 is not 303.** 14 files are read by both and disagree. Ten are diagnosed
  (elapsed-time cells, strict ISO timezones, whitespace) and four are open. The count is
  pinned by a test, but "94% agreement" is the honest headline, not "matches the original".
- **The comparison pins the reference to `TZ=UTC`.** ExcelDataReader's output for OOXML
  strict dates depends on the host's timezone; ours does not. Without that pin, the
  agreement figure would depend on where it was measured.
- The 27 files ExcelDataReader reads and this does not are mostly BIFF5-and-earlier and
  encrypted workbooks, both refused on purpose. A handful are not, and are open.
- The "streams a 500 MB workbook" claim is a design property, not a measurement. The
  largest file tested is 9.2 MB.
- `golangci-lint` was not run; it wasn't installed. `go vet` and `staticcheck` are clean.
- BIFF5 and earlier `.xls` are refused with a clear error rather than decoded. That is a
  scope decision, not a capability.
- One machine, one architecture, one Go version.

## And then inputs nobody wrote

The corpus is 303 files that exist because a human hit a bug. Fuzzing is the same idea
with the human removed, and it found a different *kind* of bug within twenty-five
seconds:

```
runtime: out of memory: cannot allocate 17179869184-byte block
	xls.(*cfb).readDIFAT  cfb.go:166
```

A compound file header declares how many sectors its allocation table occupies. I used
that number as an allocation size. A two-kilobyte file claiming `0xFFFFFFFF` sectors
asks for seventeen gigabytes and kills the process — and a Go fatal out-of-memory is not
recoverable, so no caller can defend against it.

The two techniques found completely different things and neither would have found the
other's. The corpus found **semantic** errors: a locale date format id I had never heard
of, a serial out of range, records in an order I did not know was legal. Every one needed
a real file produced by real software. The fuzzer found a **safety** error: a number used
without asking whether it could be true. No human would ever write a file claiming four
billion FAT sectors, so it will never appear in any corpus — and it is the one that gets
your service killed.

2.1 million executions after the fix, no further failures.

## Read the actual work

The chapters below are the engineering log in order, including both mistakes and the
corrections. The decision log records what was ambiguous and what got traded away.
