# Chapter 7: Someone else's test files

*2026-08-06. The library was tagged, published, and green. This is what happened after.*

Everything passed. `go vet`, `staticcheck`, twelve fixtures written by three
independent tools, and a byte-for-byte comparison against the real C#
ExcelDataReader that matched on every one.

I wrote in the README that the output was identical to the original.

It was identical on twelve files I chose. That is not the same claim, and the
difference is the whole chapter.

## What was actually available

ExcelDataReader's repository ships 379 spreadsheets in `src/TestData`. They are
not demos. They are a decade of regression fixtures, accumulated one bug report
at a time, and the filenames give it away:

```
DateFormatButNotDate.xls          Issue467_EmptyContinueLeftoverbytes.xls
ClipboardBiff8.xls                Issue477_SstZeroCount.xls
NoStylesNoRAttribute.xlsx         InvalidByteOrderValueInHeader.xls
BoolFormula.xls                   StringContinuationAfterCharacterData.xls
```

Every one of those names is a bug someone hit in production. `SstZeroCount` is a
shared string table claiming it holds zero strings. `EmptyContinueLeftoverbytes`
is a `CONTINUE` record with trailing bytes nobody expected. I would never have
thought to construct these inputs, because thinking of them requires having
already been broken by them.

The harness already existed. Pointing it at 303 `.xls` and `.xlsx` files was
about fifteen lines of shell.

## The first run

| Result | Files |
|---|---|
| Byte-identical | 207 |
| Both read it, output differs | 37 |
| Both refuse it | 33 |
| Only ExcelDataReader reads it | 26 |
| Only this library reads it | **0** |

The zero is the number I checked first. A file that parses here but not in the
reference implementation would mean this library is inventing data out of bytes
the original correctly declines to interpret. There were none, before or after.

Most of the 26 gaps were expected. `as3xls_BIFF2.xls`, `biff3.xls` and
`Issue224_Simple95.xls` are all pre-BIFF8 files this library refuses on purpose,
and `protectedsheet-xxx.xls` is encrypted. About eleven were not explained by
either, and remain open.

The 37 disagreements were the interesting part.

## Defect one: the silent read

`Issue411.xls` produced zero rows, no error, exit code 0. ExcelDataReader read
305 cells of Cyrillic text out of it.

Zero rows and no error is the worst possible failure, because every other kind
announces itself. This one hands the caller an empty spreadsheet and lets them
conclude the file was empty.

Dumping the directory of the compound file explained it:

```
dir entry "Root Entry"    type=5  size=704     start=1052
dir entry "Book"          type=2  size=530440  start=14
dir entry "MBD00002b54"   type=1  size=0       start=0
dir entry "Workbook"      type=2  size=229674  start=2032
dir entry "\x02OlePres000" type=2 size=495348  start=1060
```

Two workbook streams. `MBD00002b54` is a *storage*, meaning a directory, and
`OlePres000` is an OLE presentation cache. This file has a spreadsheet embedded
inside it.

My stream lookup scanned every directory entry in the file and took the first
name match, so it found the `Workbook` belonging to the embedded object rather
than the `Book` at the root. That stream had three sheets, a legitimate BIFF8
header, an 8-byte string table, and no cell records at all. Everything about it
looked fine. It was simply the wrong workbook.

### The comment I had written about exactly this

The fix is to walk the directory as the tree it is, which I could not do, because
I had deleted the pointers that make it a tree:

```go
// The directory is stored as a red-black tree, but the entries also sit
// contiguously in their sectors, so finding a stream by name is a linear scan
// over all of them. That is why the sibling and child pointers are not kept:
// walking the tree would be asymptotically better on a directory far larger
// than any spreadsheet has.
```

I wrote that. It is confidently, specifically wrong, and it reads like it was
written by someone who had thought about the problem.

The sibling and child pointers have nothing to do with performance. They are what
gives a directory entry its scope, meaning which storage it belongs to.
Discarding them does not cost you speed on large directories. It costs you the
ability to tell a root-level stream from one buried inside an embedded object.

Nothing forced this, either. `staticcheck` flagged two unused *functions*, not
these fields. I removed them as tidiness, then wrote a justification, and the
justification was plausible enough that I did not revisit it.

The traversal is now nine lines. Root-level entries are enumerated by following
sibling links from the root storage's child, deliberately *not* descending
through `child` pointers, since that is what would walk back into an embedded
object:

```go
walk = func(id uint32) {
        if id > maxRegSect || int(id) >= len(c.entries) || seen[id] { return }
        seen[id] = true
        e := c.entries[id]
        walk(e.leftSib)
        out = append(out, int(id))
        walk(e.rightSib)
}
walk(c.entries[c.rootIdx].child)
```

`Issue411.xls` now resolves to the root-level `Book` stream, reads its BOF, finds
version `0x0500`, and reports that BIFF5 is unsupported. That is still a file
this library cannot read, but it now says so, and turning silent data loss into
an error message is the entire improvement.

## Defect two: the fix that broke eight things

`NoStylesNoRAttribute.xlsx` disagreed by 546 cells. Its sheet is full of this:

```xml
<c t='inlineStr'><is><t></t></is></c>
```

Inline strings with no text. ExcelDataReader reports them as absent; I reported
them as the empty string.

So I made empty text produce an empty cell. Three lines, obviously correct.

It fixed four files and broke eight.

The corpus said so within a minute, which is the point. Without it I would have
shipped a change that looked right, verified it against fixtures that did not
contain the case, and been wrong in the other direction for a year.

The reason is a distinction I would not have guessed:

| Where the empty text comes from | ExcelDataReader |
|---|---|
| `<c t="inlineStr"><is><t></t></is></c>` | absent |
| An empty entry in the shared string table | the empty string |
| An empty `LABEL` record in a `.xls` | the empty string |

An empty *inline* string is a cell that was written and holds nothing. An empty
*shared* string is a value the file went to the trouble of interning. Those are
different things, and the reference implementation treats them differently.

I do not know whether that distinction is deliberate upstream or an artefact of
two code paths. Either way it is the behaviour, and matching it is what parity
means. The narrow rule fixed four files and broke none.

## Defect three: the one that was my opinion

`Issue329_Error.xls` and its `.xlsx` twin: ExcelDataReader emits nothing for a
cell holding `#DIV/0!`. I emitted the text.

I still think returning the text is more useful, because an error is information
and silently reporting it as an empty cell loses it. I had even written a comment
saying so.

But "more useful than the reference" is not "identical to the reference", and my
README claimed the second one. The claim was the problem, not the behaviour.

It is now `WithErrorValues(bool)`, defaulting to the reference behaviour, with
the text always preserved in `Cell.Raw` either way. There is also a table in the
README of known intentional differences, which is what should have existed the
moment I decided to differ.

## After

| Result | Before | After |
|---|---|---|
| Byte-identical | 207 | **214** |
| Both read it, output differs | 37 | **29** |
| Both refuse it | 33 | 33 |
| Only ExcelDataReader reads it | 26 | 27 |
| Only this library reads it | 0 | 0 |

Every file that changed verdict moved in the right direction. `Issue411.xls`
moved from "differs" to "only EDR reads it", which is a demotion on paper and a
promotion in fact: it went from lying to declining.

## Making the number a test

Numbers in a README go stale. The corpus is now pinned by
`testdata/corpus_baseline.tsv`, one row per file, holding the verdict from the
comparison and a hash of *our* dump.

Hashing our own output is what lets the check run in CI without .NET installed.
The hashes were derived from a run where the comparison actually happened, so
pinning them pins the parity result by proxy.

`TestCorpusBaselineShape` asserts the four counts, needs no corpus, and always
runs. `TestCorpusParity` re-derives all 303 dumps when `EXCEL_CORPUS` points at a
checkout.

The first version of the baseline generator was a shell script. It hashed the
empty output of a rejected file, while the Go test wrote the literal string
`REJECTED`. Seventeen files reported as changed when nothing had changed.

Which is the same mistake as everything else here. The generator and the checker
were two implementations of one idea, and they disagreed. Regeneration now runs
through the identical function the check uses, behind an environment variable.

## What I would tell myself before starting

Twelve fixtures agreeing with the reference implementation felt like strong
evidence. It wasn't. They were twelve files I chose, and a sample you choose
cannot surprise you. The corpus surprised me eight times in the first run.

If the thing you are building has an established implementation, that
implementation almost certainly ships its accumulated scar tissue as test data.
It is the cheapest high-quality input you will ever get, and running against it
took twenty minutes.
