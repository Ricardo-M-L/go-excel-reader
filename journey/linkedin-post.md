# LinkedIn post

Draft C is the current preference. It leads with intent rather than findings.
A and B are more technical framings. None are posted.

Links: writeup https://www.jonathanloor.com/go-excel-reader/ ·
repo https://github.com/jloor/go-excel-reader ·
pkg https://pkg.go.dev/github.com/jloor/go-excel-reader

**This is a separate post from the Equinox/JsxCore one.** Different project,
different lesson, and running them together would bury both. If they go out in
the same week, lead with this one: the denial-of-service finding is more concrete
than anything in the other post, and the shared theme reads better as a pattern
across two projects than as one long post.

---

## Draft C: leads with the why (current preference)

I spent a day porting a C# library to Go. The most useful twenty minutes came at
the end, when I stopped grading my own work.

By evening it was published. Every test passed, static analysis was clean, and I
had compared my output byte-for-byte against the original C# library on twelve
test files. All twelve matched.

So I wrote in the README that the output was identical to the original.

It was identical on twelve files I chose myself. That is a much weaker claim.

The original ships 379 test spreadsheets of its own, a decade of regression
files, each added because it broke something for somebody. Running against all
303 of them took twenty minutes and found three defects.

The one that mattered returned zero rows, no error, exit code zero. Not a crash,
just a silent and confident empty answer. A .xls file can contain another
spreadsheet embedded inside it, and I was reading the wrong one.

I knew about that. Hours earlier I had deleted the fields that distinguish them,
as tidying up, and written a comment explaining they were only a performance
detail.

That comment was wrong, specific, and read like it had been thought about. I
believed it because I had written it down. My own tests could never have caught
it, because the same person wrote the tests and the bug.

Then I let a fuzzer write the inputs instead. Twenty-five seconds later a
two-kilobyte file asked my library for a seventeen-gigabyte allocation and killed
the process.

The two approaches found completely different things. The corpus found errors of
meaning, and every one of them needed a real file made by real software. The
fuzzer found an error of safety, a number I had trusted without asking whether it
could be true. Neither would have found the other's.

229 of the 303 files now decode byte-identically, 14 disagree and are written
down, and zero parse in mine that the original refuses.

Written up with the corrections, because the version where everything goes
smoothly is not useful to anyone.

Writing is AI assisted. Thoughts and publishing are human-gated.

#golang #softwareengineering #testing

---

## Draft A: technical framing

I ported the core of ExcelDataReader from C# to Go in a day. Every test passed.
The memory benchmark reported 3.6 MB against a 50 MB budget.

The component those tests existed to exercise had never executed once.

Almost every string in an .xlsx lives in a shared table that cells index into,
and the entire memory design turns on it. My benchmark said 3.6 MB.

My fixtures did not contain a shared string table. The tool that generated them
writes strings inline instead. There was no such file in any fixture.

The tell was that the number was too good. A 200,000-entry table should cost
several megabytes. Checking took one command.

Four things I would want to read in someone else's post about this:

1. A passing test proves the assertion held. It does not prove the code ran.

A test whose fixture never reaches the code under test passes for free, and it
passes faster, which in a memory test reads as efficiency. Fixtures are now
written by three independent tools, two of which store strings differently on
purpose, and a test asserts that property so it cannot silently regress.

2. I shipped a claim I could not defend.

The README said peak memory "tracks the widest row, not the file size." That
described what I designed, not what I measured. Corrected in a point release. The
replacement is two workbooks with identical row and cell counts differing only in
distinct strings: 3.5 MB against 26.9 MB. Same rows. That one is falsifiable.

3. Twelve fixtures agreeing with the reference implementation is not validation.

The original ships 379 test spreadsheets accumulated over a decade of bug
reports. Running against 303 of them found three defects, including one that
returned zero rows with no error: silent data loss, caused by picking a workbook
stream belonging to an object embedded inside the file. I had deleted the
directory-tree pointers that prevent that, as cleanup, with a comment explaining
they were merely a performance detail. They are not. They are what gives a
stream its scope.

After triage: 229 of 303 byte-identical, 14 disagreements documented, 27 files
refused on purpose (BIFF5 and earlier, encrypted), zero read that the original
rejects.

4. Fuzzing found a different class of bug entirely.

Twenty-five seconds in: a 2 KB file requesting a 17 GB allocation, because the
compound-file header's sector count was used directly as an allocation size. A Go
fatal out-of-memory is unrecoverable, so a caller cannot defend against it. 2.1
million executions clean after the fix, and the crashing input is committed as a
regression test.

Reads a million-cell workbook in ~3s / 40 MB, against 95 MB for the C# original
and 221 MB for excelize. Pure Go, no CGo.

Writeup, mistakes included: https://www.jonathanloor.com/go-excel-reader/

ExcelDataReader is MIT and well-built. That it ships a decade of adversarial test
files is why this port is any good.

Writing is AI assisted. Thoughts and publishing are human-gated.

#golang #softwareengineering #testing

---

## Draft B: shorter, if A feels long for the feed

I ported a C# Excel-reading library to Go. Everything passed: tests, vet, static
analysis, and a byte-for-byte comparison against the original on twelve files.

Then I ran it against the original's own test corpus, 379 spreadsheets collected
over a decade because each one broke something. Three bugs in twenty minutes.

The worst returned zero rows, no error, exit code zero. A .xls can have another
spreadsheet embedded inside it and I was reading the wrong one. I had deleted the
fields that distinguish them hours earlier as cleanup, with a comment explaining
they were only a performance optimisation.

The comment was wrong and read like it had been thought about. I believed it
because I had written it. My own tests could never have caught that, because the
same person wrote the tests and the bug.

Then I let a fuzzer write the inputs instead, and twenty-five seconds later a
two-kilobyte file asked my library for a seventeen-gigabyte allocation and killed
the process.

The corpus found errors of meaning. The fuzzer found an error of safety. Neither
would have found the other's, and both took twenty minutes.

229 of 303 files now decode byte-identically, 14 disagree and are documented,
zero parse here that the original refuses. A million cells in ~3s / 40 MB,
against 95 MB for the C# original.

Written up with the corrections: https://www.jonathanloor.com/go-excel-reader/

Writing is AI assisted. Thoughts and publishing are human-gated.

#golang #softwareengineering #testing

---

## Follow-up comment: links and thanks

Post as the first comment. Links in the body reportedly get less reach than links
in a comment, so the drafts keep none and everything lives here. LinkedIn
comments do not render markdown, so this is written as plain text.

```
Links, for anyone who wants to dig in:

The writeup, nine chapters in order, including the parts I got wrong
https://www.jonathanloor.com/go-excel-reader/

The chapter on running against someone else's test corpus
https://www.jonathanloor.com/go-excel-reader/chapters/07-the-corpus/

The chapter on triage and fuzzing, including the denial of service
https://www.jonathanloor.com/go-excel-reader/chapters/08-triage-and-fuzzing/

The decision log: what was ambiguous, what got chosen, what got traded away
https://www.jonathanloor.com/go-excel-reader/decisions/

Source, including the parity harness and the three fuzz targets
https://github.com/jloor/go-excel-reader

Thanks are owed in a specific order.

To Johan Appelgren and the ExcelDataReader maintainers. This is a port of their
library, and more to the point, their test corpus is what found my bugs. Ten
years of collecting the files that broke a reader, kept in the repository where
anyone can point a harness at them. That is a real gift to anyone reimplementing
this, and none of the numbers in this post would exist without it. MIT licensed,
and a genuinely good codebase to read.
https://github.com/ExcelDataReader/ExcelDataReader

To the people who maintain openpyxl, xlsxwriter and xlwt. I generated test
fixtures with all three rather than writing my own, specifically so the tests
would exercise the format instead of agreeing with a writer I controlled. That
turned out to matter more than I expected: two of them store strings differently,
and the difference exposed a component of mine that had never once executed.

To the Go team for a fuzzer that ships in the standard library. It found a
denial of service twenty-five seconds after I pointed it at the parser, and it
took about twenty minutes to set up. If you maintain something that parses
untrusted bytes and you have not done this, it is a very cheap afternoon.
```

Notes on this comment:

- Order is deliberate: what I made, then whose work it stands on. The thanks are
  specific about what each project actually did for the work, because a generic
  "thanks to the maintainers" reads as decoration.
- Johan Appelgren is `appel1` on GitHub and has 297 commits to ExcelDataReader,
  which I confirmed from the commit history. `andersnm` is the other major
  contributor at 174. I have **not** verified either person's LinkedIn profile.
  Thank them by name in the comment text; only @-tag someone once you have
  confirmed the profile is really theirs.
- excelize appears in a benchmark row in Draft A and is deliberately not thanked
  or tagged. It is a comparison, not a dependency, and the post shows it using
  five times the memory on a benchmark I designed.
- There is no upstream issue to link this time. The Equinox post had four. Do not
  manufacture one to match that shape; if the timezone finding gets filed, it can
  be added here.

## Graphic

`journey/assets/corpus-parity.png`, 1200x1500, sized for the LinkedIn feed.

A unit chart: 303 squares, one per file in ExcelDataReader's test corpus,
coloured by how our output compared with theirs. The unit is meaningful, since
each square is a real spreadsheet that broke a reader once, and squares read as
spreadsheet cells, which suits the subject. A bar reading "94%" would hide both.

Cells are sorted so only consecutive categories touch, which keeps the palette on
its validated adjacent pairs. The palette passes all six checks against the dark
surface: lightness band, chroma floor, colour-vision separation, normal-vision
floor, and contrast. Every category carries a count and a text label, so identity
never depends on colour alone.

Regenerate after the numbers change: the counts are read from
`testdata/corpus_baseline.tsv` and must be updated in the generator if that file
moves. The figures in the image are the same ones the post quotes, so they go
stale together.


## Who to tag

**ExcelDataReader.** I have not verified who currently maintains it. The
repository sits under an organisation rather than a personal account, which
usually means several people. Check the recent commit history and the org members
page, and confirm any LinkedIn profile actually belongs to that person before
tagging. Tagging the wrong individual is worse than not tagging.

There is no strong reason to tag here. The post says their test corpus found my
bugs, which is unambiguously complimentary and needs no warning.

**excelize** appears in one comparison row (github.com/xuri/excelize). I would not
tag. The post shows it using five times the memory on a benchmark I designed, and
the caveat that its cursor does less type resolution is in the writeup but not in
the post. A tag turns a footnote into a callout.

## Notes

- LinkedIn's hard limit is 3000 characters and it folds at roughly 200. All
  three drafts put the hook above the fold and fit the limit; current lengths
  are noted at the top of each draft. An earlier version of Draft C ran to 4016
  characters and would simply have been rejected.
- LinkedIn does not render markdown. No bold, no bullets, so the numbered
  sections in Draft A rely on line breaks and short paragraphs.
- **File the timezone finding upstream before posting about it.** It is
  reproducible in three commands and makes their output host-dependent. Saying it
  publicly before telling them reads badly; saying it after reads as a
  contribution.
- Deliberately not claimed: that this is production-ready (v0.3.0, no users),
  that it is faster than the C# original (roughly level on time, memory is the
  difference), or that 229/303 is parity. It is not, and the post gives the real
  number.
- If someone asks what the 14 disagreements are: six are elapsed-time cells where
  the reference returns a TimeSpan, three are the timezone case, one is
  whitespace, and four are untriaged. Do not claim more than that.
- The previous version of this file claimed 214 byte-identical and 29
  disagreements, and did not mention fuzzing at all. Those numbers were correct
  when written and wrong by the time the work finished. Re-check every figure
  against `testdata/corpus_baseline.tsv` before posting.
