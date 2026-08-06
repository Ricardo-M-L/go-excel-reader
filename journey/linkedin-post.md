# LinkedIn post

Draft C is the current preference. It leads with intent rather than findings.
A and B are earlier, more technical framings. None are posted.

Links: repo https://github.com/jloor/go-excel-reader ·
writeup, see note below ·
pkg https://pkg.go.dev/github.com/jloor/go-excel-reader

**Note on the writeup link.** Unlike the Equinox project there is no deployed
site for this one yet. The `journey/` folder renders as markdown on GitHub, so
https://github.com/jloor/go-excel-reader/tree/main/journey works today. If you
want a hosted version matching the last one, that is a separate job.

---

## Draft C: leads with the why (current preference)

I spent a day porting a C# library to Go, and the most useful thing I did took
twenty minutes at the end.

The library reads Excel files: the modern .xlsx format, the old binary .xls one,
and CSV. I'd never written a line of binary spreadsheet parsing before. By
the evening it was published, every test passed, static analysis was clean, and
I had checked my output byte-for-byte against the original C# library on twelve
test files. All twelve matched.

I wrote in the README that the output was identical to the original.

It was identical on twelve files I chose. Which is a different sentence.

The original library ships 379 test spreadsheets of its own, a decade of
regression files, each one added because it broke something for somebody. So I
ran the same comparison against all 303 of them.

It found three bugs in twenty minutes.

One of them mattered a lot. A particular file returned zero rows, no error, exit
code zero. Not a crash, just a silent and confident empty answer. The cause was that a
.xls can contain another spreadsheet embedded inside it, and I was picking the
wrong one. My code found streams by scanning a flat list, ignoring the tree
structure that says which container each one belongs to.

I knew about that tree. I had deleted the fields that represent it a few hours
earlier, as tidying up, and written a comment explaining that they were only a
performance detail and the flat scan was fine.

That comment was wrong, specific, and read like it had been thought about.

I think that's the part worth sharing. Not the bug, but the fact that I generated a
confident-sounding justification for removing something I didn't understand, and
then believed it because I'd written it down. Nothing in my tests could catch
that, because my tests were also written by me.

The fix took nine lines. Finding it took someone else's test files.

If the thing you're building has an established implementation, it almost
certainly ships its accumulated scar tissue as test data. It's the cheapest
adversarial input you will ever get, and it doesn't care what you believe about
your own code.

Numbers, since people ask: it reads a million-cell spreadsheet in about 3
seconds using 40 MB, against 95 MB for the C# original and 221 MB for the most
popular Go alternative. 214 of the 303 corpus files now decode byte-identically,
29 disagree and are written down, and zero parse here that the original refuses,
which is the number I'd actually worry about.

Everything is written up, corrections included, because the version where it all
goes smoothly isn't useful to anyone.

Writing is AI assisted. Thoughts and publishing are human-gated.

#golang #softwareengineering #testing

---

## Draft A: technical framing (earlier version)

I ported the core of ExcelDataReader from C# to Go in a day. Every test passed.
The memory benchmark reported 3.6 MB against a 50 MB budget.

The component those tests existed to exercise had never executed once.

Almost every string in an .xlsx lives in a shared table that cells index into.
It's the largest allocation a reader makes and the entire memory design turns on
it. My benchmark said 3.6 MB.

My test fixtures didn't contain a shared string table. The tool that generated
them writes strings inline instead. There was no such file in any fixture.

The tell was that the number was too good. A 200,000-entry table should cost
several megabytes. Checking took one command.

Three things I'd want to read in someone else's post about this:

1. A passing test proves the assertion held. It doesn't prove the code ran.

A test whose fixture never reaches the code under test passes for free, and it
passes faster, and the speed reads as efficiency. Fixtures are now written by
three independent tools, two of which store strings differently on purpose, and
a test asserts that property directly so it can't silently regress.

2. I shipped a claim I couldn't defend.

The README said peak memory "tracks the widest row, not the file size." That
described what I designed, not what I measured. The string table is fully
resident, and on a string-heavy file it IS the memory profile. Corrected in a
point release. The replacement is two workbooks with identical row and cell
counts differing only in distinct strings: 3.5 MB vs 26.9 MB. Same rows. That
one's falsifiable.

3. Twelve fixtures agreeing with the reference implementation is not validation.

The original ships 379 test spreadsheets accumulated over a decade of bug
reports. Running against 303 of them found three defects, including one file
that returned zero rows with no error. Silent data loss, caused by my picking a
workbook stream belonging to an object embedded inside the file.

I'd deleted the directory-tree pointers that would have prevented it, as
cleanup, with a comment explaining they were merely a performance detail. They
aren't. They're what gives a stream its scope.

After fixes: 214 of 303 byte-identical, 29 disagreements written down, 27 files
it refuses on purpose (BIFF5 and earlier, encrypted), and zero files it reads
that the original rejects.

Reads a million-cell workbook in ~3s / 40 MB, vs 95 MB for the C# original and
221 MB for excelize. Pure Go, no CGo, one dependency.

Repo and full writeup, mistakes included:
https://github.com/jloor/go-excel-reader

ExcelDataReader is MIT and genuinely well-built. The fact that it ships a decade
of adversarial test files is why this port is any good.

Writing is AI assisted. Thoughts and publishing are human-gated.

#golang #softwareengineering #testing

---

## Draft B: shorter, if A feels long for the feed

I ported a C# Excel-reading library to Go. Everything passed: tests, vet,
static analysis, and a byte-for-byte comparison against the original on twelve
files.

Then I ran it against the original's own test corpus: 379 spreadsheets collected
over a decade because each one broke something.

Three bugs in twenty minutes.

The worst returned zero rows, no error, exit code zero. A .xls can have another
spreadsheet embedded inside it, and I was reading the wrong one. I'd deleted the
fields that distinguish them a few hours earlier as cleanup, with a comment
explaining they were only a performance optimisation.

The comment was wrong and read like it had been thought about. I believed it
because I'd written it.

My own tests could never have caught that. They were written by the same person
who wrote the bug.

214 of 303 files now decode byte-identically, 29 disagree and are documented,
and zero parse here that the original refuses. A million cells in ~3s / 40 MB,
against 95 MB for the C# original.

Written up with the corrections:
https://github.com/jloor/go-excel-reader

Writing is AI assisted. Thoughts and publishing are human-gated.

#golang #softwareengineering #testing

---

## Follow-up comment: all the links

Post as the first comment. Links in the body reportedly get less reach than
links in a comment, so Draft C keeps one and everything else lives here.

```
Links, for anyone who wants to dig in:

The library, a streaming reader for .xlsx, .xls and .csv, pure Go, no CGo
https://github.com/jloor/go-excel-reader

The writeup, an engineering log in order, including what I got wrong
https://github.com/jloor/go-excel-reader/tree/main/journey

The decision log: what was ambiguous, what got chosen, what got traded
https://github.com/jloor/go-excel-reader/blob/main/journey/decisions.md

Docs
https://pkg.go.dev/github.com/jloor/go-excel-reader

The project this is a port of, and the source of the test corpus that found
my bugs. MIT, and a genuinely good codebase:
https://github.com/ExcelDataReader/ExcelDataReader
```

Notes on this comment:
- Order is deliberate: what I made, then how it went, then whose work it stands
  on. The corpus credit belongs at the end and it belongs unqualified.
- There are no upstream issues to file this time. That's worth noticing rather
  than manufacturing: the last project surfaced four, this one surfaced none,
  and inventing a finding to match the previous post's shape would be dishonest.
  If the 29 remaining disagreements get diagnosed and any turn out to be
  upstream bugs, that's a follow-up post, not this one.

---

## Who to tag

**ExcelDataReader**, the library this ports.

I have not verified who currently maintains it. The repository is
github.com/ExcelDataReader/ExcelDataReader, MIT-licensed, under an organisation
rather than a personal account, which usually means several maintainers.

Before tagging anyone: check the recent commit history and the org members page,
and confirm the LinkedIn profile actually belongs to that person. Tagging the
wrong individual is worse than not tagging.

Unlike the Equinox post, there is no strong reason to tag here. That post
described real problems in someone's project and tagging was partly courtesy,
giving them a heads-up before their community saw it. This post says their test
corpus found my bugs, which is unambiguously complimentary and needs no warning.
A tag would be nice for them; its absence costs nothing.

**excelize**, which appears in one comparison row (github.com/xuri/excelize).

I would not tag. The post shows it using 5x the memory on one benchmark that
isn't quite like-for-like, since its cursor returns strings and does less type
resolution. That caveat is in the README and in the post, but a tag turns a
footnote into a callout, and it reads as picking a fight over a benchmark I
designed.

## Notes

- LinkedIn truncates around 200 characters. All three drafts put the hook above
  that line.
- LinkedIn does not render markdown. No bold, no bullets, so the numbered sections
  in Draft A rely on line breaks, hence the short paragraphs.
- Draft C deliberately spends its middle on the *comment I wrote to justify
  deleting something I didn't understand*, not on the bug. The bug is ordinary.
  Confidently rationalising a change is the part that generalises, and it is the
  part an audience can use.
- Deliberately not claimed: that this is production-ready (it is v0.2.0 with no
  users), that it is faster than the C# original (it is roughly level on time, and
  memory is the difference), or that 214/303 is parity. It isn't, and the post
  gives the real number.
- The AI framing: the last post was upfront that an agent did the work and that
  directing it and catching its errors was the skill. Same applies here and the
  drafts are written in first person throughout. If you want the framing
  explicit, the natural place is right after "I'd deleted the fields ... as
  cleanup", because the confident wrong comment is a very recognisable failure
  mode, and naming it makes the story more useful, not less.
- Before posting: the 29 documented disagreements are not individually
  diagnosed. If someone asks "what are they?", the honest answer is "tracked in
  testdata/corpus_baseline.tsv, not yet triaged one by one." Don't let a comment
  catch you claiming more.
