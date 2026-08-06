# Chapter 6 — The claim I couldn't defend

*2026-08-06. v0.1.0 was tagged, pushed, and already fetched through the module
proxy. Then someone asked a simple question.*

The question was: *is it really that memory efficient?*

The numbers said yes. 40 MB peak RSS on a 9.2 MB workbook, against 95 MB for
ExcelDataReader and 221 MB for excelize. Measured, repeated, real.

But before repeating them I reread what I had actually written in the README:

> The design goal is that peak memory tracks the widest row, not the file size.

That sentence is not a measurement. It is a description of what I set out to
build, written in the present tense as though it were a finding.

## Why it's wrong

Rows *do* stream. That half is true and remains true.

But the shared string table is held **fully resident**. It is the one structure
that is not streamed, by design (chapter 2), and on a string-heavy workbook it
*is* the memory profile — roughly 30 of the 39 MB. Memory tracks the number of
distinct strings, not the widest row and not the file size.

Nobody had challenged it. It had been true of nothing I measured, and it shipped.

## Measuring instead

The fix was to build the experiment the claim implies. Two workbooks with
**identical** row and cell counts — 200,000 rows, 1,000,000 cells — differing
only in how many distinct strings they contain:

| Workbook | Distinct strings | Peak Go heap | Peak RSS |
|---|---|---|---|
| `repeat.xlsx` (5.3 MB) | 100 | **3.5 MB** | 9.1 MB |
| `large.xlsx` (9.2 MB) | 400,000 | **26.9 MB** | 39.3 MB |

This is a better result than the claim I could not defend, and unlike the claim
it can be falsified:

- A million cells in 3.5 MB proves rows are not accumulating.
- The gap between 3.5 and 26.9 is the string table and nothing else.
- Row count appears in neither number.

The honest statement is longer and worth more: *rows stream in constant memory;
memory scales with the number of distinct strings; most real spreadsheets repeat
their values heavily, so it usually stays near the baseline.*

Plus the caveat that belongs next to it: **this is not a constant-memory
reader.** A workbook with millions of distinct strings will hold that table, and
no amount of streaming elsewhere avoids it. What is avoided is paying three times
over for it.

## Pinning the right thing

Both numbers are now tests, and the *baseline* one matters more than the total:

```go
// A million cells with a hundred distinct strings must stay near the
// streaming baseline. Accumulating rows would put this in the hundreds
// of megabytes.
const baselineBudget = 20 << 20
```

If row streaming regressed into accumulation, the string-heavy fixture would
still look plausible — it is *supposed* to use tens of megabytes, so a regression
hides inside the expected number. The repeating fixture has nowhere to hide.

A single test on the big file would have passed through the regression. Two
tests, separating the baseline from the thing that legitimately grows, is what
makes either of them meaningful.

## Fix forward, don't move the tag

v0.1.0 was already cached by the module proxy. Moving a published tag means two
different trees answer to the same version — worse than an inaccurate sentence
for a few hours.

So: v0.1.1, and a permanent record that the first release shipped a claim I could
not defend. That record is the useful part.

## The pattern, again

Chapter 3 was a test that passed without running the code. This is a claim that
sounded measured and wasn't. Same failure, different layer:

**Something asserted itself, and I did not check whether it could have been
otherwise.**

The specific tell is worth naming. "Peak memory tracks the widest row" is a
statement about a *relationship between variables* — and I had never varied
either one. I measured one file, once, and wrote a sentence describing a curve.

Any claim of the form "X scales with Y" needs at least two points, and they need
to differ in Y and nothing else. `repeat.xlsx` exists for exactly that reason and
took four minutes to generate.

A number you did not measure is a claim, and a claim in a README is a claim you
will eventually have to defend — usually to someone asking a perfectly friendly
question.
