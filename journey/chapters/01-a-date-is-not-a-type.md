# Chapter 1: A date is not a type

*2026-08-06. First code written. Out of the spec's order, deliberately.*

The plan ran xlsx first and number formats third. That order cannot work, because
the only question the xlsx reader needs answered about a number is *is this
actually a date?*, and answering it is the entire third phase.

Building xlsx first means stubbing date detection and returning later. The spec
also said never to leave a TODO in place. Those two instructions conflict, and
the honest way to satisfy the second is to reorder.

## The thing that makes spreadsheets hard

A spreadsheet does not store dates. It stores `45092.0` and, somewhere else
entirely, the fact that this particular cell should be *displayed* as a date.

The path from cell to answer has three hops:

```
<c r="D2" s="3">  →  cellXfs[3]  →  numFmtId=164  →  "yyyy-mm-dd"  →  it's a date
```

Break any link and a delivery date silently becomes the number 45217. Nothing
errors. The value is a perfectly good float.

## Trap one: there are two lists of <xf>

`styles.xml` contains `<cellStyleXfs>` and `<cellXfs>`. They hold identical-looking
`<xf>` elements. The first is named-style definitions; the second is what a cell's
`s=` attribute actually indexes. `cellStyleXfs` comes first in the file.

A parser that collects every `<xf>` it encounters binds every cell to the wrong
format, and does it consistently enough to look deliberate. The fix is one boolean
tracking which list you are inside, but you have to know the second list exists.

## Trap two: date letters inside literals

The obvious implementation of "is this format a date" is to look for `d`, `m`,
`y`, `h` or `s`.

Consider `"Paid on day "0.0`.

It contains `d`, `a` and `y`. It is a currency-ish number format with a quoted
literal. A naive scanner reports it as a date and converts 1234.5 into an
afternoon in 2093.

Things that must be skipped: quoted literals, `\`-escaped characters, `_` width
placeholders, `*` fill characters, and `[Red]` / `[$-409]` brackets. Things that
must *not* be skipped: `[h]`, `[mm]` and `[ss]`, which mean elapsed time and are
date tokens.

That single format string is now a fixture, sitting in `dates.xlsx` next to six
real date formats, precisely because it is the case a working implementation and
a broken one disagree on.

## Trap three: 'm' means two different things

`m` is months. `m` is also minutes. `mm/dd hh:mm` contains both, and they are the
same character.

Excel disambiguates by context: an `m` directly after an hour token, or directly
before a seconds token, is minutes. Otherwise it is a month. Three or more `m`s
are always a month name, since minutes have no names.

The test asserts both readings resolve correctly *inside one format string*,
which is the only version of the test that proves the context rule rather than a
global guess.

## The leap year that never happened

Excel believes 1900 was a leap year. It was not. The bug is deliberate, because
Lotus 1-2-3 had it and compatibility won.

So serial 60 is a 29th of February that did not exist, and every serial from 61
onward is one greater than a real day count.

The tempting fix is a correction term: `if serial >= 61 { serial-- }` buried in an
expression. Instead the epoch switches at the boundary, 1899-12-31 below 60 and
1899-12-30 at or above, which makes both branches checkable by hand:

| Serial | Date |
|---|---|
| 1 | 1900-01-01 |
| 59 | 1900-02-28 |
| 60 | 1900-02-28 *(the phantom; no correct answer exists)* |
| 61 | 1900-03-01 |

Serial 60 is mapped to the 28th so the value stays inside the real calendar
rather than rolling forward into March and colliding with 61.

There is also a second epoch, the 1904 system from classic Mac Excel, which is a
workbook-level property rather than a cell one. Every conversion takes the flag
rather than guessing.

## What this bought

`numfmt` was written before either reader and passed on the first run: five test
functions, including a round-trip across the leap-bug boundary.

That is not a boast. It is the setup for chapter 3, where the thing that broke
was not the hard code.

One decision recorded here: all times are returned in UTC. Excel serials carry no
timezone and no DST rules. Attaching a local zone would make the same file decode
differently on different machines, which is a bug that only appears in someone
else's CI.
