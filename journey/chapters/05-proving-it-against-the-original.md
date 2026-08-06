# Chapter 5: Making it prove itself

*2026-08-06. Everything green. Which proves less than it feels like.*

At this point the library read all three formats, every test passed, `go vet` and
`staticcheck` were clean, and it was tagged and published.

Every one of those tests encodes what I believe the formats mean. If I have
misunderstood something consistently, by reading a field at the wrong offset or
applying a rule in the wrong direction, my tests agree with me, because I wrote
both.

There is exactly one check that can catch a misunderstanding held consistently:
compare against an implementation that does not share it.

## The harness

Two dumpers emitting one line per non-empty cell:

```
sheetIndex \t sheetName \t rowIndex \t colIndex \t typeTag \t value
```

One in C# against the real ExcelDataReader 3.9.0, one in Go. Diff the output.

.NET 9 was already on the machine, the NuGet package installs in seconds, and the
whole thing took about twenty minutes.

## The one decision that made it work

Numbers are emitted as raw IEEE 754 bit patterns, not as text.

```csharp
text = BitConverter.DoubleToInt64Bits(dbl).ToString("X16");
```
```go
return "N", fmt.Sprintf("%016X", math.Float64bits(x))
```

C# and Go disagree about how to render a `double` as a string. `1E+15` versus
`1e+15`, different precision defaults, different round-trip rules. Comparing
formatted numbers would have produced a difference on nearly every non-integral
value, which is noise that buries the real signal completely.

Hashing the bits removes the question entirely. Two numbers are equal if they are
the same number, which is what I actually wanted to know.

The same reasoning applies to the type tag. `S`/`N`/`B`/`D` distinguishes a cell
holding the *string* `"42"` from one holding the *number* 42, a difference that
formatting alone would hide, and precisely the kind of bug a spreadsheet reader
has.

## What it found

Nothing. All twelve fixtures matched byte for byte, including the 1,000,000-cell
workbook and the 8.9 MB `.xls` whose string table spans hundreds of `CONTINUE`
records.

That was genuinely reassuring about the hard parts. The RK decoding, the
continuation flags byte, the date resolution chain and the leap-year handling
were all independently confirmed by an implementation with a decade of production
use.

## What I then wrote in the README

> All fixtures match exactly.

True. And it invited a reading it had not earned.

Twelve files matched. I chose all twelve. A sample you choose cannot surprise you.
It can only confirm that your implementation agrees with your understanding, on
the cases your understanding produced.

The harness was right. The corpus it ran against was the problem, and fixing that
is chapter 7.

## Worth stealing

The pattern generalises past spreadsheets. When porting something that already
exists, build the differential harness early and cheaply. Not a test suite: a
canonical dump plus `diff`.

Two things make it work:

1. **Canonicalise away everything you do not care about.** Float formatting,
   timezone rendering, key order. Whatever your two languages disagree about for
   reasons unrelated to correctness, remove it before comparing.
2. **Keep the output diffable.** One line per fact, sorted by construction. When
   something differs you want to read the difference, not decode it.

Twenty minutes bought a check that no amount of my own test-writing could replace,
and later, a place to point 303 files I had never seen.
