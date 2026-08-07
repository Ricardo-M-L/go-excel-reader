# Chapter 0: Reading the spec before running it

*2026-08-06. No parsing code yet. This is what happened before any.*

The last project's lesson was to check a plan against reality before executing it. This plan held
up much better. It named the BIFF record numbers, the CFB magic bytes, the four RK
encodings and a gotchas table that turned out to be accurate on every entry.

So the review found fewer errors, but the ones it found were the kind that are expensive
to discover halfway through.

## The toolchain didn't exist

The spec required Go 1.22+. The machine had 1.21.6.

That is not a warning; a Go toolchain refuses outright to build a module declaring a newer
version than itself. The tempting fix is to write `go 1.21` in `go.mod` and move on.

It would have failed later and confusingly. `clear()` on a map is 1.21, but `b.Loop()` in
benchmarks is much newer, and both are used in the finished code. Discovering that after
writing the benchmark means rewriting working code to satisfy a constraint I invented.

Installed Go 1.26.5 alongside the existing one. Five minutes, before anything else.

## The specified interface could not do what the spec asked of it

The target API was given precisely, with an instruction to implement exactly it:

```go
type Reader interface {
    NextSheet() bool
    Read() bool
    GetValue(col int) any
    // ... typed getters, RowIndex, Close
}

func ReadAll(r Reader) (sheets [][]Row, err error)

type Cell struct {
    Value any
    Raw   string
}
```

Two things here cannot both be true.

`ReadAll` takes a `Reader` and returns an `error`. But nothing on `Reader` reports an
error, since `Read()` returns `bool`. So a file that is corrupt halfway through looks
exactly like a file that ended. `ReadAll`'s error could only ever be `nil`.

`ReadAll` also returns `Row`s of `Cell`s, and a `Cell` carries `Raw`. Nothing on `Reader`
exposes raw text. `GetValue` returns the typed value only. So `ReadAll` cannot populate
the field it is specified to return.

Both gaps are in the specification's own requirements, not in my reading of them. I added
`Err() error` and `Cell(col) (Cell, bool)` and said so explicitly in the package doc, the
README and the commit message.

The alternative was to comply literally and ship an API that silently loses errors. I do
not think "the spec said so" survives contact with a user whose truncated file read as
empty.

## What I chose not to build

The spec was clear about scope, and I kept it:

| Not ported | Why |
|---|---|
| `.xlsb` | Explicitly out of scope. Detected and rejected by name. |
| Encryption / passwords | Out of scope. |
| `AsDataSet()` | A .NET idiom with no Go equivalent worth inventing. |
| BIFF5 and earlier `.xls` | Refused with a clear error. See below. |

The BIFF5 decision is the one worth stating. Older `.xls` files store text in a codepage
declared by a `CODEPAGE` record rather than UTF-16. Partial support there fails
*silently*, because decoding codepage bytes as Latin-1 produces text that looks like text.
Mojibake, not an error. A user has no way to distinguish a bad decode from a bad file.

So the reader parses the `BOF` record, checks the version, and refuses anything below
BIFF8 with a message naming the version it found. A clear refusal beats plausible garbage.

## Building in dependency order, not spec order

The spec's phases run xlsx (2) → numfmt (3) → xls (4). But both readers need `numfmt` to
answer the only question that matters for typing a cell: *is this number a date?*

Building xlsx first would mean stubbing date detection, then coming back. So `numfmt` went
first, out of order.

Worth noting because the spec also said never to leave a TODO in place. Those two
instructions conflict if you follow the phase order literally. The honest way to satisfy
"no stubs" is to build the dependency first and reorder the phases.

## One structural addition

The specified layout puts `xlsx/`, `xls/` and `csv/` beside a root package holding the
public `Reader` and the `Cell` type.

That is a cycle. Each backend produces cells; the root consumes them; so `Cell` cannot
live in the root if the backends must name it.

The fix is a leaf package, `internal/cells`, holding `Cell` and a `Source` interface. The
root aliases `Cell` to it and implements every typed getter *once* against `Source`.

The alternative, `GetInt` written three times, is worse than it looks. Three copies of a
coercion rule drift, and the drift is invisible: the day `xls` and `xlsx` disagree about
whether `GetInt` accepts a whole-numbered float, nothing fails.

## Twenty minutes, and what it bought

Nothing dramatic. One toolchain install, two interface methods, one package, one reorder.

None of it needed spreadsheet expertise. It needed reading the spec as a document that
could be wrong, rather than as instructions. The same move as last time, applied to a plan
that was mostly right. The errors were smaller because the plan was better. They were not
zero, and the interface one would have shipped.
