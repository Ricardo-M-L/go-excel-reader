# Decision log

Calls made, what was ambiguous, what was chosen, and what was traded away.

Separate from the chapters on purpose. The chapters are the narrative; this is the set of
places where more than one answer was defensible and one had to be picked.

---

## D1 — Install a newer Go toolchain rather than lower the module's floor
**2026-08-06**

**Ambiguity:** The spec required Go 1.22+. The machine had 1.21.6, which refuses to build
a module declaring a newer version.

**Chose:** Download and install Go 1.26.5 alongside the existing toolchain.

**Why:** The alternative — declaring `go 1.21` in `go.mod` — silently changes the contract
the library publishes. `clear()` on a map and `b.Loop()` in benchmarks both need newer
versions, and discovering that halfway through would mean rewriting working code.

**Traded:** A 67 MB download and five minutes before any code was written.

---

## D2 — Add two methods to an interface that was specified exactly
**2026-08-06**

**Ambiguity:** The target API was given precisely, with an instruction to "implement
exactly this." Two things it also specified could not be built against it.

**Chose:** Add `Err() error` and `Cell(col) (Cell, bool)` to the `Reader` interface.

**Why:** A `Read() bool` cursor has nowhere to report a mid-stream failure, so a truncated
file is indistinguishable from a short one. And `ReadAll` is specified to return an
`error` — without `Err` on the interface, that error could only ever be `nil`. `ReadAll`
is likewise specified to return cells carrying a `Raw` field, and nothing else on the
interface exposes raw text.

Both additions are implied by the specification's own requirements. The alternative was to
ship an API that quietly loses errors.

**Result:** Stated explicitly in the package doc and the README rather than slipped in.
Widening an interface somebody specified on purpose deserves to be visible.

**Traded:** Strict literal compliance, for an API that can report failure.

---

## D3 — A leaf package for the value model, not duplication across three readers
**2026-08-06**

**Ambiguity:** The specified layout has `xlsx/`, `xls/` and `csv/` as siblings, with the
public `Reader` in the root. Each backend needs to produce typed cells, and the root needs
to consume them — which is an import cycle if the `Cell` type lives in the root.

**Chose:** `internal/cells`, holding `Cell` and a `Source` interface. The root aliases
`Cell` to it and implements every typed getter exactly once against `Source`.

**Why:** The alternative is writing `GetInt` three times. Three copies of a coercion rule
drift, and the drift is invisible — the day `xls` and `xlsx` disagree about whether
`GetInt` accepts a whole-numbered float, no test fails.

**Traded:** One package not in the specified structure.

---

## D4 — Refuse BIFF5 and earlier rather than guess
**2026-08-06**

**Ambiguity:** The `.xls` container format predates BIFF8 by several versions. Older files
store text in a codepage declared by a `CODEPAGE` record instead of UTF-16.

**Chose:** Read the version from the `BOF` record and return
`ErrUnsupportedBIFF` for anything below BIFF8.

**Why:** Partial support here fails silently. Decoding codepage bytes as if they were
Latin-1 produces text that looks like text — mojibake, not an error — and a user has no
way to tell a bad decode from a bad file. A clear refusal is more useful than plausible
garbage.

**Traded:** Files from Excel 5 and 95. The spec named this as acceptable for v1.0.

---

## D5 — The shared string table as a byte blob, not a []string
**2026-08-06**

**Ambiguity:** The obvious representation of a string table is `[]string`. The spec
suggested indexing byte offsets and lazily re-reading for very large tables.

**Chose:** One contiguous `[]byte` blob plus an `[]int64` offset index, built once by
streaming the XML.

**Why:** True lazy re-reading is not actually available: the table lives inside a
compressed zip member, so seeking backwards into it means re-inflating from the start —
O(n²) across a full pass. The blob gets most of the benefit with none of that. For 400,000
entries it is two allocations and zero pointers instead of 400,000 of each, and the
garbage collector never traces into a `[]byte`.

**Result:** 26.9 MB peak heap on a table whose XML is 15.4 MB, against a 50 MB budget.

**Traded:** The table is fully resident. This reader is not constant-memory, and the
README now says so.

---

## D6 — No `unsafe.String`, despite the free win
**2026-08-06**

**Ambiguity:** `SharedStrings.At` copies when converting a blob slice to a string.
`unsafe.String` over an immutable blob would make lookups allocation-free and remove
roughly 400,000 allocations per pass.

**Chose:** Keep the copy.

**Why:** The memory budget is met with margin without it. Adding `unsafe` to an otherwise
plain Go library imposes a real cost on anyone who audits dependencies, in exchange for a
benchmark number nobody asked for. Reversible later if profiling on a real workload
justifies it.

**Traded:** Some allocation churn. Recorded in `NOTES.md` so the option isn't lost.

---

## D7 — Strict typed getters, except between int and float
**2026-08-06**

**Ambiguity:** Should `GetString` on a numeric cell format the number? Should `GetFloat`
on `"3.5"` parse it?

**Chose:** No. Wrong-type access reports `false`. The single exception is that `int64` and
`float64` interconvert, with floats accepted as integers only when integral and exact.

**Why:** Silent string-to-number coercion is how spreadsheet readers quietly corrupt data.
The exception earns its place because whether a number lands as `int64` or `float64`
depends on incidental details of how the file was written — which is not something a
caller should have to reason about. `Cell.Raw` is there for anyone who disagrees.

---

## D8 — Emit the rows the file omitted
**2026-08-06**

**Ambiguity:** A sheet whose first value is in row 5 contains no records for rows 1–4. A
reader can skip them or synthesise them.

**Chose:** Emit them as empty rows, so `RowIndex()` always matches the row number shown in
Excel.

**Why:** A user comparing output against the spreadsheet on screen is the common case, and
an index that silently shifts under a gap is worse than a few empty rows. It also matches
what ExcelDataReader does, which the parity check would have caught either way.

**Traded:** A pathologically sparse sheet costs one empty row per gap.

---

## D9 — Type inference off by default for CSV
**2026-08-06**

**Ambiguity:** The spec said treat CSV cells as strings by default, with inference
optional. But the cross-format test requires xlsx, xls and csv to produce identical
values, which needs inference on.

**Chose:** Default off, with `WithCSVTypeInference(true)` to enable it, and the
cross-format test enables it explicitly.

**Why:** A CSV carries no type information, so any inference is a guess the file cannot
confirm. Making the test ask for it keeps the default honest and makes the test's
dependency on inference visible rather than assumed.

**Related:** Date formats accepted by inference are limited to unambiguous ISO forms.
`03/04/2024` is two different dates in two different countries and nothing in the file
says which; a reader that picks one is wrong half the time and never says so.

---

## D10 — Dates in UTC, always
**2026-08-06**

**Chose:** Every `time.Time` returned is UTC.

**Why:** Excel serials carry no timezone and no DST rules. Attaching a local zone would
make the same file decode differently on different machines, and would shift values for
anything that later formats them. Documented rather than left to be discovered.

---

## D11 — Generate fixtures with three independent writers
**2026-08-06**

**Ambiguity:** Fixtures could be written by this library, hand-built as byte slices, or
generated by third-party tools.

**Chose:** openpyxl, xlsxwriter and xlwt.

**Why:** Fixtures written by our own writer prove the reader agrees with the writer, not
that either agrees with the format. The spec explicitly asked for real files rather than
mocked bytes.

**Result:** This decision was made for the right reason and still nearly failed, because
the *first* version used only openpyxl — which stores strings inline, leaving the shared
string reader untested. See `chapters/03-the-fixtures-that-tested-nothing.md`. Two writers
that differ on a property you care about is the actual requirement; "third-party" alone is
not enough.

**Traded:** A Python toolchain to regenerate fixtures. Small fixtures are committed; the
large ones are generated, so `go get` doesn't pull 18 MB of test data.

---

## D12 — Verify against the original, not only against my own expectations
**2026-08-06**

**Ambiguity:** The spec listed "produces identical output to ExcelDataReader" as a success
criterion without saying how to check it.

**Chose:** Build two dumpers — one in C# against the real ExcelDataReader 3.9.0, one in Go
— emitting a canonical line-per-cell form, and diff them.

**Why:** Every other test in the project encodes what *I* believe the format means. This
is the only check that can catch a misunderstanding I hold consistently.

**Key detail:** numbers are emitted as raw IEEE 754 bit patterns (`%016X`), not as text.
C# and Go disagree about how to render a `double`, and that disagreement would appear as a
false difference on every non-integral value — noise that buries real differences.

**Result:** Byte-identical on all twelve fixtures, including 1,000,000 cells. Setup cost
about twenty minutes.

---

## D13 — Cut v0.1.1 rather than move the v0.1.0 tag
**2026-08-06**

**Ambiguity:** The inaccurate memory claim was found after v0.1.0 was tagged, pushed, and
already fetched through the module proxy.

**Chose:** Fix forward as v0.1.1.

**Why:** The proxy had already served and cached v0.1.0. Moving a published tag means two
different trees answer to the same version, which is worse than an inaccurate sentence in
a README for a few hours.

**Traded:** A version number, and a permanent record that the first one shipped with a
claim I couldn't defend. That record is the point.

---

## D14 — Run against ExcelDataReader's own test corpus
**2026-08-06**

**Ambiguity:** The library was tagged and green, verified byte-for-byte against the
reference implementation on twelve fixtures. That looked like enough.

**Chose:** Point the same harness at all 303 `.xls` and `.xlsx` files in
ExcelDataReader's `src/TestData`.

**Why:** Twelve fixtures I wrote can only contain cases I thought of. Upstream's
corpus is a decade of regression files, each added because it broke a reader.
It is the cheapest source of adversarial input that exists for this problem, and
it was already on disk after cloning the repo to read the source.

**Result:** 207 identical, 37 differing, 26 gaps, and **zero** files read here that
the reference refuses. Three distinct defects found; after fixes, 214 identical
and 29 differing. Cost about twenty minutes to set up.

**Traded:** The comfortable belief that "verified against the original" meant
something broader than it did.

---

## D15 — Narrow the empty-string rule instead of keeping the first fix
**2026-08-06**

**Ambiguity:** ExcelDataReader reports an empty inline string as absent. The obvious
fix — treat empty text as an empty cell — repaired four corpus files and broke eight.

**Chose:** Apply the rule only to inline strings. Empty entries in the shared string
table, and empty `LABEL` records in `.xls`, stay as the empty string.

**Why:** The reference distinguishes a cell that was written and holds nothing from a
value the file deliberately interned. Whether that distinction is intentional upstream
or an artefact of two code paths, it is the behaviour, and parity means matching the
behaviour rather than the behaviour I would have designed.

**Result:** Four files fixed, none broken.

**Traded:** A rule that is harder to state. The comment in `sax.go` explains the
asymmetry rather than hiding it, because a future reader will otherwise "simplify" it
straight back into the broken version.

---

## D16 — Make the error-cell behaviour an option, defaulting to the reference
**2026-08-06**

**Ambiguity:** A cell holding `#DIV/0!` reads as absent in ExcelDataReader. This
library returned the display text, which is arguably more useful.

**Chose:** Default to the reference behaviour; `WithErrorValues(true)` restores the
text. `Cell.Raw` carries it in both modes.

**Why:** The behaviour was defensible; the README was not. It claimed output identical
to the original while deliberately differing. Given a choice between weakening the
claim and matching the reference, matching costs one option and keeps the strong claim
true.

**Traded:** A slightly larger API, and a default that is less useful than the
alternative for anyone who does not read the docs.

---

## D17 — Pin the corpus result with hashes of our own output
**2026-08-06**

**Ambiguity:** The corpus comparison needs .NET and a 379-file checkout, neither of
which belongs in this repository or in CI. But an unpinned number in a README rots.

**Chose:** Commit `testdata/corpus_baseline.tsv` — verdict plus a SHA-256 of *our*
dump for each file. `TestCorpusBaselineShape` asserts the counts with no corpus
present; `TestCorpusParity` re-derives every dump when `EXCEL_CORPUS` is set.

**Why:** Hashing our own output is what removes the .NET dependency from the check.
The hashes came from a run where the comparison really happened, so pinning them pins
the parity result by proxy. Refreshing hashes deliberately does not re-derive the
verdicts, which still require the full harness — so a change in what we agree on has
to be measured, not assumed.

**Result:** The first generator was a shell script that hashed empty output for
rejected files while the Go test wrote `REJECTED`, reporting seventeen false changes.
Regeneration now runs through the same function the check uses. The generator and the
checker being two implementations of one idea is the same mistake this project keeps
making.

---

## D18 — Buffer a bounded window of rows instead of assuming BIFF record order
**2026-08-06**

**Ambiguity:** The xls reader emitted a row whenever the row number changed, which
assumes cell records arrive grouped by row and in increasing order. Nine corpus files
proved otherwise: a sheet may emit MULBLANK records covering a block and only then the
cells that fill it, so those sheets were reported twice — once blank, once populated.

**Chose:** Accumulate rows in a map keyed by absolute row index, emitting only rows
strictly below the highest one seen, with a 1024-row window.

**Why:** ExcelDataReader buffers a block of rows and places cells by absolute index for
exactly this reason; BIFF makes no ordering guarantee within a row block. Reading the
reference settled a question that guessing would not have.

The window rather than a whole sheet: a hostile file could otherwise force buffering of
65,536 rows. A thousand covers every real displacement seen (the worst was 98) while
capping worst-case memory.

**Result:** Nine files fixed. The first implementation emitted the top buffered row
before it was complete, truncating it — caught by an existing test at exactly the
window boundary.

**Traded:** Peak memory for an xls sheet is now bounded by the window rather than by a
single row. xlsx is unaffected and still streams row by row.

---

## D19 — Pin the reference implementation's timezone when comparing
**2026-08-06**

**Ambiguity:** Several corpus files differed by exactly the local UTC offset.

**Chose:** Run ExcelDataReader under `TZ=UTC` in the parity harness, and keep returning
UTC unconditionally here.

**Why:** ExcelDataReader parses OOXML strict ISO dates and converts them through the
machine's local timezone: `strict/Open.xlsx` yields 16:00, 11:00 or 02:00 under
New York, UTC and Tokyo. Under `TZ=UTC` it agrees with this library exactly, so the
disagreement is not about the value but about whether the answer should depend on the
host.

Following the reference here would mean reproducing a bug. Pinning its timezone
compares the two readers rather than two locales.

**Traded:** "Byte-identical to ExcelDataReader" now carries a condition. Stating the
condition is better than a claim that silently depends on where it was measured.

---

## D20 — Bound every length taken from a file header
**2026-08-06**

**Ambiguity:** `readDIFAT` used the header's FAT-sector count as an allocation capacity.
Fuzzing found a 2 KB file declaring `0xFFFFFFFF` sectors, which asks for a 17 GB
allocation and kills the process with a fatal out-of-memory.

**Chose:** Reject counts exceeding what the file could physically contain, and bound
`columnIndex` by the format's 16,384-column limit for the same reason.

**Why:** A Go fatal out-of-memory is not recoverable — a caller cannot defend against it
the way it can against a panic. For any service accepting spreadsheet uploads this is a
denial of service reachable with a file small enough to paste into a chat message.

The bound is derived from the file rather than picked: a compound file cannot contain
more FAT sectors than it contains sectors.

**Result:** 2.1 million fuzz executions across three targets afterwards with no further
failures. The crashing input is committed under `testdata/fuzz/` and runs as a normal
test.
