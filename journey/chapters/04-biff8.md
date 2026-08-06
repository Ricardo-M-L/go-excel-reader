# Chapter 4: A filesystem inside a file, and one byte

*2026-08-06. The `.xls` reader. The hardest code in the project, and it worked
first try, which was not where the trouble came from.*

A modern `.xlsx` is a ZIP of XML you can read with two standard library packages.
A legacy `.xls` is neither of those things.

It is a Compound File Binary container: an entire filesystem, with sectors, a file
allocation table, and a directory tree, inside one file. The spreadsheet lives in
a stream inside it, as a flat sequence of binary records.

Nothing about the spreadsheet is visible until the filesystem works.

## Layer one: CFB

Sectors are typically 512 bytes. A FAT chains them. A directory of 128-byte
entries names the streams. There is a second, smaller allocation scheme called the
"mini FAT" for streams below 4096 bytes, which are packed into a stream belonging
to the root entry rather than given whole sectors.

Details that cost time:

- **Sector *n* starts at `(n+1) × sectorSize`.** The 512-byte header occupies the
  first sector, and for 4096-byte sectors it is padded out to a full one, which
  keeps the formula identical across versions.
- **The FAT's own sector list can outgrow the header.** 109 entries fit inline;
  beyond that they spill into DIFAT sectors, each ending with a pointer to the
  next. A chain walk, not a read.
- **Version 3 stream sizes are 32-bit.** The field is 64 bits wide and some
  writers leave junk in the top half, which reads as a multi-exabyte stream.
- **Chains can cycle.** A corrupt file will happily describe one. Every walk is
  bounded by the size of the table it walks.

Streams are read lazily: the sector chain is materialised once as a `[]uint32`,
then reads are index lookups. Opening a 200 MB workbook stream costs the chain
list and nothing else.

I also, in this chapter, deleted the directory tree's sibling and child pointers
as unused fields and wrote a comment explaining why that was fine. That decision
is chapter 7.

## Layer two: BIFF records

Records are a 2-byte type, a 2-byte length, and a payload. `BOF`, `EOF`,
`BOUNDSHEET`, `SST`, `LABELSST`, `NUMBER`, `RK`, `MULRK`, `BLANK`, `MULBLANK`,
`BOOLERR`, `FORMULA`, `XF`, `FORMAT`, `DATEMODE`.

`BOUNDSHEET` carries a sheet's name and an absolute offset to its own `BOF`, which
is what makes sheet-at-a-time reading possible without scanning.

### RK: four encodings in four bytes

`RK` packs a number into 32 bits using the low two bits as tags:

| bit 1 | bit 0 | Meaning |
|---|---|---|
| 0 | 0 | top 30 bits of an IEEE 754 double |
| 0 | 1 | same, then ÷100 |
| 1 | 0 | signed 30-bit integer |
| 1 | 1 | same, then ÷100 |

All four occur in real files. The trap is the integer variants: the value sits in
the top 30 bits, so a *logical* shift instead of an arithmetic one turns −7 into
268,435,449. The test builds each encoding by construction rather than copying
bytes out of a file, so the tag layout is explicit.

## The byte

A record payload cannot exceed 8224 bytes, so a large shared string table arrives
as one `SST` record followed by a run of `CONTINUE` records.

Concatenate them and parse the result, and almost everything works.

Almost. When a string's *characters* are split across a boundary, the first byte
of the continuation is not character data. It restates the compression flag,
because the second half of a string may be stored differently from the first
(BIFF8 stores text as UTF-16 unless every character fits in one byte).

Miss that byte and it is read as text. The table does not fail. It silently
desynchronises, and every string after that point comes back shifted by one byte
and garbled.

So the record reader keeps the boundary offsets instead of concatenating them
away, and the check lives inside the character loop, which is exactly right,
because a boundary that falls *between* strings carries no flags byte:

```go
if c.atContinuation() {
        flags, ok := c.u8()
        if !ok { return "", errors.New("xls: truncated continuation flags") }
        high = flags&0x01 != 0
}
```

### Testing something that is hard to hit on purpose

Two things make this actually tested rather than nominally tested.

The fixture generator writes strings of random length, so the split lands
mid-string somewhere in 65,535 rows. A test then walks every row asserting the
string still starts with `s<rowindex>-`, which proves the table stayed aligned
across every boundary.

And the unit test includes a negative control. It parses the same bytes *without*
the boundary information and asserts the result comes out wrong.

```go
if got == "abcd" {
        t.Fatal("parsing without continuation boundaries produced the correct string; " +
                "the boundary handling is not actually being exercised")
}
```

Without that, an assertion that `"abcd" == "abcd"` can pass for reasons having
nothing to do with the code. Chapter 3 is why that control exists.

## Refusing rather than guessing

BIFF5 and earlier store text in a codepage declared by a `CODEPAGE` record rather
than UTF-16. Supporting them halfway fails *silently*, because decoding codepage
bytes as Latin-1 produces text that looks like text.

So the reader parses `BOF`, checks the version, and refuses anything below BIFF8
with a message naming what it found. That refusal accounts for most of the 27
corpus files ExcelDataReader reads and this does not, and it is the right trade:
a clear "no" beats plausible garbage.

## Result

Every `.xls` test passed on the first run, including 65,535 rows of
continuation-spanning strings.

The hard code was fine. What broke, twice, was the layer meant to tell me whether
it was.
