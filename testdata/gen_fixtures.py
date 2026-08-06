#!/usr/bin/env python3
"""Generate the test fixtures for go-excel-reader.

These files are produced by openpyxl (xlsx) and xlwt (xls BIFF8), both of which
are independent of this library. That matters: fixtures written by our own code
would only prove the reader agrees with the writer, not that either agrees with
the format. Regenerate with:

    ./xlvenv/bin/python testdata/gen_fixtures.py testdata/
"""

import datetime
import os
import random
import sys

import openpyxl
import xlsxwriter
import xlwt

# Two xlsx writers are used on purpose. openpyxl stores text inline, as
# <c t="inlineStr"><is><t>, while xlsxwriter stores it in the shared string
# table and writes <c t="s"> with an index. Both are legal and both occur in
# the wild, and they exercise entirely different code paths in the reader.
# Using only one of them leaves half the string handling untested -- which is
# exactly what happened before xlsxwriter was added here.

# The same logical table, written into xlsx, xls and csv. The cross-format test
# asserts all three decode to identical values, so this data deliberately mixes
# strings, whole numbers and decimals.
SIMPLE_HEADER = ["Name", "Qty", "Price"]
SIMPLE_ROWS = [
    ["Widget", 10, 9.99],
    ["Gadget", 3, 19.5],
    ["Doohickey", 25, 0.5],
]

# One column per Go type the reader is specified to return.
TYPES_HEADER = ["str", "int", "float", "date", "bool", "empty"]
TYPES_ROWS = [
    ["hello", 42, 3.14159, datetime.datetime(2023, 6, 15), True, None],
    ["world", -7, -0.5, datetime.datetime(1999, 12, 31, 23, 59, 59), False, None],
]

DATE_FORMATS = [
    ("m/d/yy", datetime.datetime(2023, 1, 2)),
    ("d-mmm-yy", datetime.datetime(2023, 3, 14)),
    ("yyyy-mm-dd", datetime.datetime(2024, 12, 25)),
    # Deliberately not a date near serial 0: the 1900 system's phantom leap day
    # makes serials below 61 ambiguous, and writers disagree there. That edge is
    # tested directly in numfmt, where the expectation can be stated exactly,
    # rather than through whatever this writer happens to emit.
    ("h:mm AM/PM", datetime.datetime(2022, 5, 4, 13, 30)),
    ("yyyy-mm-dd hh:mm:ss", datetime.datetime(2020, 2, 29, 8, 5, 3)),
    ("[$-409]mmmm d, yyyy", datetime.datetime(2021, 7, 4)),
]


def write_simple_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(SIMPLE_HEADER)
    for row in SIMPLE_ROWS:
        ws.append(row)
    wb.save(path)


def write_simple_xls(path):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Sheet1")
    for c, v in enumerate(SIMPLE_HEADER):
        ws.write(0, c, v)
    for r, row in enumerate(SIMPLE_ROWS, start=1):
        ws.write(r, 0, row[0])
        ws.write(r, 1, row[1])
        ws.write(r, 2, row[2])
    wb.save(path)


def write_simple_csv(path):
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(",".join(SIMPLE_HEADER) + "\n")
        for row in SIMPLE_ROWS:
            # Render numbers the way a spreadsheet export would, so that type
            # inference has the same job it would have on a real file.
            f.write("%s,%d,%s\n" % (row[0], row[1], repr(row[2])))


def write_types_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Types"
    ws.append(TYPES_HEADER)
    for r, row in enumerate(TYPES_ROWS, start=2):
        for c, v in enumerate(row, start=1):
            if v is None:
                continue
            cell = ws.cell(row=r, column=c, value=v)
            if isinstance(v, datetime.datetime):
                cell.number_format = "yyyy-mm-dd hh:mm:ss"
    wb.save(path)


def write_types_xls(path):
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Types")
    date_style = xlwt.easyxf(num_format_str="YYYY-MM-DD HH:MM:SS")
    for c, v in enumerate(TYPES_HEADER):
        ws.write(0, c, v)
    for r, row in enumerate(TYPES_ROWS, start=1):
        ws.write(r, 0, row[0])
        ws.write(r, 1, row[1])
        ws.write(r, 2, row[2])
        ws.write(r, 3, row[3], date_style)
        # In Python bool is a subclass of int, so xlwt's generic write() would
        # store True as the number 1. Force a real BOOLERR cell.
        ws.row(r).set_cell_boolean(4, row[4])
    wb.save(path)


def write_multisheet_xlsx(path):
    wb = openpyxl.Workbook()
    first = wb.active
    first.title = "Alpha"
    first.append(["a1", "a2"])
    first.append([1, 2])

    second = wb.create_sheet("Beta")
    second.append(["b1"])
    second.append([100])

    third = wb.create_sheet("Gamma With Spaces")
    third.append(["g1", "g2", "g3"])
    wb.save(path)


def write_multisheet_xls(path):
    wb = xlwt.Workbook()
    a = wb.add_sheet("Alpha")
    a.write(0, 0, "a1")
    a.write(0, 1, "a2")
    a.write(1, 0, 1)
    a.write(1, 1, 2)

    b = wb.add_sheet("Beta")
    b.write(0, 0, "b1")
    b.write(1, 0, 100)

    c = wb.add_sheet("Gamma With Spaces")
    c.write(0, 0, "g1")
    c.write(0, 1, "g2")
    c.write(0, 2, "g3")
    wb.save(path)


def write_dates_xlsx(path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Dates"
    ws.append(["format", "value"])
    for r, (fmt, value) in enumerate(DATE_FORMATS, start=2):
        ws.cell(row=r, column=1, value=fmt)
        cell = ws.cell(row=r, column=2, value=value)
        cell.number_format = fmt
    # A number that is emphatically not a date, carrying a format whose literal
    # text contains date letters. A reader that scans for d/m/y without
    # honouring quoted literals turns this into the year 2093.
    ws.cell(row=len(DATE_FORMATS) + 2, column=1, value="not-a-date")
    trap = ws.cell(row=len(DATE_FORMATS) + 2, column=2, value=1234.5)
    trap.number_format = '"Paid on day "0.0'
    wb.save(path)


def write_gaps_xlsx(path):
    """A sheet with holes, to pin down row and column index handling."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gaps"
    ws["A1"] = "first"
    # Row 3 is skipped entirely; column B is skipped within row 2.
    ws["A2"] = "second"
    ws["C2"] = "third"
    # Column AA exercises the bijective base-26 conversion; a naive parser
    # places it at 26*0+0 or 27 instead of 26.
    ws["AA4"] = "twenty-six"
    ws["B6"] = "sixth"
    wb.save(path)


def write_sst_xlsx(path):
    """The simple table, written through the shared string table.

    Same data as simple.xlsx, different string storage, so the two files
    together prove the reader handles both representations identically.
    """
    wb = xlsxwriter.Workbook(path, {"in_memory": True})
    ws = wb.add_worksheet("Sheet1")
    for c, v in enumerate(SIMPLE_HEADER):
        ws.write_string(0, c, v)
    for r, row in enumerate(SIMPLE_ROWS, start=1):
        ws.write_string(r, 0, row[0])
        ws.write_number(r, 1, row[1])
        ws.write_number(r, 2, row[2])
    wb.close()


def write_large_xlsx(path, rows=200_000):
    """A big workbook with a large shared string table.

    Every row carries two unique strings, so the table holds 2*rows entries.
    That table is the largest single allocation a reader makes, and storing it
    as a []string is the mistake this fixture is sized to catch.
    """
    wb = xlsxwriter.Workbook(path, {"in_memory": True})
    ws = wb.add_worksheet("Large")
    rnd = random.Random(20260806)
    for i in range(rows):
        ws.write_string(i, 0, "row-%d-%s" % (i, rnd.randbytes(6).hex()))
        ws.write_string(i, 1, "tag-%d-%s" % (i, rnd.randbytes(6).hex()))
        ws.write_number(i, 2, i)
        ws.write_number(i, 3, i * 1.5)
        ws.write_boolean(i, 4, i % 2 == 0)
    wb.close()


def write_repeat_xlsx(path, rows=200_000):
    """Same shape as large.xlsx, but with a tiny shared string table.

    The two files have identical row and cell counts and differ only in how many
    distinct strings they hold. Comparing them separates the cost of streaming
    rows from the cost of the string table, which is the only part of the reader
    whose memory grows with the file.
    """
    wb = xlsxwriter.Workbook(path, {"in_memory": True})
    ws = wb.add_worksheet("Repeat")
    pool = ["category-%d" % i for i in range(100)]
    rnd = random.Random(1)
    for i in range(rows):
        ws.write_string(i, 0, rnd.choice(pool))
        ws.write_string(i, 1, rnd.choice(pool))
        ws.write_number(i, 2, i)
        ws.write_number(i, 3, i * 1.5)
        ws.write_boolean(i, 4, i % 2 == 0)
    wb.close()


def write_inline_xlsx(path, rows=2000):
    """A workbook whose strings are stored inline rather than shared."""
    wb = openpyxl.Workbook(write_only=True)
    ws = wb.create_sheet("Inline")
    for i in range(rows):
        ws.append(["inline-%d" % i, i, i * 0.25])
    wb.save(path)


def write_large_xls(path, rows=65535):
    """A big BIFF8 workbook.

    The strings vary in length on purpose. The shared string table spills across
    CONTINUE records, and a string landing across one of those boundaries is the
    case that carries a restated flags byte -- the single most delicate part of
    the BIFF reader. Random lengths guarantee the case occurs.
    """
    wb = xlwt.Workbook()
    ws = wb.add_sheet("Large")
    rnd = random.Random(20260806)
    for i in range(rows):
        n = rnd.randint(1, 60)
        ws.write(i, 0, "s%d-%s" % (i, "x" * n))
        ws.write(i, 1, i)
        ws.write(i, 2, i * 1.5)
        ws.write(i, 3, "t%d-%s" % (i, "y" * rnd.randint(1, 40)))
    wb.save(path)


def write_csv_variants(directory):
    with open(os.path.join(directory, "tab.tsv"), "w", encoding="utf-8", newline="\n") as f:
        f.write("Name\tQty\tPrice\n")
        for row in SIMPLE_ROWS:
            f.write("%s\t%d\t%s\n" % (row[0], row[1], repr(row[2])))

    # A UTF-8 BOM ahead of the header. Left in place it becomes part of the
    # first header's text, so "Name" stops comparing equal to "Name".
    with open(os.path.join(directory, "bom.csv"), "wb") as f:
        f.write(b"\xef\xbb\xbf")
        body = ",".join(SIMPLE_HEADER) + "\n"
        for row in SIMPLE_ROWS:
            body += "%s,%d,%s\n" % (row[0], row[1], repr(row[2]))
        f.write(body.encode("utf-8"))

    with open(os.path.join(directory, "quoted.csv"), "w", encoding="utf-8", newline="\n") as f:
        f.write('a,b,c\n')
        f.write('"has,comma","has ""quote""","has\nnewline"\n')
        f.write('plain,,trailing\n')

    # Latin-1 text, to exercise WithCSVEncoding.
    with open(os.path.join(directory, "latin1.csv"), "wb") as f:
        f.write("name,city\nJosé,Málaga\nRené,Köln\n".encode("latin-1"))


def main():
    out = sys.argv[1] if len(sys.argv) > 1 else "testdata"
    os.makedirs(out, exist_ok=True)
    p = lambda name: os.path.join(out, name)

    write_simple_xlsx(p("simple.xlsx"))
    write_simple_xls(p("simple.xls"))
    write_simple_csv(p("simple.csv"))
    write_types_xlsx(p("types.xlsx"))
    write_types_xls(p("types.xls"))
    write_multisheet_xlsx(p("multisheet.xlsx"))
    write_multisheet_xls(p("multisheet.xls"))
    write_dates_xlsx(p("dates.xlsx"))
    write_gaps_xlsx(p("gaps.xlsx"))
    write_sst_xlsx(p("sst.xlsx"))
    write_csv_variants(out)
    if "--no-large" not in sys.argv:
        write_inline_xlsx(p("inline.xlsx"))
        write_large_xlsx(p("large.xlsx"))
        write_repeat_xlsx(p("repeat.xlsx"))
        write_large_xls(p("large.xls"))

    for name in sorted(os.listdir(out)):
        full = p(name)
        if os.path.isfile(full) and not name.endswith(".py"):
            print("%-20s %10.1f KB" % (name, os.path.getsize(full) / 1024))


if __name__ == "__main__":
    main()
