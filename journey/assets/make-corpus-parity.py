# Unit chart: one square per file in ExcelDataReader's test corpus, coloured by
# how our output compared with the reference implementation's.
#
# A unit chart is the right form here because the unit is meaningful: each square
# is one real spreadsheet that broke a reader once. A bar of "94%" hides that.
# It is also on-theme, since the squares read as spreadsheet cells.
#
# Cells are sorted by category so only consecutive categories ever touch, which
# keeps the palette on the validated adjacent pairlist rather than all-pairs.

W, H = 1200, 1500
SURFACE   = "#1a1a19"
TEXT      = "#ffffff"
SECONDARY = "#c3c2b7"
MUTED     = "#8a8880"

# Categorical slots 1-4, dark steps, validated: all checks PASS
CATS = [
    ("#3987e5", 229, "Byte-identical to the original"),
    ("#d95926",  14, "Both read it, output differs"),
    ("#199e70",  27, "Only the original reads it"),
    ("#c98500",  33, "Both refuse it"),
]
TOTAL = sum(c[1] for c in CATS)
assert TOTAL == 303, TOTAL

COLS, CELL, GAP = 22, 40, 5
grid_w = COLS * CELL + (COLS - 1) * GAP
x0 = 80  # flush with the text column; a centred grid reads as accidental
y0 = 430

cells = []
for color, n, _ in CATS:
    cells += [color] * n

def esc(s): return s.replace("&", "&amp;").replace("<", "&lt;")

p = []
p.append(f'<svg xmlns="http://www.w3.org/2000/svg" width="{W}" height="{H}" viewBox="0 0 {W} {H}">')
p.append(f'<rect width="{W}" height="{H}" fill="{SURFACE}"/>')

MONO = "DejaVu Sans Mono, monospace"
SANS = "DejaVu Sans, sans-serif"

# --- header ---
p.append(f'<text x="80" y="120" font-family="{MONO}" font-size="54" font-weight="700" fill="{TEXT}">go-excel-reader</text>')
p.append(f'<text x="80" y="168" font-family="{SANS}" font-size="27" fill="{SECONDARY}">ExcelDataReader, ported to Go</text>')
p.append(f'<line x1="80" y1="212" x2="{W-80}" y2="212" stroke="#3a3936" stroke-width="2"/>')

# --- the claim being visualised ---
p.append(f'<text x="80" y="286" font-family="{SANS}" font-size="34" fill="{TEXT}">I checked it against the original library&#8217;s own test corpus:</text>')
p.append(f'<text x="80" y="336" font-family="{SANS}" font-size="34" fill="{TEXT}">303 spreadsheets collected over a decade, each one a real bug.</text>')
p.append(f'<text x="80" y="392" font-family="{MONO}" font-size="25" fill="{MUTED}">one square = one file</text>')

# --- unit chart ---
for i, color in enumerate(cells):
    cx = x0 + (i % COLS) * (CELL + GAP)
    cy = y0 + (i // COLS) * (CELL + GAP)
    p.append(f'<rect x="{cx}" y="{cy}" width="{CELL}" height="{CELL}" rx="4" fill="{color}"/>')

rows = (len(cells) + COLS - 1) // COLS
legend_y = y0 + rows * (CELL + GAP) + 56

# --- legend: swatch + count + label, so identity is never colour alone ---
for i, (color, n, label) in enumerate(CATS):
    ly = legend_y + i * 52
    p.append(f'<rect x="80" y="{ly-22}" width="26" height="26" rx="4" fill="{color}"/>')
    p.append(f'<text x="122" y="{ly}" font-family="{MONO}" font-size="30" font-weight="700" fill="{TEXT}">{n:>3}</text>')
    p.append(f'<text x="196" y="{ly}" font-family="{SANS}" font-size="27" fill="{SECONDARY}">{esc(label)}</text>')

# --- the line that matters, called out ---
foot = legend_y + len(CATS) * 52 + 34
p.append(f'<line x1="80" y1="{foot}" x2="{W-80}" y2="{foot}" stroke="#3a3936" stroke-width="2"/>')
p.append(f'<text x="80" y="{foot+52}" font-family="{SANS}" font-size="28" fill="{TEXT}">Zero files parse here that the original refuses.</text>')
p.append(f'<text x="80" y="{foot+92}" font-family="{SANS}" font-size="25" fill="{MUTED}">Running it took twenty minutes and found three defects.</text>')
p.append(f'<text x="{W-80}" y="{foot+92}" text-anchor="end" font-family="{MONO}" font-size="24" fill="{MUTED}">jonathanloor.com</text>')

p.append('</svg>')
open("/home/bullwinkle/excel-go/journey/assets/corpus-parity.svg","w").write("\n".join(p))
print("svg written; cells:", len(cells), "rows:", rows, "bottom:", foot+92)
