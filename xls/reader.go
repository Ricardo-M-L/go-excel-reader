// Package xls reads legacy Excel binary workbooks (.xls) in the BIFF8 format
// used by Excel 97 and later.
//
// A .xls is a Compound File Binary container (see cfb.go) whose "Workbook"
// stream is a flat sequence of BIFF records (see biff.go). Reading it is a
// two-stage job: a globals substream declares the sheets, the shared strings
// and the number formats, and each sheet then has its own substream of cell
// records located by an offset from the globals.
package xls

import (
	"errors"
	"fmt"
	"io"
	"math"
	"strconv"

	"github.com/jloor/go-excel-reader/internal/cells"
	"github.com/jloor/go-excel-reader/numfmt"
)

// Reader streams a BIFF8 workbook one row at a time.
type Reader struct {
	cfb    *cfb
	wb     *stream
	rr     *recordReader
	closer io.Closer

	sheets   []boundSheet
	sheetIdx int

	sst      []string
	xfNumFmt []int
	formats  map[int]string
	dateByID map[int]bool
	date1904 bool

	sheetOpen bool

	// errorValues makes formula errors readable as their display text.
	errorValues bool

	// fetched holds a completed row that Read has not yet handed out, which
	// happens while blank rows are being emitted to close an index gap.
	fetched     []cells.Cell
	fetchedRow  int
	haveFetched bool

	// pending holds cells belonging to the row after the one being assembled.
	// A cell record announces its own row, so the first cell of the next row
	// is only recognisable once it has already been read.
	pending     []cells.Cell
	pendingRow  int
	havePending bool

	nextRow int
	row     []cells.Cell
	rowIdx  int

	err error
}

// Open parses the workbook globals of a BIFF8 file.
func Open(ra io.ReaderAt, size int64) (*Reader, error) {
	c, err := openCFB(ra, size)
	if err != nil {
		return nil, err
	}

	// Excel 97 and later name the stream "Workbook"; Excel 5 and 95 named it
	// "Book". We accept both here so that the version check below can report
	// the real problem rather than "no workbook stream".
	wb, err := c.openStream("Workbook", "Book")
	if err != nil {
		return nil, err
	}

	r := &Reader{
		cfb:      c,
		wb:       wb,
		rr:       newRecordReader(wb),
		sheetIdx: -1,
		formats:  map[int]string{},
		dateByID: map[int]bool{},
	}
	if err := r.readGlobals(); err != nil {
		return nil, err
	}
	return r, nil
}

// SetCloser registers a resource to release when the reader is closed.
func (r *Reader) SetCloser(cl io.Closer) { r.closer = cl }

// SetErrorValues controls whether formula error cells expose their display
// text as the cell value. It must be called before reading begins.
func (r *Reader) SetErrorValues(on bool) { r.errorValues = on }

// readGlobals consumes the workbook globals substream.
func (r *Reader) readGlobals() error {
	r.rr.seek(0)

	if err := r.rr.next(); err != nil {
		return fmt.Errorf("xls: reading first record: %w", err)
	}
	if r.rr.typ != recBOF {
		return errors.New("xls: workbook stream does not begin with a BOF record")
	}
	bof, err := parseBOF(r.rr.data)
	if err != nil {
		return err
	}
	if bof.version != biff8 {
		return fmt.Errorf("%w (found version 0x%04X)", ErrUnsupportedBIFF, bof.version)
	}
	if bof.substream != substreamWorkbook {
		return errors.New("xls: workbook stream does not begin with a globals substream")
	}

	for {
		if err := r.rr.next(); err != nil {
			if errors.Is(err, io.EOF) {
				break
			}
			return err
		}

		switch r.rr.typ {
		case recEOF:
			return nil

		case recBoundSheet:
			bs, err := parseBoundSheet(r.rr.data)
			if err != nil {
				return err
			}
			// Chart, macro and module sheets hold no cells; listing them would
			// produce phantom empty sheets in the iteration.
			if bs.kind == sheetKindWorksheet {
				r.sheets = append(r.sheets, bs)
			}

		case recSST:
			sst, err := parseSST(r.rr.data, r.rr.contStarts)
			if err != nil {
				return err
			}
			r.sst = sst

		case recXF:
			id, err := parseXF(r.rr.data)
			if err != nil {
				return err
			}
			r.xfNumFmt = append(r.xfNumFmt, id)

		case recFormat:
			id, code, err := parseFormat(r.rr.data)
			if err != nil {
				return err
			}
			r.formats[id] = code

		case recDateMode:
			// A single flag selecting the 1904 epoch, as used by workbooks
			// authored on classic Mac Excel.
			if len(r.rr.data) >= 2 {
				r.date1904 = r.rr.data[0]|r.rr.data[1] != 0
			}
		}
	}
	return nil
}

// SheetCount returns the number of worksheets.
func (r *Reader) SheetCount() int { return len(r.sheets) }

// SheetName returns the name of the current sheet.
func (r *Reader) SheetName() string {
	if r.sheetIdx < 0 || r.sheetIdx >= len(r.sheets) {
		return ""
	}
	return r.sheets[r.sheetIdx].name
}

// NextSheet advances to the next worksheet.
func (r *Reader) NextSheet() bool {
	r.sheetIdx++
	if r.sheetIdx >= len(r.sheets) {
		r.sheetOpen = false
		return false
	}

	r.row, r.rowIdx, r.nextRow = nil, -1, 0
	r.haveFetched, r.havePending = false, false
	r.sheetOpen = true

	// The BOUNDSHEET offset points straight at this sheet's BOF record, so no
	// scanning is needed to find it.
	r.rr.seek(r.sheets[r.sheetIdx].offset)
	if err := r.rr.next(); err != nil {
		r.noteErr(fmt.Errorf("xls: sheet %q: %w", r.SheetName(), err))
		r.sheetOpen = false
		return true
	}
	if r.rr.typ != recBOF {
		r.noteErr(fmt.Errorf("xls: sheet %q does not start with a BOF record", r.SheetName()))
		r.sheetOpen = false
	}
	return true
}

// Read advances to the next row of the current sheet.
//
// Rows the file never wrote are still reported as empty rows, so RowIndex
// always lines up with the row number Excel shows.
func (r *Reader) Read() bool {
	if !r.sheetOpen {
		return false
	}
	if !r.haveFetched {
		if !r.fetchRow() {
			return false
		}
	}

	if r.fetchedRow > r.nextRow {
		r.row = nil
		r.rowIdx = r.nextRow
		r.nextRow++
		return true
	}

	r.row = r.fetched
	r.rowIdx = r.fetchedRow
	r.nextRow = r.rowIdx + 1
	r.haveFetched = false
	return true
}

// fetchRow assembles the next populated row from the record stream.
func (r *Reader) fetchRow() bool {
	var (
		cur    []cells.Cell
		curRow = -1
	)
	if r.havePending {
		cur, curRow = r.pending, r.pendingRow
		r.pending, r.havePending = nil, false
	}

	for {
		if err := r.rr.next(); err != nil {
			if !errors.Is(err, io.EOF) {
				r.noteErr(err)
			}
			r.sheetOpen = false
			return r.finishRow(cur, curRow)
		}

		if r.rr.typ == recEOF {
			r.sheetOpen = false
			return r.finishRow(cur, curRow)
		}

		row, placed, ok := r.decodeCellRecord()
		if !ok {
			continue
		}

		if curRow == -1 {
			curRow = int(row)
		}
		if int(row) != curRow {
			// The record belongs to a later row, so the row being assembled is
			// complete. Hold the new cells until the next call.
			r.pending, r.pendingRow, r.havePending = placed, int(row), true
			return r.finishRow(cur, curRow)
		}
		cur = mergeCells(cur, placed)
	}
}

func (r *Reader) finishRow(cur []cells.Cell, curRow int) bool {
	if curRow < 0 {
		return false
	}
	r.fetched, r.fetchedRow, r.haveFetched = cur, curRow, true
	return true
}

// mergeCells overlays src onto dst, growing dst as needed.
func mergeCells(dst, src []cells.Cell) []cells.Cell {
	for len(dst) < len(src) {
		dst = append(dst, cells.Cell{})
	}
	for i, c := range src {
		if c.Value != nil || c.Raw != "" {
			dst[i] = c
		}
	}
	return dst
}

// placeAt returns a slice holding cell at column col, padded with empties.
func placeAt(col int, cs ...cells.Cell) []cells.Cell {
	out := make([]cells.Cell, col+len(cs))
	copy(out[col:], cs)
	return out
}

// decodeCellRecord turns the current record into cells, if it is a cell record.
//
// It returns the row the cells belong to and a slice positioned at their
// columns. Records that are not cell data return ok=false and are skipped.
func (r *Reader) decodeCellRecord() (row uint16, placed []cells.Cell, ok bool) {
	switch r.rr.typ {
	case recBlank:
		c := newCursor(r.rr.data, nil)
		h, good := parseCellHeader(c)
		if !good {
			return 0, nil, false
		}
		// A BLANK is a formatted but valueless cell. It still occupies its
		// column, so it widens the row.
		return h.row, placeAt(int(h.col), cells.Cell{}), true

	case recMulBlank:
		c := newCursor(r.rr.data, nil)
		rw, ok1 := c.u16()
		colFirst, ok2 := c.u16()
		if !ok1 || !ok2 {
			return 0, nil, false
		}
		n := (c.remaining() - 2) / 2
		if n < 0 {
			return 0, nil, false
		}
		out := make([]cells.Cell, n)
		return rw, placeAt(int(colFirst), out...), true

	case recNumber:
		c := newCursor(r.rr.data, nil)
		h, good := parseCellHeader(c)
		if !good {
			return 0, nil, false
		}
		v, good := c.f64()
		if !good {
			return 0, nil, false
		}
		return h.row, placeAt(int(h.col), r.numericCell(h.xf, v)), true

	case recRK:
		c := newCursor(r.rr.data, nil)
		h, good := parseCellHeader(c)
		if !good {
			return 0, nil, false
		}
		raw, good := c.u32()
		if !good {
			return 0, nil, false
		}
		return h.row, placeAt(int(h.col), r.numericCell(h.xf, decodeRK(raw))), true

	case recMulRK:
		c := newCursor(r.rr.data, nil)
		rw, ok1 := c.u16()
		colFirst, ok2 := c.u16()
		if !ok1 || !ok2 {
			return 0, nil, false
		}
		// Each entry is a format index plus a packed value; the trailing two
		// bytes are the last column, which we can derive instead.
		n := (c.remaining() - 2) / 6
		out := make([]cells.Cell, 0, maxInt(n, 0))
		for i := 0; i < n; i++ {
			xf, ok1 := c.u16()
			raw, ok2 := c.u32()
			if !ok1 || !ok2 {
				break
			}
			out = append(out, r.numericCell(xf, decodeRK(raw)))
		}
		return rw, placeAt(int(colFirst), out...), true

	case recLabel, recRString:
		c := newCursor(r.rr.data, nil)
		h, good := parseCellHeader(c)
		if !good {
			return 0, nil, false
		}
		s, err := readLongString(c)
		if err != nil {
			r.noteErr(err)
			return 0, nil, false
		}
		return h.row, placeAt(int(h.col), cells.Cell{Value: s, Raw: s}), true

	case recLabelSST:
		c := newCursor(r.rr.data, nil)
		h, good := parseCellHeader(c)
		if !good {
			return 0, nil, false
		}
		idx, good := c.u32()
		if !good {
			return 0, nil, false
		}
		if int(idx) >= len(r.sst) {
			r.noteErr(fmt.Errorf("xls: shared string index %d out of range (table holds %d)", idx, len(r.sst)))
			return h.row, placeAt(int(h.col), cells.Cell{}), true
		}
		s := r.sst[idx]
		return h.row, placeAt(int(h.col), cells.Cell{Value: s, Raw: s}), true

	case recBoolErr:
		c := newCursor(r.rr.data, nil)
		h, good := parseCellHeader(c)
		if !good {
			return 0, nil, false
		}
		val, ok1 := c.u8()
		isErr, ok2 := c.u8()
		if !ok1 || !ok2 {
			return 0, nil, false
		}
		if isErr != 0 {
			return h.row, placeAt(int(h.col), r.errorCell(errorText(val))), true
		}
		b := val != 0
		return h.row, placeAt(int(h.col), cells.Cell{Value: b, Raw: strconv.FormatBool(b)}), true

	case recFormula:
		return r.decodeFormula()

	default:
		return 0, nil, false
	}
}

// decodeFormula reads a FORMULA record and, when its cached result is a string,
// the STRING record that carries it.
func (r *Reader) decodeFormula() (uint16, []cells.Cell, bool) {
	c := newCursor(r.rr.data, nil)
	h, good := parseCellHeader(c)
	if !good {
		return 0, nil, false
	}
	if c.remaining() < 8 {
		return 0, nil, false
	}
	result := c.data[c.pos : c.pos+8]
	value, needsString := formulaResult(result)

	if needsString {
		s, err := r.readFollowingString()
		if err != nil {
			r.noteErr(err)
			return h.row, placeAt(int(h.col), cells.Cell{}), true
		}
		return h.row, placeAt(int(h.col), cells.Cell{Value: s, Raw: s}), true
	}

	switch v := value.(type) {
	case float64:
		return h.row, placeAt(int(h.col), r.numericCell(h.xf, v)), true
	case errValue:
		return h.row, placeAt(int(h.col), r.errorCell(string(v))), true
	case bool:
		return h.row, placeAt(int(h.col), cells.Cell{Value: v, Raw: strconv.FormatBool(v)}), true
	case string:
		return h.row, placeAt(int(h.col), cells.Cell{Value: v, Raw: v}), true
	default:
		return h.row, placeAt(int(h.col), cells.Cell{}), true
	}
}

// errorCell builds a cell for a formula error such as #DIV/0!.
func (r *Reader) errorCell(text string) cells.Cell {
	if r.errorValues {
		return cells.Cell{Value: text, Raw: text}
	}
	return cells.Cell{Raw: text}
}

// readFollowingString consumes the STRING record that carries a formula's
// cached text result.
//
// Shared-formula and array bookkeeping records may sit between the FORMULA and
// its STRING, so a few record types are stepped over. Anything else means the
// result is absent and the cell is left empty rather than desynchronising the
// record stream.
func (r *Reader) readFollowingString() (string, error) {
	const (
		recShrFmla uint16 = 0x04BC
		recArray   uint16 = 0x0221
	)
	for {
		typ, err := r.rr.peek()
		if err != nil {
			return "", err
		}
		switch typ {
		case recShrFmla, recArray:
			if err := r.rr.next(); err != nil {
				return "", err
			}
		case recString:
			if err := r.rr.next(); err != nil {
				return "", err
			}
			return readLongString(newCursor(r.rr.data, r.rr.contStarts))
		default:
			return "", nil
		}
	}
}

// numericCell types a stored double using the cell's number format.
func (r *Reader) numericCell(xf uint16, v float64) cells.Cell {
	if r.isDateXF(int(xf)) {
		return cells.Cell{Value: numfmt.SerialToTime(v, r.date1904), Raw: formatFloat(v)}
	}
	// A whole number is returned as int64 so that a column of counts reads the
	// same here as it does from an xlsx, where the file's own text distinguishes
	// integers from decimals. BIFF has no such text, so the value itself decides.
	if v == math.Trunc(v) && !math.IsInf(v, 0) && math.Abs(v) < 1<<53 {
		n := int64(v)
		return cells.Cell{Value: n, Raw: strconv.FormatInt(n, 10)}
	}
	return cells.Cell{Value: v, Raw: formatFloat(v)}
}

func formatFloat(v float64) string { return strconv.FormatFloat(v, 'g', -1, 64) }

// isDateXF resolves a cell's format index to whether it denotes a date.
func (r *Reader) isDateXF(xf int) bool {
	if xf < 0 || xf >= len(r.xfNumFmt) {
		return false
	}
	id := r.xfNumFmt[xf]
	if cached, ok := r.dateByID[id]; ok {
		return cached
	}

	var isDate bool
	if code, ok := r.formats[id]; ok {
		isDate = numfmt.IsDateFormat(code)
	} else if numfmt.IsDateBuiltin(id) {
		isDate = true
	} else if code, ok := numfmt.Builtin(id); ok {
		isDate = numfmt.IsDateFormat(code)
	}

	r.dateByID[id] = isDate
	return isDate
}

// Row returns the cells of the current row.
func (r *Reader) Row() []cells.Cell { return r.row }

// RowIndex returns the zero-based index of the current row.
func (r *Reader) RowIndex() int { return r.rowIdx }

// Err returns the first error encountered while streaming.
func (r *Reader) Err() error { return r.err }

func (r *Reader) noteErr(err error) {
	if r.err == nil {
		r.err = err
	}
}

// Close releases the reader's resources.
func (r *Reader) Close() error {
	if r.closer != nil {
		err := r.closer.Close()
		r.closer = nil
		return err
	}
	return nil
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}

var _ cells.Source = (*Reader)(nil)
