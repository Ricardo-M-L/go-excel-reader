package xlsx

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
	"strings"

	"github.com/jloor/go-excel-reader/internal/cells"
	"github.com/jloor/go-excel-reader/numfmt"
)

// newDecoder builds an XML decoder for a package part.
//
// OPC requires parts to be UTF-8 or UTF-16, so rather than pulling in a charset
// registry we accept the encodings the spec permits and reject anything else
// with a clear message. Silently misdecoding text is worse than refusing it.
func newDecoder(r io.Reader) *xml.Decoder {
	d := xml.NewDecoder(r)
	d.CharsetReader = func(charset string, input io.Reader) (io.Reader, error) {
		switch strings.ToLower(charset) {
		case "", "utf-8", "utf8", "us-ascii", "ascii":
			return input, nil
		}
		return nil, fmt.Errorf("unsupported XML charset %q", charset)
	}
	return d
}

// columnIndex converts the letter part of a cell reference to a zero-based
// column index: A is 0, Z is 25, AA is 26.
//
// The letters are bijective base-26, not ordinary base-26: there is no zero
// digit, so A..Z are 1..26 and AA is 26+1. Treating 'A' as 0 makes every
// two-letter column off by 26, which is the classic way a reader lands every
// value past column Z in the wrong place.
func columnIndex(ref string) (int, bool) {
	col := 0
	i := 0
	for ; i < len(ref); i++ {
		c := ref[i]
		if c >= 'a' && c <= 'z' {
			c -= 'a' - 'A'
		}
		if c < 'A' || c > 'Z' {
			break
		}
		col = col*26 + int(c-'A') + 1
	}
	if i == 0 {
		return 0, false
	}
	return col - 1, true
}

// sheetScanner streams one worksheet part, yielding a row at a time.
//
// The sheet is the only part that can be arbitrarily large, so it is never held
// in memory: the decoder pulls from the zip's decompressing reader and each row
// overwrites the buffer used by the previous one.
type sheetScanner struct {
	rc  io.Closer
	dec *xml.Decoder
	own *Reader

	row    []cells.Cell
	rowIdx int // zero-based index of the row in buf
	// nextImplicitRow tracks where an unnumbered <row> lands. The r attribute
	// is optional, and files written by streaming producers frequently omit it.
	nextImplicitRow int

	err  error
	done bool
}

// next advances to the following populated row in the sheet part. It reports
// false at end of sheet or on error.
func (s *sheetScanner) next() bool {
	if s.done || s.err != nil {
		return false
	}
	for {
		tok, err := s.dec.Token()
		if err == io.EOF {
			s.done = true
			return false
		}
		if err != nil {
			s.err = err
			return false
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "row" {
			continue
		}
		if err := s.readRow(start); err != nil {
			s.err = err
			return false
		}
		return true
	}
}

func (s *sheetScanner) readRow(start xml.StartElement) error {
	rowNum := 0
	for _, a := range start.Attr {
		if a.Name.Local == "r" {
			if n, err := strconv.Atoi(a.Value); err == nil {
				rowNum = n
			}
		}
	}
	if rowNum > 0 {
		s.rowIdx = rowNum - 1
	} else {
		s.rowIdx = s.nextImplicitRow
	}
	s.nextImplicitRow = s.rowIdx + 1

	s.row = s.row[:0]
	prevCol := -1

	for {
		tok, err := s.dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return err
		}

		switch t := tok.(type) {
		case xml.StartElement:
			if t.Name.Local != "c" {
				if err := s.dec.Skip(); err != nil {
					return err
				}
				continue
			}
			col, cell, err := s.readCell(t, prevCol)
			if err != nil {
				return err
			}
			prevCol = col
			// Grow to the cell's column, leaving gaps as empty cells. A sheet
			// omits cells that were never touched, so column position comes
			// from the r attribute and not from arrival order.
			for len(s.row) < col {
				s.row = append(s.row, cells.Cell{})
			}
			if len(s.row) == col {
				s.row = append(s.row, cell)
			} else {
				s.row[col] = cell
			}
		case xml.EndElement:
			if t.Name.Local == "row" {
				return nil
			}
		}
	}
}

// readCell parses one <c> element and returns its column index and value.
func (s *sheetScanner) readCell(start xml.StartElement, prevCol int) (int, cells.Cell, error) {
	var (
		ref      string
		typ      string
		styleIdx = -1
	)
	for _, a := range start.Attr {
		switch a.Name.Local {
		case "r":
			ref = a.Value
		case "t":
			typ = a.Value
		case "s":
			if n, err := strconv.Atoi(a.Value); err == nil {
				styleIdx = n
			}
		}
	}

	col := prevCol + 1
	if ref != "" {
		if c, ok := columnIndex(ref); ok {
			col = c
		}
	}

	var (
		raw    string
		hasVal bool
	)
	for {
		tok, err := s.dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return col, cells.Cell{}, err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "v":
				text, err := collectText(s.dec, false)
				if err != nil {
					return col, cells.Cell{}, err
				}
				raw, hasVal = text, true
			case "is":
				// An inline string carries its text in the cell itself rather
				// than in the shared table, optionally split across <r> runs.
				text, err := collectText(s.dec, true)
				if err != nil {
					return col, cells.Cell{}, err
				}
				raw, hasVal = text, true
			default:
				// <f> holds the formula source, which we deliberately ignore in
				// favour of the cached <v> result that follows it.
				if err := s.dec.Skip(); err != nil {
					return col, cells.Cell{}, err
				}
			}
		case xml.EndElement:
			if t.Name.Local == "c" {
				if !hasVal {
					return col, cells.Cell{}, nil
				}
				return col, s.own.makeCell(raw, typ, styleIdx), nil
			}
		}
	}
	if !hasVal {
		return col, cells.Cell{}, nil
	}
	return col, s.own.makeCell(raw, typ, styleIdx), nil
}

// collectText accumulates character data until the current element closes.
//
// With skipPhonetic set it ignores <rPh> subtrees, which carry furigana
// pronunciation hints for East Asian text. Those hints contain <t> elements of
// their own, so a reader that concatenates every <t> it encounters splices
// pronunciation guides into the middle of the value.
func collectText(dec *xml.Decoder, skipPhonetic bool) (string, error) {
	var (
		sb    strings.Builder
		depth = 1
	)
	for depth > 0 {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return "", err
		}
		switch t := tok.(type) {
		case xml.StartElement:
			if skipPhonetic && (t.Name.Local == "rPh" || t.Name.Local == "phoneticPr") {
				if err := dec.Skip(); err != nil {
					return "", err
				}
				continue
			}
			depth++
		case xml.EndElement:
			depth--
		case xml.CharData:
			sb.Write(t)
		}
	}
	return sb.String(), nil
}

// makeCell applies the cell's declared type and number format to produce a
// Go-typed value.
//
// The mapping is where a spreadsheet reader earns its keep. A cell's t
// attribute distinguishes strings from numbers from booleans, but it says
// nothing about dates: a date is a plain number whose style happens to carry a
// date format. That second lookup is the difference between returning
// time.Time and returning 45217.
func (r *Reader) makeCell(raw, typ string, styleIdx int) cells.Cell {
	switch typ {
	case "s":
		// Shared string: the value is an index into the workbook string table.
		idx, err := strconv.Atoi(strings.TrimSpace(raw))
		if err != nil {
			return cells.Cell{Value: raw, Raw: raw}
		}
		sst := r.sharedStrings()
		str, ok := sst.At(idx)
		if !ok {
			// A dangling index means a corrupt table. Surface it rather than
			// pretending the cell was empty.
			r.noteErr(fmt.Errorf("xlsx: shared string index %d out of range (table holds %d)", idx, sst.Len()))
			return cells.Cell{Raw: raw}
		}
		// An empty entry in the shared table is a value: the file went to the
		// trouble of interning it. Only inline text is treated as absent below.
		return cells.Cell{Value: str, Raw: str}

	case "inlineStr":
		// Text stored in the cell itself. An empty <t> here means the cell was
		// written but holds nothing, so it reads as empty rather than as "".
		//
		// This asymmetry with the shared table is not ours; it is what
		// ExcelDataReader does, and matching it is what the corpus comparison
		// is for. A whole-sheet export can otherwise differ by hundreds of
		// cells purely in how blanks are represented.
		return textCell(raw)

	case "str":
		// A formula's cached string result, which may legitimately be empty.
		return cells.Cell{Value: raw, Raw: raw}

	case "e":
		// An error such as #DIV/0!. It has no Go equivalent, so by default the
		// cell reads as empty and the text stays in Raw, which matches
		// ExcelDataReader. WithErrorValues surfaces the text as the value.
		return r.errorCell(raw)

	case "b":
		return cells.Cell{Value: raw == "1" || strings.EqualFold(raw, "true"), Raw: raw}

	case "d":
		// ISO 8601 dates, permitted since ECMA-376 2nd edition and written by
		// some non-Excel producers.
		if t, err := parseISODate(raw); err == nil {
			return cells.Cell{Value: t, Raw: raw}
		}
		return cells.Cell{Value: raw, Raw: raw}

	default: // "n" or absent
		if strings.TrimSpace(raw) == "" {
			return cells.Cell{Raw: raw}
		}
		f, err := strconv.ParseFloat(raw, 64)
		if err != nil {
			return cells.Cell{Value: raw, Raw: raw}
		}
		if styleIdx >= 0 && r.styles.IsDate(styleIdx) {
			return cells.Cell{Value: numfmt.SerialToTime(f, r.date1904), Raw: raw}
		}
		// Prefer int64 when the file wrote a whole number. Excel stores every
		// number as a double, but a caller reading a row of counts should not
		// have to convert each one, and the file's own text tells us whether a
		// fractional part was intended.
		if isIntegerLiteral(raw) {
			if n, err := strconv.ParseInt(raw, 10, 64); err == nil {
				return cells.Cell{Value: n, Raw: raw}
			}
		}
		return cells.Cell{Value: f, Raw: raw}
	}
}

// textCell builds a cell holding text.
//
// A cell whose text is empty reads as empty rather than as the empty string.
// A file may write a string cell with no content, and reporting that as a
// present-but-empty value makes it indistinguishable from a deliberate "" while
// making IsEmpty answer false for a cell the user sees as blank. The raw text is
// preserved either way.
func textCell(s string) cells.Cell {
	if s == "" {
		return cells.Cell{}
	}
	return cells.Cell{Value: s, Raw: s}
}

// errorCell builds a cell for a formula error such as #DIV/0!.
func (r *Reader) errorCell(text string) cells.Cell {
	if r.errorValues {
		return cells.Cell{Value: text, Raw: text}
	}
	return cells.Cell{Raw: text}
}

// isIntegerLiteral reports whether the raw text denotes a whole number written
// without a fractional part or exponent.
func isIntegerLiteral(raw string) bool {
	if raw == "" {
		return false
	}
	for i := 0; i < len(raw); i++ {
		switch c := raw[i]; {
		case c >= '0' && c <= '9':
		case (c == '-' || c == '+') && i == 0:
		default:
			return false
		}
	}
	return true
}
