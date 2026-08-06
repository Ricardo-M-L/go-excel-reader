package xlsx

import (
	"encoding/xml"
	"fmt"
	"io"
	"strconv"

	"github.com/jloor/go-excel-reader/numfmt"
)

// Styles resolves a cell's style index to the question the reader actually
// needs answered: is this number a date?
//
// The path is indirect by design in the file format. A cell carries s="3",
// which indexes cellXfs; that xf carries numFmtId="164", which indexes either
// the built-in format table or the numFmts the file defines itself; and only
// that format string reveals whether the cell's float64 is a quantity or a
// moment in time. Every hop is a chance to get it wrong, and getting it wrong
// is how a reader turns a delivery date into the number 45217.
type Styles struct {
	// xfNumFmt maps a cellXfs index to a numFmtId.
	xfNumFmt []int
	// numFmts holds the format codes this file defines (ids 164 and up).
	numFmts map[int]string
	// dateByFmtID caches the date test per numFmtId. A sheet asks this question
	// once per cell, and parsing a format string per cell would dominate the
	// cost of reading a numeric column.
	dateByFmtID map[int]bool
}

// numFmtID returns the number format id for a cell style index.
func (s *Styles) numFmtID(xfIndex int) (int, bool) {
	if s == nil || xfIndex < 0 || xfIndex >= len(s.xfNumFmt) {
		return 0, false
	}
	return s.xfNumFmt[xfIndex], true
}

// IsDate reports whether cells carrying the given style index hold a date.
func (s *Styles) IsDate(xfIndex int) bool {
	id, ok := s.numFmtID(xfIndex)
	if !ok {
		return false
	}
	if cached, ok := s.dateByFmtID[id]; ok {
		return cached
	}

	var isDate bool
	if custom, ok := s.numFmts[id]; ok {
		// A file-defined format always wins: a file may legally redefine a
		// built-in id, and the string it supplies is what Excel displays.
		isDate = numfmt.IsDateFormat(custom)
	} else if numfmt.IsDateBuiltin(id) {
		isDate = true
	} else if builtin, ok := numfmt.Builtin(id); ok {
		isDate = numfmt.IsDateFormat(builtin)
	}

	s.dateByFmtID[id] = isDate
	return isDate
}

// NumFmt returns the number format code for a cell style index, or "" if the
// style or format is unknown.
func (s *Styles) NumFmt(xfIndex int) string {
	id, ok := s.numFmtID(xfIndex)
	if !ok {
		return ""
	}
	if custom, ok := s.numFmts[id]; ok {
		return custom
	}
	if builtin, ok := numfmt.Builtin(id); ok {
		return builtin
	}
	return ""
}

// newStyles returns an empty style table, used when a workbook has no
// styles.xml. Such a file has no date formats, so every number stays a number.
func newStyles() *Styles {
	return &Styles{
		numFmts:     map[int]string{},
		dateByFmtID: map[int]bool{},
	}
}

// parseStyles streams xl/styles.xml.
func parseStyles(r io.Reader) (*Styles, error) {
	dec := newDecoder(r)
	st := newStyles()

	// styles.xml contains two <xf> lists. cellStyleXfs holds named-style
	// definitions and cellXfs holds the per-cell records that a cell's s=
	// attribute actually indexes. They look identical and appear in that
	// order, so reading <xf> elements without tracking which list we are in
	// silently binds every cell to the wrong format.
	inCellXfs := false

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("styles.xml: %w", err)
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "cellXfs":
				inCellXfs = true
			case "numFmt":
				id, code := -1, ""
				for _, a := range t.Attr {
					switch a.Name.Local {
					case "numFmtId":
						if n, err := strconv.Atoi(a.Value); err == nil {
							id = n
						}
					case "formatCode":
						code = a.Value
					}
				}
				if id >= 0 {
					st.numFmts[id] = code
				}
			case "xf":
				if !inCellXfs {
					continue
				}
				id := 0
				for _, a := range t.Attr {
					if a.Name.Local == "numFmtId" {
						if n, err := strconv.Atoi(a.Value); err == nil {
							id = n
						}
					}
				}
				st.xfNumFmt = append(st.xfNumFmt, id)
			}
		case xml.EndElement:
			if t.Name.Local == "cellXfs" {
				inCellXfs = false
			}
		}
	}

	return st, nil
}
