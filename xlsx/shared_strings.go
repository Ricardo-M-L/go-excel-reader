package xlsx

import (
	"encoding/xml"
	"fmt"
	"io"
)

// SharedStrings is the workbook's string table.
//
// Almost every string in an xlsx lives here, not in the sheet. A sheet cell of
// type "s" stores an index into this table, so a spreadsheet with a million
// repeated country names stores each distinct name once. That is good for file
// size and bad for a naive reader, because the obvious representation -- a
// []string -- is the single largest allocation a reader makes on a big file.
//
// We store the table as one contiguous byte blob plus an offset index. For a
// table of n strings that replaces n separate heap allocations and n pointers
// with two allocations and zero pointers. The size win is real (a []string
// header costs 16 bytes per entry before the bytes themselves) but the garbage
// collector win is larger: the GC never has to trace into a []byte, so a
// 200,000-entry table stops being 200,000 objects it must scan on every cycle.
//
// Strings are returned by slicing the blob, which is safe because the blob is
// never mutated after parsing and Go strings are immutable.
type SharedStrings struct {
	blob []byte
	// offs holds n+1 entries; string i occupies blob[offs[i]:offs[i+1]].
	offs []int64
}

// Len returns the number of strings in the table.
func (s *SharedStrings) Len() int {
	if s == nil || len(s.offs) == 0 {
		return 0
	}
	return len(s.offs) - 1
}

// At returns the string at index i. The bool reports whether i was in range;
// an out-of-range index means the sheet references a string the table does not
// contain, which is a corrupt file rather than an empty cell.
func (s *SharedStrings) At(i int) (string, bool) {
	if s == nil || i < 0 || i >= s.Len() {
		return "", false
	}
	return string(s.blob[s.offs[i]:s.offs[i+1]]), true
}

// parseSharedStrings streams sharedStrings.xml into a SharedStrings table.
//
// The document is never materialised as a tree. We walk tokens and append text
// directly into the blob, so peak memory is the size of the strings themselves
// plus the growth slack of one slice, not a multiple of the file size.
func parseSharedStrings(r io.Reader) (*SharedStrings, error) {
	dec := newDecoder(r)
	ss := &SharedStrings{offs: []int64{0}}

	// depth tracking for the one genuinely awkward part of this schema:
	// a <si> may contain <t> directly, or a sequence of <r> runs each with
	// their own <t>, and it may also contain <rPh> phonetic hints which
	// themselves contain <t>. Concatenating every <t> we see would silently
	// append Japanese furigana to the middle of the string.
	var (
		inSI  bool
		inT   bool
		phDep int
	)

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("sharedStrings.xml: %w", err)
		}

		switch t := tok.(type) {
		case xml.StartElement:
			switch t.Name.Local {
			case "si":
				inSI = true
			case "rPh", "phoneticPr":
				if inSI {
					phDep++
				}
			case "t":
				if inSI && phDep == 0 {
					inT = true
				}
			}
		case xml.CharData:
			if inT {
				ss.blob = append(ss.blob, t...)
			}
		case xml.EndElement:
			switch t.Name.Local {
			case "si":
				if inSI {
					inSI = false
					ss.offs = append(ss.offs, int64(len(ss.blob)))
				}
			case "rPh", "phoneticPr":
				if phDep > 0 {
					phDep--
				}
			case "t":
				inT = false
			}
		}
	}

	return ss, nil
}
