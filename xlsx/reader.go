// Package xlsx reads ECMA-376 OpenXML spreadsheets (.xlsx).
//
// An xlsx file is a ZIP archive of XML parts. The reader streams the parts it
// cannot bound in size -- the worksheets and the shared string table -- so that
// peak memory tracks the widest row rather than the size of the file.
package xlsx

import (
	"archive/zip"
	"encoding/xml"
	"errors"
	"fmt"
	"io"
	"path"
	"strings"
	"time"

	"github.com/jloor/go-excel-reader/internal/cells"
)

// ErrNotXLSX reports an archive that is a valid ZIP but not a spreadsheet.
var ErrNotXLSX = errors.New("xlsx: missing xl/workbook.xml; not an OpenXML spreadsheet")

// sheetRef is one entry of the workbook's sheet list.
type sheetRef struct {
	name  string
	relID string
	path  string
}

// Reader streams an xlsx workbook one row at a time.
type Reader struct {
	zr     *zip.Reader
	parts  map[string]*zip.File
	closer io.Closer

	sheets   []sheetRef
	sheetIdx int
	date1904 bool

	styles *Styles

	sst       *SharedStrings
	sstLoaded bool

	scan *sheetScanner
	// pending marks a row that the scanner has produced but that Read has not
	// yet delivered, because blank rows are being emitted ahead of it.
	pending bool
	nextRow int

	row    []cells.Cell
	rowIdx int

	err error
}

// Open reads the workbook structure from a ZIP archive.
//
// The caller keeps ownership of ra. Nothing beyond the workbook metadata, the
// style table and, on demand, the shared strings is read up front; sheet
// contents stream as they are iterated.
func Open(ra io.ReaderAt, size int64) (*Reader, error) {
	zr, err := zip.NewReader(ra, size)
	if err != nil {
		return nil, fmt.Errorf("xlsx: %w", err)
	}

	r := &Reader{
		zr:       zr,
		parts:    make(map[string]*zip.File, len(zr.File)),
		sheetIdx: -1,
	}
	for _, f := range zr.File {
		r.parts[normalizePart(f.Name)] = f
	}

	wb := r.part("xl/workbook.xml")
	if wb == nil {
		return nil, ErrNotXLSX
	}
	rc, err := wb.Open()
	if err != nil {
		return nil, fmt.Errorf("xlsx: opening workbook.xml: %w", err)
	}
	sheets, date1904, err := parseWorkbook(rc)
	rc.Close()
	if err != nil {
		return nil, err
	}
	r.date1904 = date1904

	rels, err := r.loadRels()
	if err != nil {
		return nil, err
	}
	r.sheets = r.resolveSheetPaths(sheets, rels)

	if err := r.loadStyles(); err != nil {
		return nil, err
	}
	return r, nil
}

// SetCloser registers a resource to release when the reader is closed. It lets
// the caller hand over an open file without the reader having to know how the
// bytes were obtained.
func (r *Reader) SetCloser(c io.Closer) { r.closer = c }

// normalizePart canonicalises a zip entry name for lookup. OPC part names are
// case-sensitive in principle, but producers disagree in practice, so lookups
// go through a lowered, slash-normalised key.
func normalizePart(name string) string {
	return strings.ToLower(strings.TrimPrefix(path.Clean("/"+name), "/"))
}

func (r *Reader) part(name string) *zip.File { return r.parts[normalizePart(name)] }

func (r *Reader) loadStyles() error {
	f := r.part("xl/styles.xml")
	if f == nil {
		// A workbook without styles has no date formats, so every number is
		// simply a number.
		r.styles = newStyles()
		return nil
	}
	rc, err := f.Open()
	if err != nil {
		return fmt.Errorf("xlsx: opening styles.xml: %w", err)
	}
	defer rc.Close()

	st, err := parseStyles(rc)
	if err != nil {
		return err
	}
	r.styles = st
	return nil
}

func (r *Reader) loadRels() (map[string]string, error) {
	f := r.part("xl/_rels/workbook.xml.rels")
	if f == nil {
		return nil, nil
	}
	rc, err := f.Open()
	if err != nil {
		return nil, fmt.Errorf("xlsx: opening workbook.xml.rels: %w", err)
	}
	defer rc.Close()
	return parseRels(rc)
}

// sharedStrings loads the string table on first use.
//
// Loading is deferred because a workbook's string table can be its largest
// part, and a caller that only reads numeric columns should never pay for it.
func (r *Reader) sharedStrings() *SharedStrings {
	if r.sstLoaded {
		return r.sst
	}
	r.sstLoaded = true
	r.sst = &SharedStrings{}

	f := r.part("xl/sharedStrings.xml")
	if f == nil {
		return r.sst
	}
	rc, err := f.Open()
	if err != nil {
		r.noteErr(fmt.Errorf("xlsx: opening sharedStrings.xml: %w", err))
		return r.sst
	}
	defer rc.Close()

	ss, err := parseSharedStrings(rc)
	if err != nil {
		r.noteErr(err)
		return r.sst
	}
	r.sst = ss
	return r.sst
}

// resolveSheetPaths turns relationship ids into part names.
func (r *Reader) resolveSheetPaths(sheets []sheetRef, rels map[string]string) []sheetRef {
	out := make([]sheetRef, 0, len(sheets))
	fallback := 0
	for _, s := range sheets {
		if target, ok := rels[s.relID]; ok && target != "" {
			s.path = resolveTarget(target)
		} else {
			// Some producers omit the relationship. Fall back to the
			// conventional layout rather than dropping the sheet.
			fallback++
			s.path = fmt.Sprintf("xl/worksheets/sheet%d.xml", fallback)
		}
		out = append(out, s)
	}
	return out
}

// resolveTarget converts a relationship target to a package part name. Targets
// are relative to the part that declared them (xl/workbook.xml), except when
// they are absolute.
func resolveTarget(target string) string {
	if strings.HasPrefix(target, "/") {
		return strings.TrimPrefix(target, "/")
	}
	return path.Join("xl", target)
}

// SheetCount returns the number of sheets in the workbook.
func (r *Reader) SheetCount() int { return len(r.sheets) }

// SheetName returns the name of the current sheet.
func (r *Reader) SheetName() string {
	if r.sheetIdx < 0 || r.sheetIdx >= len(r.sheets) {
		return ""
	}
	return r.sheets[r.sheetIdx].name
}

// NextSheet advances to the next worksheet, reporting false when none remain.
func (r *Reader) NextSheet() bool {
	r.closeScanner()
	r.sheetIdx++
	if r.sheetIdx >= len(r.sheets) {
		return false
	}

	r.row, r.rowIdx, r.nextRow, r.pending = nil, -1, 0, false

	f := r.part(r.sheets[r.sheetIdx].path)
	if f == nil {
		r.noteErr(fmt.Errorf("xlsx: sheet %q references missing part %q",
			r.sheets[r.sheetIdx].name, r.sheets[r.sheetIdx].path))
		return true // the sheet exists in the workbook even if its part does not
	}
	rc, err := f.Open()
	if err != nil {
		r.noteErr(fmt.Errorf("xlsx: opening %s: %w", r.sheets[r.sheetIdx].path, err))
		return true
	}
	r.scan = &sheetScanner{rc: rc, dec: newDecoder(rc), own: r}
	return true
}

func (r *Reader) closeScanner() {
	if r.scan == nil {
		return
	}
	if r.scan.err != nil {
		r.noteErr(r.scan.err)
	}
	if r.scan.rc != nil {
		r.scan.rc.Close()
	}
	r.scan = nil
}

// Read advances to the next row of the current sheet.
//
// Rows omitted from the file are still reported, as empty rows, so that
// RowIndex always matches the row number a user sees in Excel. A sheet whose
// first value sits in row 5 therefore yields four empty rows first; without
// that, every row index below a gap would be wrong.
func (r *Reader) Read() bool {
	if r.scan == nil {
		return false
	}
	if !r.pending {
		if !r.scan.next() {
			if r.scan.err != nil {
				r.noteErr(r.scan.err)
			}
			return false
		}
		r.pending = true
	}

	if r.scan.rowIdx > r.nextRow {
		r.row = nil
		r.rowIdx = r.nextRow
		r.nextRow++
		return true
	}

	r.row = r.scan.row
	r.rowIdx = r.scan.rowIdx
	r.nextRow = r.rowIdx + 1
	r.pending = false
	return true
}

// Row returns the cells of the current row. The slice is reused by the next
// call to Read.
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
	r.closeScanner()
	if r.closer != nil {
		err := r.closer.Close()
		r.closer = nil
		if err != nil {
			return err
		}
	}
	return nil
}

// parseWorkbook reads the sheet list and the workbook's date system.
func parseWorkbook(rd io.Reader) ([]sheetRef, bool, error) {
	dec := newDecoder(rd)
	var (
		sheets   []sheetRef
		date1904 bool
	)

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, false, fmt.Errorf("xlsx: workbook.xml: %w", err)
		}
		start, ok := tok.(xml.StartElement)
		if !ok {
			continue
		}

		switch start.Name.Local {
		case "workbookPr":
			for _, a := range start.Attr {
				// Excel writes date1904; some producers write the older
				// spelling. Both select the 1904 epoch.
				if a.Name.Local == "date1904" || a.Name.Local == "dateCompatibility" {
					date1904 = a.Value == "1" || strings.EqualFold(a.Value, "true")
				}
			}
		case "sheet":
			var s sheetRef
			for _, a := range start.Attr {
				switch {
				case a.Name.Local == "name":
					s.name = a.Value
				case a.Name.Local == "id" && strings.Contains(a.Name.Space, "relationships"):
					s.relID = a.Value
				}
			}
			sheets = append(sheets, s)
		}
	}
	return sheets, date1904, nil
}

// parseRels reads a relationship part into an id-to-target map.
func parseRels(rd io.Reader) (map[string]string, error) {
	dec := newDecoder(rd)
	rels := map[string]string{}

	for {
		tok, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("xlsx: workbook.xml.rels: %w", err)
		}
		start, ok := tok.(xml.StartElement)
		if !ok || start.Name.Local != "Relationship" {
			continue
		}
		var id, target string
		for _, a := range start.Attr {
			switch a.Name.Local {
			case "Id":
				id = a.Value
			case "Target":
				target = a.Value
			}
		}
		if id != "" {
			rels[id] = target
		}
	}
	return rels, nil
}

// parseISODate parses the date forms permitted for cells of type "d".
func parseISODate(s string) (time.Time, error) {
	layouts := []string{
		time.RFC3339Nano,
		time.RFC3339,
		"2006-01-02T15:04:05.999999999",
		"2006-01-02T15:04:05",
		"2006-01-02T15:04",
		"2006-01-02",
	}
	for _, l := range layouts {
		if t, err := time.Parse(l, s); err == nil {
			return t.UTC(), nil
		}
	}
	return time.Time{}, fmt.Errorf("xlsx: %q is not an ISO 8601 date", s)
}

// compile-time guarantee that the reader satisfies the shared source contract.
var _ cells.Source = (*Reader)(nil)
