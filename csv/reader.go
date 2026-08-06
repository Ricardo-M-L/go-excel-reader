// Package csv reads delimited text files behind the same streaming cursor as
// the spreadsheet formats.
//
// A CSV has no sheets, no types and no styles, so most of the work here is
// deciding how much structure to invent. The default is to invent none: every
// field is a string, exactly as it appeared. Type inference is available but
// opt-in, because guessing wrong is silent and a CSV carries no metadata to
// check the guess against.
package csv

import (
	"bufio"
	"encoding/csv"
	"errors"
	"fmt"
	"io"
	"strconv"
	"strings"
	"time"

	"golang.org/x/text/encoding"
	"golang.org/x/text/transform"

	"github.com/jloor/go-excel-reader/internal/cells"
)

// Config controls how a delimited file is interpreted.
type Config struct {
	// Delimiter separates fields. Zero means comma.
	Delimiter rune
	// Encoding decodes the byte stream. Nil means UTF-8.
	Encoding encoding.Encoding
	// InferTypes turns bare text into typed values. Off by default.
	InferTypes bool
	// SheetName is reported by SheetName. Zero means "Sheet1".
	SheetName string
}

// Reader streams a delimited file one record at a time.
type Reader struct {
	cr     *csv.Reader
	cfg    Config
	closer io.Closer

	row    []cells.Cell
	rowIdx int

	sheetStarted bool
	exhausted    bool
	err          error
}

// New wraps an io.Reader as a delimited-file reader.
func New(r io.Reader, cfg Config) (*Reader, error) {
	if cfg.Delimiter == 0 {
		cfg.Delimiter = ','
	}
	if cfg.SheetName == "" {
		cfg.SheetName = "Sheet1"
	}

	if cfg.Encoding != nil {
		r = transform.NewReader(r, cfg.Encoding.NewDecoder())
	}

	br := bufio.NewReader(r)
	if err := stripBOM(br); err != nil {
		return nil, err
	}

	cr := csv.NewReader(br)
	cr.Comma = cfg.Delimiter
	// Rows in the wild are ragged; a short row is data, not a syntax error.
	cr.FieldsPerRecord = -1
	// Unescaped quotes inside unquoted fields are common enough in exported
	// data that rejecting them would fail on files every spreadsheet opens.
	cr.LazyQuotes = true
	// The record slice is rebuilt into cells on every Read, so letting the
	// decoder reuse its backing array costs nothing and saves an allocation
	// per row.
	cr.ReuseRecord = true

	return &Reader{cr: cr, cfg: cfg, rowIdx: -1}, nil
}

// SetCloser registers a resource to release when the reader is closed.
func (r *Reader) SetCloser(c io.Closer) { r.closer = c }

// stripBOM consumes a leading UTF-8 byte order mark.
//
// A BOM is invisible to the eye and to most tooling, but left in place it
// becomes part of the first field's text, so the first column header stops
// comparing equal to its own name. It is the most common reason a CSV "looks
// identical" yet fails a lookup.
func stripBOM(br *bufio.Reader) error {
	const bom = "\xef\xbb\xbf"
	prefix, err := br.Peek(len(bom))
	if err != nil && !errors.Is(err, io.EOF) {
		return fmt.Errorf("csv: %w", err)
	}
	if string(prefix) == bom {
		if _, err := br.Discard(len(bom)); err != nil {
			return fmt.Errorf("csv: %w", err)
		}
	}
	return nil
}

// SheetCount reports one sheet; a delimited file has no sheet structure.
func (r *Reader) SheetCount() int { return 1 }

// SheetName returns the synthetic sheet name.
func (r *Reader) SheetName() string {
	if !r.sheetStarted {
		return ""
	}
	return r.cfg.SheetName
}

// NextSheet advances to the file's single sheet, then reports false.
func (r *Reader) NextSheet() bool {
	if r.sheetStarted {
		return false
	}
	r.sheetStarted = true
	return true
}

// Read advances to the next record.
func (r *Reader) Read() bool {
	if !r.sheetStarted || r.exhausted {
		return false
	}
	rec, err := r.cr.Read()
	if errors.Is(err, io.EOF) {
		r.exhausted = true
		return false
	}
	if err != nil {
		r.noteErr(fmt.Errorf("csv: %w", err))
		r.exhausted = true
		return false
	}

	r.row = r.row[:0]
	for _, f := range rec {
		r.row = append(r.row, r.makeCell(f))
	}
	r.rowIdx++
	return true
}

func (r *Reader) makeCell(field string) cells.Cell {
	if !r.cfg.InferTypes {
		if field == "" {
			return cells.Cell{Raw: field}
		}
		return cells.Cell{Value: field, Raw: field}
	}
	return cells.Cell{Value: infer(field), Raw: field}
}

// infer maps text to a Go value, in the order integer, float, date, string.
//
// The date formats accepted here are deliberately limited to unambiguous ones.
// "03/04/2024" is the fourth of March in one country and the third of April in
// another, and nothing in the file says which; a reader that picks one is wrong
// half the time and never says so. Callers who know their format can read Raw.
func infer(field string) any {
	if field == "" {
		return nil
	}
	if n, err := strconv.ParseInt(field, 10, 64); err == nil {
		return n
	}
	if f, err := strconv.ParseFloat(field, 64); err == nil {
		// Reject the forms ParseFloat accepts but a spreadsheet would not,
		// so that "NaN" and "Inf" stay text.
		if !strings.ContainsAny(field, "nNiI") || strings.ContainsAny(field, "eE") {
			return f
		}
	}
	for _, layout := range []string{
		time.RFC3339Nano,
		time.RFC3339,
		"2006-01-02T15:04:05",
		"2006-01-02 15:04:05",
		"2006-01-02 15:04",
		"2006-01-02",
	} {
		if t, err := time.Parse(layout, field); err == nil {
			return t.UTC()
		}
	}
	return field
}

// Row returns the cells of the current record.
func (r *Reader) Row() []cells.Cell { return r.row }

// RowIndex returns the zero-based index of the current record.
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

var _ cells.Source = (*Reader)(nil)
