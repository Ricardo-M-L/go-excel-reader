// Package excel reads Excel and CSV files through a single streaming cursor.
//
// It is a Go port of the core of ExcelDataReader, covering the .xlsx, .xls
// (BIFF8) and .csv formats. There is no CGo and no dependency beyond
// golang.org/x/text.
//
// The cursor model is deliberately close to the original's IExcelDataReader:
// advance a sheet, advance a row, read cells by column index.
//
// Rows stream: worksheets are parsed as a token stream and never accumulated,
// so the number of rows does not affect peak memory. The one structure held
// fully in memory is the workbook's shared string table, which is why memory
// scales with the number of distinct strings rather than with the file size.
// A million cells drawn from a small set of values decode in roughly 9MB; the
// same million cells with a distinct string in every one cost about 39MB.
//
//	r, err := excel.OpenFile("report.xlsx")
//	if err != nil {
//	        return err
//	}
//	defer r.Close()
//
//	for r.NextSheet() {
//	        for r.Read() {
//	                for i := 0; i < r.FieldCount(); i++ {
//	                        fmt.Println(r.GetValue(i))
//	                }
//	        }
//	}
//	return r.Err()
package excel

import (
	"bytes"
	"fmt"
	"io"
	"os"
	"time"

	"golang.org/x/text/encoding"

	xcsv "github.com/jloor/go-excel-reader/csv"
	"github.com/jloor/go-excel-reader/internal/cells"
	"github.com/jloor/go-excel-reader/xls"
	"github.com/jloor/go-excel-reader/xlsx"
)

// Cell is a single parsed cell: a Go-typed Value and the Raw text it came from.
type Cell = cells.Cell

// Row is one row of cells.
type Row []Cell

// Reader is the streaming cursor over a workbook.
//
// Two methods go beyond the shape of the C# original, both because the rest of
// this API cannot be honest without them. Err surfaces failures that happen
// mid-stream, which a bool-returning Read has no way to report and which would
// otherwise make a truncated file look like a short one. Cell exposes the raw
// text alongside the typed value, which ReadAll needs in order to populate the
// Raw field it is specified to return.
type Reader interface {
	// Sheet navigation.
	NextSheet() bool
	SheetName() string
	SheetCount() int

	// Row iteration within the current sheet.
	Read() bool
	FieldCount() int

	// Typed getters. Each reports false when the cell is empty, out of range,
	// or holds a different kind of value; int and float interconvert.
	GetValue(col int) any
	GetString(col int) (string, bool)
	GetFloat(col int) (float64, bool)
	GetInt(col int) (int64, bool)
	GetBool(col int) (bool, bool)
	GetTime(col int) (time.Time, bool)
	IsEmpty(col int) bool

	// Cell returns the whole cell, including the raw text it was parsed from.
	Cell(col int) (Cell, bool)

	// RowIndex is the zero-based index of the current row, counting rows the
	// file omitted, so it always matches the row number shown in Excel.
	RowIndex() int

	// Err returns the first error encountered while streaming, if any.
	Err() error

	Close() error
}

// reader adapts a format backend to the public interface.
//
// Every backend produces rows of cells and nothing more; the value coercion
// lives here, once, so that xlsx, xls and csv cannot drift apart in how they
// answer GetInt on a whole-numbered float.
type reader struct {
	src cells.Source
}

func (r *reader) NextSheet() bool   { return r.src.NextSheet() }
func (r *reader) SheetName() string { return r.src.SheetName() }
func (r *reader) SheetCount() int   { return r.src.SheetCount() }
func (r *reader) Read() bool        { return r.src.Read() }
func (r *reader) FieldCount() int   { return len(r.src.Row()) }
func (r *reader) RowIndex() int     { return r.src.RowIndex() }
func (r *reader) Err() error        { return r.src.Err() }
func (r *reader) Close() error      { return r.src.Close() }

func (r *reader) Cell(col int) (Cell, bool) {
	row := r.src.Row()
	if col < 0 || col >= len(row) {
		return Cell{}, false
	}
	return row[col], true
}

func (r *reader) GetValue(col int) any {
	c, _ := r.Cell(col)
	return c.Value
}

func (r *reader) GetString(col int) (string, bool) {
	c, ok := r.Cell(col)
	if !ok {
		return "", false
	}
	return cells.String(c)
}

func (r *reader) GetFloat(col int) (float64, bool) {
	c, ok := r.Cell(col)
	if !ok {
		return 0, false
	}
	return cells.Float(c)
}

func (r *reader) GetInt(col int) (int64, bool) {
	c, ok := r.Cell(col)
	if !ok {
		return 0, false
	}
	return cells.Int(c)
}

func (r *reader) GetBool(col int) (bool, bool) {
	c, ok := r.Cell(col)
	if !ok {
		return false, false
	}
	return cells.Bool(c)
}

func (r *reader) GetTime(col int) (time.Time, bool) {
	c, ok := r.Cell(col)
	if !ok {
		return time.Time{}, false
	}
	return cells.Time(c)
}

func (r *reader) IsEmpty(col int) bool {
	c, ok := r.Cell(col)
	return !ok || c.IsEmpty()
}

// Option configures a spreadsheet reader. Options that apply only to delimited
// text are CSVOption values instead.
type Option func(*options)

type options struct {
	errorValues bool
}

// WithErrorValues makes formula error cells such as #DIV/0! readable through
// GetValue and GetString as their display text.
//
// The default is off, matching ExcelDataReader, which reports an error cell as
// having no value. Either way the text is preserved in Cell.Raw, so this option
// only changes whether an error is a value or an absence.
func WithErrorValues(on bool) Option {
	return func(o *options) { o.errorValues = on }
}

func buildOptions(opts []Option) options {
	var o options
	for _, f := range opts {
		f(&o)
	}
	return o
}

// OpenFile opens a spreadsheet by path, detecting its format from content.
//
// The returned Reader holds the file open until Close. Because an *os.File
// supports random access, xlsx and xls read straight from disk without the
// whole file ever being buffered.
func OpenFile(path string, opts ...Option) (Reader, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}

	r, err := openSeeker(f, f, buildOptions(opts))
	if err != nil {
		f.Close()
		return nil, fmt.Errorf("%s: %w", path, err)
	}
	return r, nil
}

// OpenReader opens a spreadsheet from an io.ReadSeeker, detecting its format
// from the leading bytes.
//
// If r does not also implement io.ReaderAt, the xlsx and xls paths buffer the
// stream in memory, because both formats require random access. Pass an
// *os.File, a *bytes.Reader or your own ReaderAt to stream instead.
func OpenReader(r io.ReadSeeker, opts ...Option) (Reader, error) {
	return openSeeker(r, nil, buildOptions(opts))
}

func openSeeker(rs io.ReadSeeker, closer io.Closer, o options) (Reader, error) {
	format, err := Detect(rs)
	if err != nil {
		return nil, err
	}
	if _, err := rs.Seek(0, io.SeekStart); err != nil {
		return nil, err
	}

	switch format {
	case FormatXLSX:
		return openXLSX(rs, closer, o)
	case FormatXLS:
		return openXLS(rs, closer, o)
	case FormatCSV:
		return openCSV(rs, closer, nil)
	default:
		return nil, ErrUnknownFormat
	}
}

// OpenXLSX opens an OpenXML spreadsheet without format detection.
func OpenXLSX(r io.ReadSeeker, opts ...Option) (Reader, error) {
	return openXLSX(r, nil, buildOptions(opts))
}

// OpenXLS opens a BIFF8 spreadsheet without format detection.
func OpenXLS(r io.ReadSeeker, opts ...Option) (Reader, error) {
	return openXLS(r, nil, buildOptions(opts))
}

// OpenCSV opens a delimited text file.
func OpenCSV(r io.Reader, opts ...CSVOption) (Reader, error) {
	return openCSV(r, nil, opts)
}

func openXLSX(rs io.ReadSeeker, closer io.Closer, o options) (Reader, error) {
	ra, size, err := randomAccess(rs)
	if err != nil {
		return nil, err
	}
	x, err := xlsx.Open(ra, size)
	if err != nil {
		return nil, err
	}
	x.SetErrorValues(o.errorValues)
	if closer != nil {
		x.SetCloser(closer)
	}
	return &reader{src: x}, nil
}

func openXLS(rs io.ReadSeeker, closer io.Closer, o options) (Reader, error) {
	ra, size, err := randomAccess(rs)
	if err != nil {
		return nil, err
	}
	x, err := xls.Open(ra, size)
	if err != nil {
		return nil, err
	}
	x.SetErrorValues(o.errorValues)
	if closer != nil {
		x.SetCloser(closer)
	}
	return &reader{src: x}, nil
}

func openCSV(r io.Reader, closer io.Closer, opts []CSVOption) (Reader, error) {
	var cfg csvConfig
	for _, o := range opts {
		o(&cfg)
	}
	c, err := xcsv.New(r, cfg)
	if err != nil {
		return nil, err
	}
	if closer != nil {
		c.SetCloser(closer)
	}
	return &reader{src: c}, nil
}

// randomAccess adapts a ReadSeeker to the ReaderAt that ZIP and CFB both need.
//
// Neither container can be read front to back: a ZIP's directory lives at the
// end of the file and a CFB's sector chains jump arbitrarily. When the caller's
// stream cannot seek arbitrarily we buffer it, which is the one case where this
// library holds a whole file in memory.
func randomAccess(rs io.ReadSeeker) (io.ReaderAt, int64, error) {
	if ra, ok := rs.(io.ReaderAt); ok {
		size, err := rs.Seek(0, io.SeekEnd)
		if err != nil {
			return nil, 0, err
		}
		if _, err := rs.Seek(0, io.SeekStart); err != nil {
			return nil, 0, err
		}
		return ra, size, nil
	}

	if _, err := rs.Seek(0, io.SeekStart); err != nil {
		return nil, 0, err
	}
	data, err := io.ReadAll(rs)
	if err != nil {
		return nil, 0, err
	}
	return bytes.NewReader(data), int64(len(data)), nil
}

// CSVOption configures the delimited-text reader.
type CSVOption func(*csvConfig)

// csvConfig is an alias so that options declared here can configure the csv
// package's exported Config without either package importing the other.
type csvConfig = xcsv.Config

// WithCSVDelimiter sets the field separator. The default is a comma.
func WithCSVDelimiter(r rune) CSVOption {
	return func(c *csvConfig) { c.Delimiter = r }
}

// WithCSVEncoding decodes the stream from a non-UTF-8 encoding, for example
// charmap.Windows1252 from golang.org/x/text/encoding/charmap.
func WithCSVEncoding(enc encoding.Encoding) CSVOption {
	return func(c *csvConfig) { c.Encoding = enc }
}

// WithCSVTypeInference turns bare text into typed values, trying integer, then
// float, then unambiguous ISO dates. It is off by default: a CSV carries no
// type information, so any inference is a guess the file cannot confirm.
func WithCSVTypeInference(on bool) CSVOption {
	return func(c *csvConfig) { c.InferTypes = on }
}

// ReadAll drains a Reader into per-sheet rows.
//
// It is the convenience path and it does hold everything in memory; the cursor
// API exists for anything large enough to care.
func ReadAll(r Reader) ([][]Row, error) {
	var sheets [][]Row
	for r.NextSheet() {
		var rows []Row
		for r.Read() {
			n := r.FieldCount()
			row := make(Row, n)
			for i := 0; i < n; i++ {
				c, _ := r.Cell(i)
				row[i] = c
			}
			rows = append(rows, row)
		}
		sheets = append(sheets, rows)
	}
	return sheets, r.Err()
}
