// Package cells defines the value model shared by every format backend.
//
// It exists so that the typed getters (GetString, GetFloat, ...) are written
// exactly once, in the root excel package, rather than duplicated across the
// xlsx, xls and csv readers. It is a leaf package: the format backends and the
// root package both import it, and it imports nothing of ours in return, which
// is what keeps the dependency graph acyclic.
package cells

import (
	"math"
	"strconv"
	"strings"
	"time"

	"github.com/jloor/go-excel-reader/numfmt"
)

// Cell is a single parsed cell.
//
// Value holds the Go-typed interpretation and is one of:
// nil, string, float64, int64, bool or time.Time.
//
// Raw holds the value as it literally appeared in the file, before typing.
// For xlsx that is the text of the <v> element, for xls it is the decimal
// rendering of the stored number or the stored string, and for csv it is the
// unmodified field. Raw is the escape hatch for callers who disagree with our
// typing, and it is never lost.
type Cell struct {
	Value any
	Raw   string

	numberFormat string
	date1904     bool
}

// IsEmpty reports whether the cell holds no value.
func (c Cell) IsEmpty() bool { return c.Value == nil }

// Number builds a numeric cell with the format metadata needed to reproduce
// the text Excel displays without changing the cell's typed Value.
func Number(value any, raw, numberFormat string, date1904 bool) Cell {
	return Cell{Value: value, Raw: raw, numberFormat: numberFormat, date1904: date1904}
}

// FormattedValue returns the text Excel displays for the cell. Numeric cells
// use their workbook number format; other cell types use their natural display
// representation. Formula errors remain available here even when Value is nil.
func (c Cell) FormattedValue() string {
	if c.numberFormat != "" {
		if value, err := strconv.ParseFloat(strings.TrimSpace(c.Raw), 64); err == nil {
			return numfmt.FormatValueIn(value, c.numberFormat, c.date1904)
		}
	}

	switch value := c.Value.(type) {
	case nil:
		return c.Raw
	case string:
		return value
	case bool:
		return strings.ToUpper(strconv.FormatBool(value))
	case time.Time:
		return value.Format(time.RFC3339)
	default:
		if c.Raw != "" {
			return c.Raw
		}
		return ""
	}
}

// Source is the format-specific engine that sits behind a public Reader.
//
// Each backend (xlsx, xls, csv) implements it, and the root package wraps a
// Source to produce the exported Reader interface. Splitting it this way means
// a backend only has to know how to produce rows of Cells; it never reimplements
// value coercion.
//
// Read must return the rows of the current sheet in order. Row returns the
// cells of the row most recently returned by Read; the slice is only valid
// until the next call to Read and backends are free to reuse it.
type Source interface {
	NextSheet() bool
	SheetName() string
	SheetCount() int
	Read() bool
	Row() []Cell
	RowIndex() int
	Err() error
	Close() error
}

// The functions below implement the coercion policy for the typed getters.
//
// The policy is deliberately strict: a getter returns ok=false when the cell
// holds a different kind of value, rather than attempting to parse strings into
// numbers or dates. Silent string-to-number coercion is how spreadsheet readers
// quietly corrupt data, and callers who want it can reach for Cell.Raw.
//
// The single exception is that int64 and float64 interconvert, because whether
// a numeric cell lands as one or the other depends on incidental details of how
// the file was written, which is not something a caller should have to reason
// about.

// String returns the cell's string value.
func String(c Cell) (string, bool) {
	s, ok := c.Value.(string)
	return s, ok
}

// Float returns the cell's value as a float64, widening an integer if needed.
func Float(c Cell) (float64, bool) {
	switch v := c.Value.(type) {
	case float64:
		return v, true
	case int64:
		return float64(v), true
	default:
		return 0, false
	}
}

// Int returns the cell's value as an int64. A float is accepted only when it is
// integral and survives the round trip exactly, so 3.0 converts but 3.5 does not.
func Int(c Cell) (int64, bool) {
	switch v := c.Value.(type) {
	case int64:
		return v, true
	case float64:
		if math.IsNaN(v) || math.IsInf(v, 0) || v != math.Trunc(v) {
			return 0, false
		}
		if v < math.MinInt64 || v >= math.MaxInt64 {
			return 0, false
		}
		return int64(v), true
	default:
		return 0, false
	}
}

// Bool returns the cell's boolean value.
func Bool(c Cell) (bool, bool) {
	b, ok := c.Value.(bool)
	return b, ok
}

// Time returns the cell's date/time value.
func Time(c Cell) (time.Time, bool) {
	t, ok := c.Value.(time.Time)
	return t, ok
}
