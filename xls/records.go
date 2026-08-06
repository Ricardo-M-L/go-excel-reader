package xls

import (
	"errors"
	"fmt"
	"math"
	"unicode/utf16"
)

// BIFF record type numbers. Only the records that carry data this reader needs
// are listed; everything else in a workbook stream is skipped by type.
const (
	recFormula    uint16 = 0x0006
	recEOF        uint16 = 0x000A
	recCalcCount  uint16 = 0x000C
	recDateMode   uint16 = 0x0022
	recCodePage   uint16 = 0x0042
	recContinue   uint16 = 0x003C
	recBoundSheet uint16 = 0x0085
	recMulRK      uint16 = 0x00BD
	recMulBlank   uint16 = 0x00BE
	recRString    uint16 = 0x00D6
	recXF         uint16 = 0x00E0
	recSST        uint16 = 0x00FC
	recLabelSST   uint16 = 0x00FD
	recDimension  uint16 = 0x0200
	recBlank      uint16 = 0x0201
	recNumber     uint16 = 0x0203
	recLabel      uint16 = 0x0204
	recBoolErr    uint16 = 0x0205
	recString     uint16 = 0x0207
	recRow        uint16 = 0x0208
	recIndex      uint16 = 0x020B
	recRK         uint16 = 0x027E
	recFormat     uint16 = 0x041E
	recBOF        uint16 = 0x0809
)

// BIFF versions, as they appear in the BOF record.
const (
	biff2 uint16 = 0x0200
	biff3 uint16 = 0x0300
	biff4 uint16 = 0x0400
	biff5 uint16 = 0x0500 // also BIFF7
	biff8 uint16 = 0x0600
)

// Substream kinds, from the BOF record's dt field.
const (
	substreamWorkbook  uint16 = 0x0005
	substreamWorksheet uint16 = 0x0010
)

// ErrUnsupportedBIFF reports a workbook older than BIFF8.
//
// BIFF5 and earlier store text in a codepage declared by a CODEPAGE record
// rather than in UTF-16, and their record layouts differ in ways that are not
// worth guessing at. Refusing them with a clear message beats decoding them
// into mojibake.
var ErrUnsupportedBIFF = errors.New("xls: only BIFF8 (Excel 97 and later) workbooks are supported")

func float64frombits(b uint64) float64 { return math.Float64frombits(b) }

// bofInfo is the part of a BOF record this reader uses.
type bofInfo struct {
	version   uint16
	substream uint16
}

func parseBOF(data []byte) (bofInfo, error) {
	if len(data) < 4 {
		return bofInfo{}, errors.New("xls: truncated BOF record")
	}
	return bofInfo{
		version:   uint16(data[0]) | uint16(data[1])<<8,
		substream: uint16(data[2]) | uint16(data[3])<<8,
	}, nil
}

// boundSheet describes one sheet listed in the workbook globals.
type boundSheet struct {
	name   string
	offset int64
	hidden bool
	// kind distinguishes worksheets from chart, macro and module sheets, none
	// of which hold cell data.
	kind byte
}

// sheetKindWorksheet is the BOUNDSHEET type byte for an ordinary worksheet.
const sheetKindWorksheet byte = 0x00

// parseBoundSheet reads a BOUNDSHEET record.
//
// The offset it carries is an absolute position within the workbook stream,
// pointing at the sheet's own BOF record. That is what makes sheet-at-a-time
// reading possible without scanning the whole file.
func parseBoundSheet(data []byte) (boundSheet, error) {
	c := newCursor(data, nil)
	pos, ok := c.u32()
	if !ok {
		return boundSheet{}, errors.New("xls: truncated BOUNDSHEET record")
	}
	grbit, ok := c.u16()
	if !ok {
		return boundSheet{}, errors.New("xls: truncated BOUNDSHEET record")
	}
	name, err := readShortString(c)
	if err != nil {
		return boundSheet{}, err
	}
	return boundSheet{
		name:   name,
		offset: int64(pos),
		// The low two bits of grbit hold the visibility state; anything
		// non-zero means hidden or very hidden. The high byte is the sheet kind.
		hidden: grbit&0x03 != 0,
		kind:   byte(grbit >> 8),
	}, nil
}

// readShortString reads a ShortXLUnicodeString: a one-byte character count
// followed by the flags byte and the characters.
func readShortString(c *cursor) (string, error) {
	cch, ok := c.u8()
	if !ok {
		return "", errors.New("xls: truncated string length")
	}
	return readStringChars(c, int(cch), true)
}

// readLongString reads an XLUnicodeString: a two-byte character count followed
// by the flags byte and the characters.
func readLongString(c *cursor) (string, error) {
	cch, ok := c.u16()
	if !ok {
		return "", errors.New("xls: truncated string length")
	}
	return readStringChars(c, int(cch), true)
}

// readStringChars decodes cch characters, honouring the compression flag and
// the rich-text and phonetic extensions.
//
// BIFF8 stores text as UTF-16LE, except that a string whose characters all fit
// in one byte is stored compressed, one byte per character, holding Latin-1
// code points. The flags byte says which, and the same string may switch
// between the two halfway through when a CONTINUE record intervenes.
func readStringChars(c *cursor, cch int, hasFlags bool) (string, error) {
	var (
		high   bool
		rich   int
		extLen int
	)
	if hasFlags {
		flags, ok := c.u8()
		if !ok {
			return "", errors.New("xls: truncated string flags")
		}
		high = flags&0x01 != 0
		if flags&0x08 != 0 { // fRichSt
			n, ok := c.u16()
			if !ok {
				return "", errors.New("xls: truncated rich-text run count")
			}
			rich = int(n)
		}
		if flags&0x04 != 0 { // fExtSt
			n, ok := c.u32()
			if !ok {
				return "", errors.New("xls: truncated phonetic extension length")
			}
			extLen = int(n)
		}
	}

	units := make([]uint16, 0, cch)
	for i := 0; i < cch; i++ {
		// A CONTINUE that resumes a string's characters restates the
		// compression flag in its first byte. This is the single most
		// error-prone detail of the whole format: miss the byte and every
		// string after the first split comes out shifted by one and garbled.
		if c.atContinuation() {
			flags, ok := c.u8()
			if !ok {
				return "", errors.New("xls: truncated continuation flags")
			}
			high = flags&0x01 != 0
		}

		if high {
			u, ok := c.u16()
			if !ok {
				return "", errors.New("xls: truncated string characters")
			}
			units = append(units, u)
		} else {
			b, ok := c.u8()
			if !ok {
				return "", errors.New("xls: truncated string characters")
			}
			units = append(units, uint16(b))
		}
	}

	// Rich-text run data and phonetic data follow the characters. They may
	// themselves span a CONTINUE, but those boundaries carry no flags byte
	// because they do not resume a character array.
	c.skip(rich * 4)
	c.skip(extLen)

	return string(utf16.Decode(units)), nil
}

// parseSST decodes the shared string table.
//
// The table is almost always longer than a single record can hold, so it
// arrives as a first record plus a run of CONTINUEs. They are concatenated
// before parsing, with the boundaries retained so that split strings can find
// their restated flags byte.
func parseSST(data []byte, contStarts map[int]bool) ([]string, error) {
	c := newCursor(data, contStarts)
	if _, ok := c.u32(); !ok { // total reference count, not needed
		return nil, errors.New("xls: truncated SST record")
	}
	unique, ok := c.u32()
	if !ok {
		return nil, errors.New("xls: truncated SST record")
	}

	// The declared count is trusted only as a capacity hint. A corrupt or
	// hostile file could claim four billion strings, so growth is driven by
	// what actually parses.
	const sane = 1 << 20
	capHint := int(unique)
	if capHint < 0 || capHint > sane {
		capHint = sane
	}

	out := make([]string, 0, capHint)
	for i := uint32(0); i < unique; i++ {
		if c.remaining() <= 0 {
			break
		}
		s, err := readLongString(c)
		if err != nil {
			return out, fmt.Errorf("xls: shared string %d: %w", i, err)
		}
		out = append(out, s)
	}
	return out, nil
}

// parseFormat reads a FORMAT record: a format id and its format string.
func parseFormat(data []byte) (int, string, error) {
	c := newCursor(data, nil)
	id, ok := c.u16()
	if !ok {
		return 0, "", errors.New("xls: truncated FORMAT record")
	}
	s, err := readLongString(c)
	if err != nil {
		return 0, "", err
	}
	return int(id), s, nil
}

// parseXF reads the number format index out of an XF record.
//
// An XF holds font, alignment, borders and much else; the only field that
// matters for typing a value is the number format index at offset 2.
func parseXF(data []byte) (int, error) {
	if len(data) < 4 {
		return 0, errors.New("xls: truncated XF record")
	}
	return int(uint16(data[2]) | uint16(data[3])<<8), nil
}

// cellHeader is the row, column and format index that every cell record starts with.
type cellHeader struct {
	row uint16
	col uint16
	xf  uint16
}

func parseCellHeader(c *cursor) (cellHeader, bool) {
	row, ok1 := c.u16()
	col, ok2 := c.u16()
	xf, ok3 := c.u16()
	return cellHeader{row: row, col: col, xf: xf}, ok1 && ok2 && ok3
}

// decodeRK expands an RK value.
//
// RK packs a number into four bytes using the low two bits as tags: bit 1 says
// the remaining 30 bits are a signed integer rather than the top half of an
// IEEE 754 double, and bit 0 says the result was multiplied by 100 to make it
// fit. All four combinations occur in real files, and getting the integer
// variant's sign extension wrong turns large negative numbers positive.
func decodeRK(rk uint32) float64 {
	var v float64
	if rk&0x02 != 0 {
		// Signed 30-bit integer: shift arithmetically to keep the sign.
		v = float64(int32(rk) >> 2)
	} else {
		// The 30 significant bits are the most significant half of a double;
		// the low 34 bits of the mantissa were discarded when it was stored.
		v = math.Float64frombits(uint64(rk&0xFFFFFFFC) << 32)
	}
	if rk&0x01 != 0 {
		v /= 100
	}
	return v
}

// errValue distinguishes a formula error from an ordinary string result, so
// that the reader can decide how to present it rather than having to guess
// whether a value that looks like "#REF!" was text a user typed.
type errValue string

// errorText maps a BIFF error code to the text Excel displays.
func errorText(code byte) string {
	switch code {
	case 0x00:
		return "#NULL!"
	case 0x07:
		return "#DIV/0!"
	case 0x0F:
		return "#VALUE!"
	case 0x17:
		return "#REF!"
	case 0x1D:
		return "#NAME?"
	case 0x24:
		return "#NUM!"
	case 0x2A:
		return "#N/A"
	default:
		return fmt.Sprintf("#ERR%02X", code)
	}
}

// formulaResult decodes the eight result bytes of a FORMULA record.
//
// A formula caches its last computed value in place. When the top two bytes are
// 0xFFFF the value is not a number, and the first byte says what it is instead:
// a string that arrives in a following STRING record, a boolean, or an error.
func formulaResult(b []byte) (value any, needsString bool) {
	if len(b) < 8 {
		return nil, false
	}
	if b[6] != 0xFF || b[7] != 0xFF {
		bits := uint64(b[0]) | uint64(b[1])<<8 | uint64(b[2])<<16 | uint64(b[3])<<24 |
			uint64(b[4])<<32 | uint64(b[5])<<40 | uint64(b[6])<<48 | uint64(b[7])<<56
		return math.Float64frombits(bits), false
	}

	switch b[0] {
	case 0x00: // string, delivered by the next STRING record
		return nil, true
	case 0x01:
		return b[2] != 0, false
	case 0x02:
		return errValue(errorText(b[2])), false
	case 0x03: // empty string
		return "", false
	default:
		return nil, false
	}
}
