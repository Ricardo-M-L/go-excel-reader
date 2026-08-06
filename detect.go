package excel

import (
	"archive/zip"
	"errors"
	"fmt"
	"io"
	"unicode/utf8"
)

// Format identifies a spreadsheet container.
type Format int

const (
	FormatUnknown Format = iota
	FormatXLSX
	FormatXLS
	FormatCSV
)

func (f Format) String() string {
	switch f {
	case FormatXLSX:
		return "xlsx"
	case FormatXLS:
		return "xls"
	case FormatCSV:
		return "csv"
	default:
		return "unknown"
	}
}

var (
	// ErrUnknownFormat reports a stream that matches no supported format.
	ErrUnknownFormat = errors.New("excel: unrecognised file format")
	// ErrXLSBUnsupported reports an OpenXML Binary workbook, which this
	// version does not read.
	ErrXLSBUnsupported = errors.New("excel: xlsb (OpenXML Binary) workbooks are not supported")
)

// Magic byte sequences that identify the two binary containers.
var (
	// zipMagic starts every ZIP local file header, and so every xlsx.
	zipMagic = []byte{'P', 'K', 0x03, 0x04}
	// cfbMagic starts a Compound File Binary header, the OLE2 container that
	// holds a BIFF workbook.
	cfbMagic = []byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}
)

// detectSniffLen is how much of a candidate text file we inspect. It is enough
// to catch binary content without reading a large file twice.
const detectSniffLen = 8192

// Detect identifies the format of a stream from its content.
//
// The stream is rewound to its start before returning, so a caller can hand the
// same reader straight to an Open function.
//
// Detection is by content rather than by file extension, because the extension
// is the least reliable thing about a spreadsheet: files named .xls are
// routinely xlsx, CSV, or HTML that Excel happens to open.
func Detect(rs io.ReadSeeker) (Format, error) {
	if _, err := rs.Seek(0, io.SeekStart); err != nil {
		return FormatUnknown, err
	}

	head := make([]byte, detectSniffLen)
	n, err := io.ReadFull(rs, head)
	if err != nil && !errors.Is(err, io.EOF) && !errors.Is(err, io.ErrUnexpectedEOF) {
		return FormatUnknown, err
	}
	head = head[:n]

	if _, err := rs.Seek(0, io.SeekStart); err != nil {
		return FormatUnknown, err
	}

	switch {
	case hasPrefix(head, cfbMagic):
		// A CFB container. It could in principle hold something other than a
		// workbook, but the stream lookup in the xls reader will say so.
		return FormatXLS, nil

	case hasPrefix(head, zipMagic):
		// A ZIP could be xlsx, xlsb, ODS, or an unrelated archive. Only the
		// central directory can tell them apart, so look inside.
		return detectZip(rs)

	case n == 0:
		// An empty stream is an empty CSV, not an error: reading it yields no
		// rows, which is the honest answer.
		return FormatCSV, nil

	case looksLikeText(head):
		return FormatCSV, nil

	default:
		return FormatUnknown, ErrUnknownFormat
	}
}

// detectZip distinguishes xlsx from the other ZIP-based spreadsheet formats by
// the parts the archive contains.
func detectZip(rs io.ReadSeeker) (Format, error) {
	ra, size, err := randomAccess(rs)
	if err != nil {
		return FormatUnknown, err
	}
	defer func() { _, _ = rs.Seek(0, io.SeekStart) }()

	zr, err := zip.NewReader(ra, size)
	if err != nil {
		return FormatUnknown, fmt.Errorf("excel: reading zip container: %w", err)
	}

	var sawBin bool
	for _, f := range zr.File {
		switch normalizeZipName(f.Name) {
		case "xl/workbook.xml":
			return FormatXLSX, nil
		case "xl/workbook.bin":
			sawBin = true
		}
	}
	if sawBin {
		return FormatUnknown, ErrXLSBUnsupported
	}
	return FormatUnknown, fmt.Errorf("%w: zip archive has no xl/workbook.xml", ErrUnknownFormat)
}

func normalizeZipName(name string) string {
	out := make([]byte, 0, len(name))
	for i := 0; i < len(name); i++ {
		c := name[i]
		if c == '\\' {
			c = '/'
		}
		if c >= 'A' && c <= 'Z' {
			c += 'a' - 'A'
		}
		out = append(out, c)
	}
	s := string(out)
	if len(s) > 0 && s[0] == '/' {
		s = s[1:]
	}
	return s
}

func hasPrefix(b, prefix []byte) bool {
	if len(b) < len(prefix) {
		return false
	}
	for i := range prefix {
		if b[i] != prefix[i] {
			return false
		}
	}
	return true
}

// looksLikeText reports whether a sample plausibly holds delimited text.
//
// This is the last resort in detection, and it needs to be strict: falling back
// to the CSV reader on a binary file produces pages of mojibake instead of an
// error. A NUL byte is the strongest signal, since text files do not contain
// them and every binary format we might meet does.
func looksLikeText(sample []byte) bool {
	for _, b := range sample {
		if b == 0x00 {
			return false
		}
	}

	// Reject anything that is not valid UTF-8, except for a truncation at the
	// tail of the sample, which only means the sample cut a rune in half.
	for len(sample) > 0 {
		r, size := utf8.DecodeRune(sample)
		if r == utf8.RuneError && size <= 1 {
			if len(sample) < utf8.UTFMax {
				break // an incomplete rune at the sample boundary
			}
			return false
		}
		sample = sample[size:]
	}
	return true
}
