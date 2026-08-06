package xls

import (
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"unicode/utf16"
)

// This file implements enough of the Compound File Binary format (MS-CFB) to
// pull one named stream out of a .xls file.
//
// CFB is a filesystem inside a file: fixed-size sectors, a FAT that chains them
// together, and a directory of named entries. A .xls is a CFB whose "Workbook"
// stream holds the actual BIFF records. Nothing about the spreadsheet is
// visible until this layer is working, which is why it comes first.
//
// Reference: https://learn.microsoft.com/en-us/openspecs/windows_protocols/ms-cfb/

var (
	// ErrNotCFB reports a file that does not begin with the CFB signature.
	ErrNotCFB = errors.New("xls: not a compound file (bad OLE2 signature)")
	// ErrNoWorkbook reports a compound file with no workbook stream.
	ErrNoWorkbook = errors.New("xls: compound file contains no Workbook stream")
)

// cfbSignature is the fixed 8-byte header magic, historically the "docfile" id.
var cfbSignature = [8]byte{0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1}

// maxRegSect is the largest value that names a real sector. Everything above
// it is a sentinel: DIFSECT (FFFFFFFC), FATSECT (FFFFFFFD), ENDOFCHAIN
// (FFFFFFFE) and FREESECT (FFFFFFFF). A chain walk only needs to know where
// the sector numbers stop, so it tests against this bound rather than against
// each sentinel in turn.
const maxRegSect uint32 = 0xFFFFFFFA

// dirEntrySize is the fixed size of a directory record.
const dirEntrySize = 128

// cfb is an opened compound file.
type cfb struct {
	ra   io.ReaderAt
	size int64

	sectorSize     int
	miniSectorSize int
	miniCutoff     uint32

	fat     []uint32
	miniFAT []uint32
	entries []dirEntry

	// miniStream is the root entry's stream, inside which mini sectors live.
	miniStream *stream
}

// dirEntry is one directory record.
//
// The directory is stored as a red-black tree, but the entries also sit
// contiguously in their sectors, so finding a stream by name is a linear scan
// over all of them. That is why the sibling and child pointers are not kept:
// walking the tree would be asymptotically better on a directory far larger
// than any spreadsheet has.
type dirEntry struct {
	name    string
	objType byte
	start   uint32
	size    int64
}

const (
	objTypeStream  byte = 2
	objTypeRootDir byte = 5
)

// openCFB parses the header, FAT and directory of a compound file.
func openCFB(ra io.ReaderAt, size int64) (*cfb, error) {
	header := make([]byte, 512)
	if _, err := ra.ReadAt(header, 0); err != nil {
		return nil, fmt.Errorf("xls: reading compound file header: %w", err)
	}
	for i, b := range cfbSignature {
		if header[i] != b {
			return nil, ErrNotCFB
		}
	}

	major := binary.LittleEndian.Uint16(header[26:28])
	sectorShift := binary.LittleEndian.Uint16(header[30:32])
	miniShift := binary.LittleEndian.Uint16(header[32:34])

	c := &cfb{
		ra:             ra,
		size:           size,
		sectorSize:     1 << sectorShift,
		miniSectorSize: 1 << miniShift,
		miniCutoff:     binary.LittleEndian.Uint32(header[56:60]),
	}
	if c.sectorSize < 128 || c.sectorSize > 1<<20 {
		return nil, fmt.Errorf("xls: implausible sector size %d", c.sectorSize)
	}
	if c.miniSectorSize <= 0 || c.miniSectorSize > c.sectorSize {
		return nil, fmt.Errorf("xls: implausible mini sector size %d", c.miniSectorSize)
	}

	numFATSectors := binary.LittleEndian.Uint32(header[44:48])
	firstDirSector := binary.LittleEndian.Uint32(header[48:52])
	firstMiniFAT := binary.LittleEndian.Uint32(header[60:64])
	numMiniFAT := binary.LittleEndian.Uint32(header[64:68])
	firstDIFAT := binary.LittleEndian.Uint32(header[68:72])
	numDIFAT := binary.LittleEndian.Uint32(header[72:76])

	fatSectors, err := c.readDIFAT(header, firstDIFAT, numDIFAT, numFATSectors)
	if err != nil {
		return nil, err
	}
	if err := c.readFAT(fatSectors); err != nil {
		return nil, err
	}
	if err := c.readDirectory(firstDirSector, major); err != nil {
		return nil, err
	}
	if err := c.readMiniFAT(firstMiniFAT, numMiniFAT); err != nil {
		return nil, err
	}
	return c, nil
}

// sectorOffset returns the byte offset of a sector.
//
// Sector numbering starts after the 512-byte header, so sector 0 begins at one
// sector size in. For 4096-byte sectors the header is padded out to a full
// sector, which keeps this formula the same for both versions.
func (c *cfb) sectorOffset(sector uint32) int64 {
	return int64(sector+1) * int64(c.sectorSize)
}

func (c *cfb) readSector(sector uint32) ([]byte, error) {
	off := c.sectorOffset(sector)
	if off < 0 || off+int64(c.sectorSize) > c.size {
		return nil, fmt.Errorf("xls: sector %d lies outside the file", sector)
	}
	buf := make([]byte, c.sectorSize)
	if _, err := c.ra.ReadAt(buf, off); err != nil {
		return nil, fmt.Errorf("xls: reading sector %d: %w", sector, err)
	}
	return buf, nil
}

// readDIFAT collects the list of sectors that make up the FAT.
//
// The first 109 FAT sector numbers live in the header. Beyond that they spill
// into DIFAT sectors, each of which ends with a pointer to the next, which is
// why this is a chain walk rather than a single read.
func (c *cfb) readDIFAT(header []byte, firstDIFAT, numDIFAT, numFATSectors uint32) ([]uint32, error) {
	sectors := make([]uint32, 0, numFATSectors)

	for i := 0; i < 109 && uint32(len(sectors)) < numFATSectors; i++ {
		s := binary.LittleEndian.Uint32(header[76+i*4 : 80+i*4])
		if s > maxRegSect {
			break
		}
		sectors = append(sectors, s)
	}

	perSector := c.sectorSize/4 - 1
	next := firstDIFAT
	for guard := uint32(0); next <= maxRegSect && uint32(len(sectors)) < numFATSectors; guard++ {
		if guard > numDIFAT+1 {
			return nil, errors.New("xls: DIFAT chain does not terminate")
		}
		buf, err := c.readSector(next)
		if err != nil {
			return nil, err
		}
		for i := 0; i < perSector && uint32(len(sectors)) < numFATSectors; i++ {
			s := binary.LittleEndian.Uint32(buf[i*4 : i*4+4])
			if s > maxRegSect {
				continue
			}
			sectors = append(sectors, s)
		}
		next = binary.LittleEndian.Uint32(buf[perSector*4:])
	}

	if len(sectors) == 0 {
		return nil, errors.New("xls: compound file has no FAT sectors")
	}
	return sectors, nil
}

func (c *cfb) readFAT(fatSectors []uint32) error {
	c.fat = make([]uint32, 0, len(fatSectors)*(c.sectorSize/4))
	for _, s := range fatSectors {
		buf, err := c.readSector(s)
		if err != nil {
			return err
		}
		for i := 0; i+4 <= len(buf); i += 4 {
			c.fat = append(c.fat, binary.LittleEndian.Uint32(buf[i:i+4]))
		}
	}
	return nil
}

func (c *cfb) readMiniFAT(first, count uint32) error {
	if count == 0 || first > maxRegSect {
		return nil
	}
	chain, err := c.chainOf(first, int64(count)*int64(c.sectorSize))
	if err != nil {
		return err
	}
	c.miniFAT = make([]uint32, 0, len(chain)*(c.sectorSize/4))
	for _, s := range chain {
		buf, err := c.readSector(s)
		if err != nil {
			return err
		}
		for i := 0; i+4 <= len(buf); i += 4 {
			c.miniFAT = append(c.miniFAT, binary.LittleEndian.Uint32(buf[i:i+4]))
		}
	}
	return nil
}

func (c *cfb) readDirectory(first uint32, major uint16) error {
	chain, err := c.chainOf(first, -1)
	if err != nil {
		return err
	}

	for _, s := range chain {
		buf, err := c.readSector(s)
		if err != nil {
			return err
		}
		for off := 0; off+dirEntrySize <= len(buf); off += dirEntrySize {
			e := buf[off : off+dirEntrySize]
			nameLen := int(binary.LittleEndian.Uint16(e[64:66]))
			entry := dirEntry{
				objType: e[66],
				start:   binary.LittleEndian.Uint32(e[116:120]),
			}

			size := binary.LittleEndian.Uint64(e[120:128])
			if major == 3 {
				// Version 3 defines only the low 32 bits; some writers leave
				// junk in the high half, which would otherwise be read as a
				// multi-exabyte stream.
				size &= 0xFFFFFFFF
			}
			entry.size = int64(size)
			entry.name = decodeEntryName(e[0:64], nameLen)

			c.entries = append(c.entries, entry)
		}
	}

	if len(c.entries) == 0 {
		return errors.New("xls: compound file has an empty directory")
	}

	// The root entry doubles as the container for every mini sector.
	root := c.entries[0]
	if root.objType != objTypeRootDir {
		for _, e := range c.entries {
			if e.objType == objTypeRootDir {
				root = e
				break
			}
		}
	}
	if root.size > 0 && root.start <= maxRegSect {
		chain, err := c.chainOf(root.start, root.size)
		if err != nil {
			return err
		}
		c.miniStream = &stream{cfb: c, chain: chain, size: root.size, unit: c.sectorSize}
	}
	return nil
}

// decodeEntryName converts a directory entry's UTF-16LE name.
func decodeEntryName(raw []byte, nameLen int) string {
	// nameLen counts bytes including the terminating NUL.
	if nameLen < 2 || nameLen > len(raw) {
		nameLen = len(raw)
	}
	n := (nameLen - 2) / 2
	if n <= 0 {
		return ""
	}
	u := make([]uint16, n)
	for i := 0; i < n; i++ {
		u[i] = binary.LittleEndian.Uint16(raw[i*2 : i*2+2])
	}
	return string(utf16.Decode(u))
}

// chainOf walks the FAT from a starting sector, returning the sector list.
//
// Materialising the chain up front turns every later read into an index
// lookup. A corrupt file can describe a cycle, so the walk is bounded by the
// size of the FAT itself.
func (c *cfb) chainOf(start uint32, byteSize int64) ([]uint32, error) {
	var (
		chain []uint32
		want  = -1
	)
	if byteSize >= 0 {
		want = int((byteSize + int64(c.sectorSize) - 1) / int64(c.sectorSize))
	}

	for s := start; s <= maxRegSect; {
		if int(s) >= len(c.fat) {
			return nil, fmt.Errorf("xls: sector %d is outside the FAT", s)
		}
		chain = append(chain, s)
		if want >= 0 && len(chain) >= want {
			break
		}
		if len(chain) > len(c.fat) {
			return nil, errors.New("xls: sector chain contains a cycle")
		}
		s = c.fat[s]
	}
	return chain, nil
}

// miniChainOf is chainOf for the mini FAT.
func (c *cfb) miniChainOf(start uint32, byteSize int64) ([]uint32, error) {
	var chain []uint32
	want := int((byteSize + int64(c.miniSectorSize) - 1) / int64(c.miniSectorSize))

	for s := start; s <= maxRegSect && len(chain) < want; {
		if int(s) >= len(c.miniFAT) {
			return nil, fmt.Errorf("xls: mini sector %d is outside the mini FAT", s)
		}
		chain = append(chain, s)
		if len(chain) > len(c.miniFAT) {
			return nil, errors.New("xls: mini sector chain contains a cycle")
		}
		s = c.miniFAT[s]
	}
	return chain, nil
}

// openStream returns a reader over the named stream.
func (c *cfb) openStream(names ...string) (*stream, error) {
	for _, want := range names {
		for _, e := range c.entries {
			if e.objType != objTypeStream || e.name != want {
				continue
			}
			return c.streamFor(e)
		}
	}
	return nil, ErrNoWorkbook
}

func (c *cfb) streamFor(e dirEntry) (*stream, error) {
	if e.size == 0 {
		return &stream{cfb: c, size: 0, unit: c.sectorSize}, nil
	}

	// Small streams are packed into the mini stream rather than given whole
	// sectors of their own, so which FAT applies depends on the stream's size.
	if uint32(e.size) < c.miniCutoff {
		if c.miniStream == nil {
			return nil, errors.New("xls: stream needs the mini FAT, which is absent")
		}
		chain, err := c.miniChainOf(e.start, e.size)
		if err != nil {
			return nil, err
		}
		return &stream{cfb: c, chain: chain, size: e.size, unit: c.miniSectorSize, mini: true}, nil
	}

	chain, err := c.chainOf(e.start, e.size)
	if err != nil {
		return nil, err
	}
	return &stream{cfb: c, chain: chain, size: e.size, unit: c.sectorSize}, nil
}

// stream is a seekable reader over a sector chain.
//
// Reads pull only the sectors they touch, so opening a 200MB workbook stream
// costs the chain list and nothing else.
type stream struct {
	cfb   *cfb
	chain []uint32
	size  int64
	unit  int
	mini  bool
	pos   int64
}

// Size returns the stream length in bytes.
func (s *stream) Size() int64 { return s.size }

func (s *stream) Read(p []byte) (int, error) {
	n, err := s.ReadAt(p, s.pos)
	s.pos += int64(n)
	return n, err
}

func (s *stream) Seek(offset int64, whence int) (int64, error) {
	var abs int64
	switch whence {
	case io.SeekStart:
		abs = offset
	case io.SeekCurrent:
		abs = s.pos + offset
	case io.SeekEnd:
		abs = s.size + offset
	default:
		return 0, fmt.Errorf("xls: invalid whence %d", whence)
	}
	if abs < 0 {
		return 0, errors.New("xls: negative stream position")
	}
	s.pos = abs
	return abs, nil
}

func (s *stream) ReadAt(p []byte, off int64) (int, error) {
	if off >= s.size {
		return 0, io.EOF
	}
	if off < 0 {
		return 0, errors.New("xls: negative stream offset")
	}

	read := 0
	for read < len(p) && off+int64(read) < s.size {
		cur := off + int64(read)
		idx := int(cur / int64(s.unit))
		within := int(cur % int64(s.unit))
		if idx >= len(s.chain) {
			break
		}

		want := s.unit - within
		if remaining := int(s.size - cur); want > remaining {
			want = remaining
		}
		if want > len(p)-read {
			want = len(p) - read
		}

		var (
			base int64
			err  error
		)
		if s.mini {
			base = int64(s.chain[idx])*int64(s.cfb.miniSectorSize) + int64(within)
			_, err = s.cfb.miniStream.ReadAt(p[read:read+want], base)
		} else {
			base = s.cfb.sectorOffset(s.chain[idx]) + int64(within)
			_, err = s.cfb.ra.ReadAt(p[read:read+want], base)
		}
		if err != nil {
			return read, err
		}
		read += want
	}

	if read < len(p) {
		return read, io.EOF
	}
	return read, nil
}
