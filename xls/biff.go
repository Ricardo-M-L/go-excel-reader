package xls

import (
	"encoding/binary"
	"fmt"
	"io"
)

// The BIFF record stream is a flat sequence of records, each a 2-byte type, a
// 2-byte payload length, and that many bytes of payload. The only structural
// complication is CONTINUE: a record's payload may not exceed 8224 bytes, so
// anything longer is emitted as a first record followed by CONTINUE records
// carrying the remainder.

// recordReader walks the record stream of a workbook substream.
//
// The payload buffer is reused between records. Callers that need to keep a
// payload past the next call must copy it; in practice only the shared string
// table does, and it copies into its own storage as it parses.
type recordReader struct {
	s   *stream
	off int64

	typ  uint16
	data []byte
	// contStarts holds the offsets within data at which a CONTINUE record's
	// payload begins. The shared string table needs these because a string
	// split across a CONTINUE carries an extra flags byte at the resume point;
	// every other record can ignore them and treat data as one buffer.
	contStarts map[int]bool

	// recOff is the stream offset of the record just returned, which is what
	// BOUNDSHEET offsets point at.
	recOff int64

	hdr [4]byte
}

func newRecordReader(s *stream) *recordReader {
	return &recordReader{s: s, contStarts: make(map[int]bool)}
}

// seek repositions the reader at an absolute stream offset.
func (r *recordReader) seek(off int64) { r.off = off }

// readHeader reads one record header at off.
func (r *recordReader) readHeader(off int64) (typ, length uint16, err error) {
	if off+4 > r.s.Size() {
		return 0, 0, io.EOF
	}
	if _, err := r.s.ReadAt(r.hdr[:], off); err != nil {
		return 0, 0, err
	}
	return binary.LittleEndian.Uint16(r.hdr[0:2]), binary.LittleEndian.Uint16(r.hdr[2:4]), nil
}

// peek returns the type of the next record without consuming it.
func (r *recordReader) peek() (uint16, error) {
	typ, _, err := r.readHeader(r.off)
	return typ, err
}

// next reads the following record, absorbing any CONTINUE records that extend it.
func (r *recordReader) next() error {
	typ, length, err := r.readHeader(r.off)
	if err != nil {
		return err
	}
	r.recOff = r.off
	r.typ = typ
	r.off += 4

	r.data = r.data[:0]
	clear(r.contStarts)

	if err := r.appendPayload(int(length)); err != nil {
		return err
	}

	// Absorb CONTINUE records. Peeking costs one header read per record, which
	// is cheaper than the alternative of buffering the whole substream.
	for {
		ctyp, clen, err := r.readHeader(r.off)
		if err == io.EOF {
			break
		}
		if err != nil {
			return err
		}
		if ctyp != recContinue {
			break
		}
		r.off += 4
		r.contStarts[len(r.data)] = true
		if err := r.appendPayload(int(clen)); err != nil {
			return err
		}
	}
	return nil
}

func (r *recordReader) appendPayload(length int) error {
	if length == 0 {
		return nil
	}
	if r.off+int64(length) > r.s.Size() {
		return fmt.Errorf("xls: record payload of %d bytes runs past end of stream", length)
	}
	start := len(r.data)
	r.data = append(r.data, make([]byte, length)...)
	if _, err := r.s.ReadAt(r.data[start:start+length], r.off); err != nil {
		return err
	}
	r.off += int64(length)
	return nil
}

// cursor reads primitive values out of a record payload.
type cursor struct {
	data       []byte
	contStarts map[int]bool
	pos        int
}

func newCursor(data []byte, contStarts map[int]bool) *cursor {
	return &cursor{data: data, contStarts: contStarts}
}

func (c *cursor) remaining() int { return len(c.data) - c.pos }

func (c *cursor) atContinuation() bool {
	if c.contStarts == nil {
		return false
	}
	return c.contStarts[c.pos]
}

func (c *cursor) u8() (byte, bool) {
	if c.remaining() < 1 {
		return 0, false
	}
	v := c.data[c.pos]
	c.pos++
	return v, true
}

func (c *cursor) u16() (uint16, bool) {
	if c.remaining() < 2 {
		return 0, false
	}
	v := binary.LittleEndian.Uint16(c.data[c.pos : c.pos+2])
	c.pos += 2
	return v, true
}

func (c *cursor) u32() (uint32, bool) {
	if c.remaining() < 4 {
		return 0, false
	}
	v := binary.LittleEndian.Uint32(c.data[c.pos : c.pos+4])
	c.pos += 4
	return v, true
}

func (c *cursor) f64() (float64, bool) {
	if c.remaining() < 8 {
		return 0, false
	}
	bits := binary.LittleEndian.Uint64(c.data[c.pos : c.pos+8])
	c.pos += 8
	return float64frombits(bits), true
}

func (c *cursor) skip(n int) bool {
	if c.remaining() < n {
		c.pos = len(c.data)
		return false
	}
	c.pos += n
	return true
}
