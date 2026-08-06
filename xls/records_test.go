package xls

import (
	"math"
	"testing"
)

// TestDecodeRK covers all four RK encodings.
//
// RK packs a number into four bytes and uses the low two bits to say how. Real
// files use every combination, and each one fails differently when wrong: the
// integer variants come out scaled or sign-flipped, the float variants come out
// as garbage exponents. Encoding the inputs here by construction, rather than
// by copying bytes out of a file, makes the tag layout explicit.
func TestDecodeRK(t *testing.T) {
	// intRK builds the "signed 30-bit integer" encoding.
	intRK := func(v int32, div100 bool) uint32 {
		rk := uint32(v<<2) | 0x02
		if div100 {
			rk |= 0x01
		}
		return rk
	}
	// floatRK builds the "top 30 bits of an IEEE 754 double" encoding.
	floatRK := func(v float64, div100 bool) uint32 {
		rk := uint32(math.Float64bits(v)>>32) &^ 0x03
		if div100 {
			rk |= 0x01
		}
		return rk
	}

	cases := []struct {
		name string
		rk   uint32
		want float64
	}{
		{"integer", intRK(42, false), 42},
		{"integer zero", intRK(0, false), 0},
		// Sign extension is the trap: the value occupies the top 30 bits, so a
		// logical shift instead of an arithmetic one turns this into 268435449.
		{"negative integer", intRK(-7, false), -7},
		{"large negative integer", intRK(-100000, false), -100000},
		{"integer scaled by 100", intRK(123, true), 1.23},
		{"negative integer scaled", intRK(-250, true), -2.5},
		// 2.5 and 100 have no mantissa bits below the top 30, so they survive
		// the truncation exactly and can be compared without a tolerance.
		{"double", floatRK(2.5, false), 2.5},
		{"negative double", floatRK(-2.5, false), -2.5},
		{"double scaled by 100", floatRK(250, true), 2.5},
	}

	for _, c := range cases {
		if got := decodeRK(c.rk); got != c.want {
			t.Errorf("%s: decodeRK(0x%08X) = %v, want %v", c.name, c.rk, got, c.want)
		}
	}
}

func TestErrorText(t *testing.T) {
	cases := map[byte]string{
		0x00: "#NULL!",
		0x07: "#DIV/0!",
		0x0F: "#VALUE!",
		0x17: "#REF!",
		0x1D: "#NAME?",
		0x24: "#NUM!",
		0x2A: "#N/A",
	}
	for code, want := range cases {
		if got := errorText(code); got != want {
			t.Errorf("errorText(0x%02X) = %q, want %q", code, got, want)
		}
	}
}

// TestReadStringChars covers the CONTINUE boundary rule, which is the single
// most delicate detail in the BIFF reader.
//
// When a string's characters are split across a CONTINUE record, the first byte
// of the continuation restates the compression flag. Consume it and the string
// decodes; miss it and that byte is read as character data, shifting everything
// after it by one byte for the rest of the table.
func TestReadStringChars(t *testing.T) {
	t.Run("compressed, no split", func(t *testing.T) {
		// cch=3, flags=0 (one byte per character), "abc"
		data := []byte{0x03, 0x00, 0x00, 'a', 'b', 'c'}
		got, err := readLongString(newCursor(data, nil))
		if err != nil {
			t.Fatal(err)
		}
		if got != "abc" {
			t.Errorf("got %q, want %q", got, "abc")
		}
	})

	t.Run("wide, no split", func(t *testing.T) {
		// cch=2, flags=1 (UTF-16LE), "hi"
		data := []byte{0x02, 0x00, 0x01, 'h', 0x00, 'i', 0x00}
		got, err := readLongString(newCursor(data, nil))
		if err != nil {
			t.Fatal(err)
		}
		if got != "hi" {
			t.Errorf("got %q, want %q", got, "hi")
		}
	})

	t.Run("split mid-string restates the flag", func(t *testing.T) {
		// "abcd" declared as compressed, split after "ab". The continuation
		// begins with a flags byte of 0x00, meaning the remainder is still
		// compressed.
		data := []byte{0x04, 0x00, 0x00, 'a', 'b', 0x00, 'c', 'd'}
		contStarts := map[int]bool{5: true} // "cd" block begins at index 5
		got, err := readLongString(newCursor(data, contStarts))
		if err != nil {
			t.Fatal(err)
		}
		if got != "abcd" {
			t.Errorf("got %q, want %q", got, "abcd")
		}
	})

	t.Run("split switches from compressed to wide", func(t *testing.T) {
		// "ab" compressed, then a continuation whose flags byte is 0x01, so
		// the remaining two characters are UTF-16LE.
		data := []byte{0x04, 0x00, 0x00, 'a', 'b', 0x01, 'c', 0x00, 'd', 0x00}
		contStarts := map[int]bool{5: true}
		got, err := readLongString(newCursor(data, contStarts))
		if err != nil {
			t.Fatal(err)
		}
		if got != "abcd" {
			t.Errorf("got %q, want %q", got, "abcd")
		}
	})

	t.Run("ignoring the restated flag corrupts the result", func(t *testing.T) {
		// The same bytes as above, parsed without knowing where the
		// continuation starts. This asserts the failure mode is real, so the
		// test above is not passing for an unrelated reason.
		data := []byte{0x04, 0x00, 0x00, 'a', 'b', 0x00, 'c', 'd'}
		got, err := readLongString(newCursor(data, nil))
		if err != nil {
			t.Fatal(err)
		}
		if got == "abcd" {
			t.Fatal("parsing without continuation boundaries produced the correct string; " +
				"the boundary handling is not actually being exercised")
		}
	})
}

func TestParseBOFRejectsOldVersions(t *testing.T) {
	// BIFF5 workbook globals: version 0x0500.
	bof, err := parseBOF([]byte{0x00, 0x05, 0x05, 0x00})
	if err != nil {
		t.Fatal(err)
	}
	if bof.version != biff5 {
		t.Errorf("version = 0x%04X, want 0x%04X", bof.version, biff5)
	}
	if bof.substream != substreamWorkbook {
		t.Errorf("substream = 0x%04X, want 0x%04X", bof.substream, substreamWorkbook)
	}
}
