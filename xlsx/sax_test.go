package xlsx

import "testing"

// TestColumnIndex covers the bijective base-26 conversion of column letters.
//
// The letters have no zero digit: A..Z are 1..26 and AA is 27, which then maps
// to zero-based index 26. Treating them as ordinary base-26 puts AA at 27 and
// shifts every value past column Z one place to the right.
func TestColumnIndex(t *testing.T) {
	cases := []struct {
		ref  string
		want int
		ok   bool
	}{
		{"A1", 0, true},
		{"B1", 1, true},
		{"Z1", 25, true},
		{"AA1", 26, true},
		{"AB1", 27, true},
		{"AZ1", 51, true},
		{"BA1", 52, true},
		{"ZZ1", 701, true},
		{"AAA1", 702, true},
		// XFD is the last column Excel allows.
		{"XFD1048576", 16383, true},

		// A reference without a row number is still a valid column.
		{"C", 2, true},
		// Lowercase appears in files written by non-Excel producers.
		{"aa1", 26, true},

		{"", 0, false},
		{"1", 0, false},
		{"123", 0, false},
	}

	for _, c := range cases {
		got, ok := columnIndex(c.ref)
		if ok != c.ok {
			t.Errorf("columnIndex(%q) ok = %v, want %v", c.ref, ok, c.ok)
			continue
		}
		if ok && got != c.want {
			t.Errorf("columnIndex(%q) = %d, want %d", c.ref, got, c.want)
		}
	}
}

func TestIsIntegerLiteral(t *testing.T) {
	cases := map[string]bool{
		"0":     true,
		"42":    true,
		"-7":    true,
		"+7":    true,
		"":      false,
		"3.0":   false, // a written decimal point means the author meant a decimal
		"1e5":   false,
		"1.5":   false,
		"abc":   false,
		"4-2":   false,
		"1 000": false,
	}
	for in, want := range cases {
		if got := isIntegerLiteral(in); got != want {
			t.Errorf("isIntegerLiteral(%q) = %v, want %v", in, got, want)
		}
	}
}
