package numfmt

import (
	"testing"
	"time"
)

func TestIsDateFormat(t *testing.T) {
	cases := []struct {
		format string
		want   bool
		why    string
	}{
		// Plainly dates.
		{"m/d/yy", true, "built-in short date"},
		{"yyyy-mm-dd", true, "iso style"},
		{"[$-409]d-mmm-yy", true, "locale hint must not hide the date tokens"},
		{"h:mm AM/PM", true, "time with meridiem"},
		{"[h]:mm:ss", true, "elapsed hours"},
		{"mmss.0", true, "elapsed-style built-in 47"},
		{"d", true, "a bare day token"},

		// Plainly not dates.
		{"General", false, "general is not a date"},
		{"0.00%", false, "percent"},
		{"#,##0", false, "thousands"},
		{"$#,##0.00", false, "currency"},
		{"0.00E+00", false, "scientific: the E must not read as a token"},
		{"@", false, "text"},
		{"", false, "empty"},

		// The cases that catch a naive implementation.
		{`"Paid on "0`, false, "the d in a quoted literal is not a day"},
		{`\d0`, false, "an escaped d is not a day"},
		{"[Red]#,##0", false, "a colour is not a date"},
		{`_(* #,##0_);_(* \(#,##0\);_(* "-"_);_(@_)`, false, "accounting format"},
		{"[>=100]0.0;[<100]0.00", false, "conditions are not dates"},

		// A date only in the first section is still a date; a date only in a
		// later section is not, because sign does not change a cell's type.
		{"m/d/yy;@", true, "first section decides"},
	}

	for _, c := range cases {
		if got := IsDateFormat(c.format); got != c.want {
			t.Errorf("IsDateFormat(%q) = %v, want %v (%s)", c.format, got, c.want, c.why)
		}
	}
}

func TestSerialToTime(t *testing.T) {
	utc := func(y int, m time.Month, d, hh, mm, ss int) time.Time {
		return time.Date(y, m, d, hh, mm, ss, 0, time.UTC)
	}

	cases := []struct {
		serial   float64
		date1904 bool
		want     time.Time
		why      string
	}{
		{1, false, utc(1900, 1, 1, 0, 0, 0), "serial 1 is the epoch"},
		{59, false, utc(1900, 2, 28, 0, 0, 0), "last day before the phantom"},
		{60, false, utc(1900, 2, 28, 0, 0, 0), "phantom 1900-02-29 folds back"},
		{61, false, utc(1900, 3, 1, 0, 0, 0), "the leap-bug offset kicks in"},
		{25569, false, utc(1970, 1, 1, 0, 0, 0), "unix epoch"},
		{25569.5, false, utc(1970, 1, 1, 12, 0, 0), "fractional day is noon"},
		{44927, false, utc(2023, 1, 1, 0, 0, 0), "a modern date"},
		{0, true, utc(1904, 1, 1, 0, 0, 0), "1904 system epoch"},
		{24107, true, utc(1970, 1, 1, 0, 0, 0), "unix epoch in the 1904 system"},
	}

	for _, c := range cases {
		got := SerialToTime(c.serial, c.date1904)
		if !got.Equal(c.want) {
			t.Errorf("SerialToTime(%v, %v) = %v, want %v (%s)",
				c.serial, c.date1904, got, c.want, c.why)
		}
		if got.Location() != time.UTC {
			t.Errorf("SerialToTime(%v) location = %v, want UTC", c.serial, got.Location())
		}
	}
}

// TestSerialRoundTrip checks that TimeToSerial inverts SerialToTime across the
// leap-bug boundary, which is where an off-by-one would hide.
func TestSerialRoundTrip(t *testing.T) {
	for _, date1904 := range []bool{false, true} {
		for serial := 1.0; serial < 50000; serial += 137 {
			if !date1904 && serial >= 59 && serial <= 60 {
				continue // the phantom day has no inverse
			}
			tm := SerialToTime(serial, date1904)
			back := TimeToSerial(tm, date1904)
			if back != serial {
				t.Fatalf("round trip failed for serial %v (1904=%v): got %v via %v",
					serial, date1904, back, tm)
			}
		}
	}
}

func TestFormatValueDates(t *testing.T) {
	// 44927 is 2023-01-01; adding a fraction gives a time of day.
	const newYear2023 = 44927

	cases := []struct {
		value  float64
		format string
		want   string
	}{
		{newYear2023, "m/d/yyyy", "1/1/2023"},
		{newYear2023, "m/d/yy", "1/1/23"},
		{newYear2023, "yyyy-mm-dd", "2023-01-01"},
		{newYear2023, "[$-409]d-mmm-yy", "1-Jan-23"},
		{newYear2023, "dddd", "Sunday"},
		{newYear2023, "ddd", "Sun"},
		{newYear2023, "mmmm", "January"},
		{newYear2023, "mmmmm", "J"},
		{0.5, "h:mm AM/PM", "12:00 PM"},
		{0.75, "h:mm AM/PM", "6:00 PM"},
		{0.25, "h:mm:ss", "6:00:00"},
		{0.5, "hh:mm", "12:00"},
		{newYear2023 + 0.5, "m/d/yyyy h:mm", "1/1/2023 12:00"},
		// 'm' means minutes after an hour token and months otherwise; the same
		// character in the same format string must resolve both ways.
		{newYear2023 + 0.5, "mm/dd hh:mm", "01/01 12:00"},
		{1.5, "[h]:mm", "36:00"},
		{2.25, "[h]:mm", "54:00"},
	}

	for _, c := range cases {
		if got := FormatValue(c.value, c.format); got != c.want {
			t.Errorf("FormatValue(%v, %q) = %q, want %q", c.value, c.format, got, c.want)
		}
	}
}

func TestFormatValueNumbers(t *testing.T) {
	cases := []struct {
		value  float64
		format string
		want   string
	}{
		{1234.5678, "#,##0.00", "1,234.57"},
		{1234, "#,##0", "1,234"},
		{1234567, "#,##0", "1,234,567"},
		{0.5, "0.00", "0.50"},
		{0.1234, "0.00%", "12.34%"},
		{0.5, "0%", "50%"},
		{-1234.5, "#,##0.00", "-1,234.50"},
		{5, "0", "5"},
		{5.6, "0", "6"},
		{0, "0.00", "0.00"},
		{12345.678, "0.00E+00", "1.23E+04"},
		{12345.678, "##0.0E+0", "12.3E+3"},
		{1234567, "#,##0,", "1,235"},

		// Section selection.
		{-5, "#,##0;(#,##0)", "(5)"},
		{5, "#,##0;(#,##0)", "5"},
		{0, `0.0;-0.0;"zero"`, "zero"},

		// Literals must survive intact.
		{42, `"Qty: "0`, "Qty: 42"},
		{42, `0" units"`, "42 units"},
	}

	for _, c := range cases {
		if got := FormatValue(c.value, c.format); got != c.want {
			t.Errorf("FormatValue(%v, %q) = %q, want %q", c.value, c.format, got, c.want)
		}
	}
}

func TestBuiltin(t *testing.T) {
	if f, ok := Builtin(14); !ok || f != "m/d/yyyy" {
		t.Errorf("Builtin(14) = %q, %v", f, ok)
	}
	if _, ok := Builtin(200); ok {
		t.Error("Builtin(200) should not resolve; customs are file-defined")
	}
	for _, id := range []int{14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47} {
		if !IsDateBuiltin(id) {
			t.Errorf("IsDateBuiltin(%d) = false, want true", id)
		}
	}
	for _, id := range []int{0, 1, 2, 3, 9, 10, 11, 37, 49} {
		if IsDateBuiltin(id) {
			t.Errorf("IsDateBuiltin(%d) = true, want false", id)
		}
	}
	// Every built-in flagged as a date must also parse as one, or the two
	// paths that decide "is this a date" would disagree with each other.
	for id, format := range builtinFormats {
		if IsDateBuiltin(id) && !IsDateFormat(format) {
			t.Errorf("built-in %d (%q) is a date by id but not by format", id, format)
		}
	}
}
