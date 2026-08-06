package numfmt

import (
	"math"
	"time"
)

// Excel stores dates as a floating point "serial" number: the integer part
// counts days from an epoch, the fractional part is the time of day.
//
// There are two epochs. The 1900 system (Windows default) and the 1904 system
// (historic Mac default, still selectable and still found in the wild). Which
// one applies is a property of the workbook, not of the cell, which is why
// every conversion here takes a date1904 flag rather than guessing.
//
// All conversions produce times in UTC. Excel serials carry no timezone and no
// DST rules; attaching a local zone to them would silently shift values for any
// program that later formats them, and would make the same file decode
// differently on different machines. Callers who know the intended zone can
// reinterpret the wall-clock fields themselves.

var (
	// epoch1900 is 1899-12-30, the base used for serials at or after the
	// phantom leap day. See SerialToTime for why there are two bases.
	epoch1900 = time.Date(1899, time.December, 30, 0, 0, 0, 0, time.UTC)
	// epoch1900Early is 1899-12-31, the base used for serials before it.
	epoch1900Early = time.Date(1899, time.December, 31, 0, 0, 0, 0, time.UTC)
	// epoch1904 is the base of the 1904 date system, where serial 0 is the epoch.
	epoch1904 = time.Date(1904, time.January, 1, 0, 0, 0, 0, time.UTC)
)

// SerialToTime converts an Excel date serial to a time.Time in UTC.
//
// In the 1900 system Excel deliberately reproduces a Lotus 1-2-3 bug: it
// believes 1900 was a leap year, so serial 60 denotes a 29th of February that
// never existed. Every serial from 61 onwards is therefore one greater than a
// naive day count would suggest. We handle this by switching the epoch at the
// boundary rather than by adding a correction, which keeps both branches easy
// to check by hand: serial 1 is 1900-01-01 and serial 61 is 1900-03-01.
//
// Serial 60 itself has no correct answer. We map it to 1900-02-28, matching the
// day Excel displays immediately before it, so the value stays inside the real
// calendar instead of silently rolling forward into March.
func SerialToTime(serial float64, date1904 bool) time.Time {
	if math.IsNaN(serial) || math.IsInf(serial, 0) {
		return time.Time{}
	}

	base := epoch1904
	if !date1904 {
		if serial < 60 {
			base = epoch1900Early
		} else {
			base = epoch1900
		}
	}

	days := math.Floor(serial)
	frac := serial - days

	// The fractional day is converted through milliseconds. A float64 has more
	// than enough precision for a day expressed in nanoseconds, but the values
	// that reach us are usually the result of Excel's own decimal rounding, so
	// an exact conversion produces times like 09:59:59.9999997. Rounding to the
	// millisecond removes that noise without losing anything a spreadsheet can
	// actually represent.
	const msPerDay = 24 * 60 * 60 * 1000
	ms := math.Round(frac * msPerDay)

	// Rounding can carry into the next day (a serial of x.9999999 becomes a
	// full day of milliseconds); fold the carry into the day count.
	if ms >= msPerDay {
		ms -= msPerDay
		days++
	}

	t := base.AddDate(0, 0, int(days))
	return t.Add(time.Duration(ms) * time.Millisecond)
}

// TimeToSerial converts a time to an Excel date serial. It is the inverse of
// SerialToTime and is provided mainly so that round trips can be tested.
func TimeToSerial(t time.Time, date1904 bool) float64 {
	t = t.UTC()
	day := time.Date(t.Year(), t.Month(), t.Day(), 0, 0, 0, 0, time.UTC)

	var serial float64
	if date1904 {
		serial = math.Round(day.Sub(epoch1904).Hours() / 24)
	} else {
		serial = math.Round(day.Sub(epoch1900).Hours() / 24)
		// Undo the phantom leap day for dates before 1900-03-01.
		if serial < 61 {
			serial = math.Round(day.Sub(epoch1900Early).Hours() / 24)
		}
	}

	secs := float64(t.Hour()*3600+t.Minute()*60+t.Second()) +
		float64(t.Nanosecond())/1e9
	return serial + secs/86400
}
