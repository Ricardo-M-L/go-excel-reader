// Package numfmt parses Excel number format codes.
//
// It answers the one question the readers cannot work without — is this
// number actually a date? — and, as a standalone utility, renders a float
// through a format code the way Excel would display it.
//
// A spreadsheet cell stores a date as a bare float64. Nothing in the cell says
// "date"; the only signal is the number format attached to it through the
// style table. Getting this wrong is the single most common way a spreadsheet
// reader corrupts data, which is why IsDateFormat is deliberately conservative
// about what counts as a date token and careful about what counts as a literal.
//
// This is a port of ExcelDataReader's Core/NumberFormat.
package numfmt

import (
	"math"
	"strconv"
	"strings"
)

// builtinFormats holds the number formats Excel defines implicitly. A file only
// stores format strings for ids it invented (customs start at 164); anything
// below that is looked up from this table.
//
// The gaps are real: ids 23-36 and 41-44 are reserved or locale-specific and
// have no portable definition, so a file referencing them falls back to General.
var builtinFormats = map[int]string{
	0:  "General",
	1:  "0",
	2:  "0.00",
	3:  "#,##0",
	4:  "#,##0.00",
	5:  "$#,##0_);($#,##0)",
	6:  "$#,##0_);[Red]($#,##0)",
	7:  "$#,##0.00_);($#,##0.00)",
	8:  "$#,##0.00_);[Red]($#,##0.00)",
	9:  "0%",
	10: "0.00%",
	11: "0.00E+00",
	12: "# ?/?",
	13: "# ??/??",
	14: "m/d/yyyy",
	15: "d-mmm-yy",
	16: "d-mmm",
	17: "mmm-yy",
	18: "h:mm AM/PM",
	19: "h:mm:ss AM/PM",
	20: "h:mm",
	21: "h:mm:ss",
	22: "m/d/yyyy h:mm",
	37: "#,##0 ;(#,##0)",
	38: "#,##0 ;[Red](#,##0)",
	39: "#,##0.00;(#,##0.00)",
	40: "#,##0.00;[Red](#,##0.00)",
	45: "mm:ss",
	46: "[h]:mm:ss",
	47: "mmss.0",
	48: "##0.0E+0",
	49: "@",
}

// Builtin returns the format string for a built-in numFmtId.
func Builtin(id int) (string, bool) {
	f, ok := builtinFormats[id]
	return f, ok
}

// builtinDateIDs are the built-in format ids that always denote a date or time.
//
// We answer these from the id alone rather than by parsing the format string,
// because a file is free to omit the string entirely for a built-in id, and
// because locale-dependent built-ins (14 is written "mm-dd-yy" in some locales)
// would otherwise have to be pattern-matched in every locale Excel ships.
var builtinDateIDs = map[int]bool{
	14: true, 15: true, 16: true, 17: true, 18: true, 19: true,
	20: true, 21: true, 22: true, 45: true, 46: true, 47: true,
}

// IsDateBuiltin reports whether a built-in numFmtId always denotes a date or time.
//
// Beyond the well-known ids, two ranges are reserved for locale-specific
// formats and are entirely dates: 27-36 and 50-58 hold the Japanese, Chinese,
// Korean and Thai calendar layouts. ECMA-376 leaves their format strings
// undefined, so a file using one supplies no <numFmt> to inspect and there is
// nothing to pattern-match. A reader that only knows ids 0-49 silently returns
// a raw serial for every date in a CJK-locale workbook.
func IsDateBuiltin(id int) bool {
	if builtinDateIDs[id] {
		return true
	}
	return (id >= 27 && id <= 36) || (id >= 50 && id <= 58)
}

// Excel serial numbers are bounded by the calendar it can represent. The range
// matches OLE Automation dates: the lower bound is year 100 and the upper bound
// is the end of 9999.
const (
	minSerial = -657435.0
	maxSerial = 2958466.0
)

// SerialInRange reports whether a serial denotes a date Excel could represent.
//
// A cell can carry a date format and a value that is not a date -- a stray
// 1.0E12, or a number that acquired a neighbour's formatting. Converting those
// anyway yields confident nonsense: 1.0E12 becomes the year 2,737,908,906.
// Callers should fall back to returning the number.
func SerialInRange(serial float64) bool {
	if math.IsNaN(serial) || math.IsInf(serial, 0) {
		return false
	}
	return serial >= minSerial && serial <= maxSerial
}

// IsDateFormat reports whether a format code renders its value as a date or time.
//
// Only the first section is examined. An Excel format may carry up to four
// sections (positive;negative;zero;text) and a cell's type does not change with
// its sign, so the positive section decides.
func IsDateFormat(format string) bool {
	secs := splitSections(format)
	if len(secs) == 0 {
		return false
	}
	return sectionIsDate(secs[0])
}

// sectionIsDate scans one format section for date or time tokens.
//
// The work here is almost entirely about what to *ignore*. A literal "day" in
// quotes is not a date token; neither is the 'd' in an escaped \d, nor anything
// inside a [Red] colour or a [$-409] locale hint. Skipping those correctly is
// what separates this from a naive strings.ContainsAny, which reports the
// currency format $#,##0 as a date because of the 'd'... it does not, but
// "Paid"#,##0 would fool it, and files in the wild contain exactly that.
func sectionIsDate(section string) bool {
	for i := 0; i < len(section); {
		c := section[i]
		switch c {
		case '"':
			i = skipQuoted(section, i)
		case '\\':
			i += 2
		case '_', '*':
			// _x reserves the width of x; *x repeats x as fill. Either way the
			// following character is not a format token.
			i += 2
		case '[':
			end := strings.IndexByte(section[i:], ']')
			if end < 0 {
				return false
			}
			inner := section[i+1 : i+end]
			// [h], [mm] and [ss] request elapsed time and are date tokens.
			// Everything else in brackets is a colour, a condition or a locale.
			if isElapsed(inner) {
				return true
			}
			i += end + 1
		case 'y', 'Y', 'd', 'D', 'h', 'H', 's', 'S':
			return true
		case 'm', 'M':
			return true
		case 'g', 'G':
			// "General" is not a date, and is the only token starting with g
			// that we recognise; a bare 'g' is an era marker we do not support.
			if hasPrefixFold(section[i:], "general") {
				i += len("general")
				continue
			}
			i++
		case 'a', 'A':
			// AM/PM markers imply a time; a lone 'a' does not.
			if hasPrefixFold(section[i:], "am/pm") || hasPrefixFold(section[i:], "a/p") {
				return true
			}
			i++
		default:
			i++
		}
	}
	return false
}

// isElapsed reports whether a bracket body is an elapsed-time token such as
// [h], [mm] or [ss]. Elapsed tokens use a single repeated unit character.
func isElapsed(inner string) bool {
	if inner == "" {
		return false
	}
	u := lowerByte(inner[0])
	if u != 'h' && u != 'm' && u != 's' {
		return false
	}
	for i := 0; i < len(inner); i++ {
		if lowerByte(inner[i]) != u {
			return false
		}
	}
	return true
}

// splitSections splits a format code on semicolons that are not inside a quoted
// literal, a bracket group or an escape.
func splitSections(format string) []string {
	var (
		out   []string
		start int
	)
	for i := 0; i < len(format); {
		switch format[i] {
		case '"':
			i = skipQuoted(format, i)
		case '\\', '_', '*':
			i += 2
		case '[':
			end := strings.IndexByte(format[i:], ']')
			if end < 0 {
				i++
				continue
			}
			i += end + 1
		case ';':
			out = append(out, format[start:i])
			i++
			start = i
		default:
			i++
		}
	}
	out = append(out, format[start:])
	return out
}

// skipQuoted returns the index just past the quoted literal starting at i.
func skipQuoted(s string, i int) int {
	end := strings.IndexByte(s[i+1:], '"')
	if end < 0 {
		return len(s)
	}
	return i + end + 2
}

func lowerByte(b byte) byte {
	if b >= 'A' && b <= 'Z' {
		return b + 32
	}
	return b
}

func hasPrefixFold(s, prefix string) bool {
	if len(s) < len(prefix) {
		return false
	}
	return strings.EqualFold(s[:len(prefix)], prefix)
}

// FormatValue renders a value through an Excel format code, using the 1900 date
// system. It is a display helper; the readers themselves only need IsDateFormat.
//
// Supported: literals and escapes, digit placeholders (0 # ?), thousands
// grouping and comma scaling, decimals, percent, scientific notation, the text
// placeholder @, colour and locale brackets (ignored), conditional sections,
// elapsed time, and the full set of date and time tokens.
//
// Not supported: fraction formats such as "# ?/?" (rendered as plain decimals)
// and era calendars. These fall back to a readable approximation rather than an
// error, because a display helper that panics on an exotic format is worse than
// one that is occasionally imprecise.
func FormatValue(value float64, fmtString string) string {
	return FormatValueIn(value, fmtString, false)
}

// FormatValueIn is FormatValue with an explicit choice of date system.
func FormatValueIn(value float64, fmtString string, date1904 bool) string {
	if strings.TrimSpace(fmtString) == "" {
		return formatGeneral(value)
	}

	section, negate := selectSection(fmtString, value)
	if section == "" {
		return ""
	}
	if hasPrefixFold(strings.TrimSpace(section), "general") {
		return formatGeneral(value)
	}

	if sectionIsDate(section) {
		return formatDateSection(value, section, date1904)
	}

	// When a dedicated negative section was chosen it supplies its own sign,
	// usually as a literal minus or a pair of parentheses, and the value is
	// rendered as its magnitude. Otherwise a negative value needs the minus
	// prepended here, because no placeholder in the section will produce one.
	v := math.Abs(value)
	addMinus := !negate && value < 0
	return formatNumberSection(v, section, addMinus)
}

// selectSection picks the format section that applies to value and reports
// whether the value should be rendered as its absolute value.
//
// Excel's rules: with one section it applies to everything and negatives get a
// leading minus. With two, the second handles negatives. With three or more,
// the third handles zero. Sections carrying an explicit [condition] are tested
// in order and the last unconditional section is the fallback.
func selectSection(format string, value float64) (section string, abs bool) {
	secs := splitSections(format)
	if len(secs) == 0 {
		return "", false
	}

	if cond, ok := selectConditional(secs, value); ok {
		return cond, true
	}

	switch {
	case len(secs) == 1:
		return secs[0], false
	case value < 0 && len(secs) >= 2:
		return secs[1], true
	case value == 0 && len(secs) >= 3:
		return secs[2], false
	default:
		return secs[0], false
	}
}

// selectConditional handles sections prefixed with a comparison such as [>=100].
func selectConditional(secs []string, value float64) (string, bool) {
	var (
		any      bool
		fallback string
		hasFall  bool
	)
	for _, s := range secs {
		op, operand, rest, ok := parseCondition(s)
		if !ok {
			if !hasFall {
				fallback, hasFall = s, true
			}
			continue
		}
		any = true
		if compare(value, op, operand) {
			return rest, true
		}
	}
	if any && hasFall {
		return fallback, true
	}
	return "", false
}

// parseCondition extracts a leading [<op><number>] guard from a section.
func parseCondition(s string) (op string, operand float64, rest string, ok bool) {
	t := strings.TrimSpace(s)
	if !strings.HasPrefix(t, "[") {
		return "", 0, "", false
	}
	end := strings.IndexByte(t, ']')
	if end < 0 {
		return "", 0, "", false
	}
	inner := t[1:end]
	for _, o := range []string{">=", "<=", "<>", ">", "<", "="} {
		if strings.HasPrefix(inner, o) {
			n, err := strconv.ParseFloat(strings.TrimSpace(inner[len(o):]), 64)
			if err != nil {
				return "", 0, "", false
			}
			return o, n, t[end+1:], true
		}
	}
	return "", 0, "", false
}

func compare(v float64, op string, operand float64) bool {
	switch op {
	case ">=":
		return v >= operand
	case "<=":
		return v <= operand
	case "<>":
		return v != operand
	case ">":
		return v > operand
	case "<":
		return v < operand
	case "=":
		return v == operand
	}
	return false
}

// formatGeneral approximates Excel's General format, which shows up to about
// eleven significant digits and switches to scientific notation at the extremes.
func formatGeneral(v float64) string {
	if v == 0 {
		return "0"
	}
	a := math.Abs(v)
	if a >= 1e11 || a < 1e-10 {
		s := strconv.FormatFloat(v, 'E', 5, 64)
		return s
	}
	return strconv.FormatFloat(v, 'f', -1, 64)
}

// --- date rendering ---

// dtok is one token of a parsed date format section.
type dtok struct {
	kind byte   // 'y' 'm' 'd' 'h' 's' 'a' 'e' 'f' 'L'
	n    int    // repeat count for token characters
	lit  string // literal text, for kind 'L' and the raw AM/PM spelling
	unit byte   // elapsed unit, for kind 'e'
}

func formatDateSection(serial float64, section string, date1904 bool) string {
	t := SerialToTime(serial, date1904)
	toks := tokenizeDate(section)
	resolveMinutes(toks)

	ampm := false
	for _, tk := range toks {
		if tk.kind == 'a' {
			ampm = true
			break
		}
	}

	var b strings.Builder
	for _, tk := range toks {
		switch tk.kind {
		case 'L':
			b.WriteString(tk.lit)
		case 'y':
			if tk.n <= 2 {
				b.WriteString(pad2(t.Year() % 100))
			} else {
				b.WriteString(padN(t.Year(), 4))
			}
		case 'm':
			switch {
			case tk.n == 1:
				b.WriteString(strconv.Itoa(int(t.Month())))
			case tk.n == 2:
				b.WriteString(pad2(int(t.Month())))
			case tk.n == 3:
				b.WriteString(t.Format("Jan"))
			case tk.n == 4:
				b.WriteString(t.Format("January"))
			default:
				b.WriteString(t.Format("January")[:1])
			}
		case 'n': // minutes, produced by resolveMinutes
			if tk.n == 1 {
				b.WriteString(strconv.Itoa(t.Minute()))
			} else {
				b.WriteString(pad2(t.Minute()))
			}
		case 'd':
			switch {
			case tk.n == 1:
				b.WriteString(strconv.Itoa(t.Day()))
			case tk.n == 2:
				b.WriteString(pad2(t.Day()))
			case tk.n == 3:
				b.WriteString(t.Format("Mon"))
			default:
				b.WriteString(t.Format("Monday"))
			}
		case 'h':
			h := t.Hour()
			if ampm {
				h %= 12
				if h == 0 {
					h = 12
				}
			}
			if tk.n == 1 {
				b.WriteString(strconv.Itoa(h))
			} else {
				b.WriteString(pad2(h))
			}
		case 's':
			if tk.n == 1 {
				b.WriteString(strconv.Itoa(t.Second()))
			} else {
				b.WriteString(pad2(t.Second()))
			}
		case 'f':
			frac := float64(t.Nanosecond()) / 1e9
			s := strconv.FormatFloat(frac, 'f', tk.n, 64)
			b.WriteString(s[1:]) // keep ".xxx", drop the leading zero
		case 'a':
			b.WriteString(meridiem(t.Hour(), tk.lit))
		case 'e':
			b.WriteString(padN(int(elapsed(serial, tk.unit)), tk.n))
		}
	}
	return b.String()
}

func elapsed(serial float64, unit byte) int64 {
	switch unit {
	case 'h':
		return int64(math.Floor(serial * 24))
	case 'm':
		return int64(math.Floor(serial * 24 * 60))
	default:
		return int64(math.Floor(serial * 24 * 60 * 60))
	}
}

// meridiem renders an AM/PM marker in the same style as the format code used,
// so "am/pm" stays lowercase and "A/P" stays abbreviated.
func meridiem(hour int, spelling string) string {
	pm := hour >= 12
	short := strings.Contains(spelling, "/") && len(spelling) <= 3
	var s string
	switch {
	case short && pm:
		s = "P"
	case short:
		s = "A"
	case pm:
		s = "PM"
	default:
		s = "AM"
	}
	if spelling != "" && spelling == strings.ToLower(spelling) {
		return strings.ToLower(s)
	}
	return s
}

func tokenizeDate(section string) []dtok {
	var toks []dtok
	push := func(t dtok) { toks = append(toks, t) }
	lit := func(s string) { push(dtok{kind: 'L', lit: s}) }

	for i := 0; i < len(section); {
		c := section[i]
		switch {
		case c == '"':
			j := skipQuoted(section, i)
			if j-1 <= len(section) && j-1 > i+1 {
				lit(section[i+1 : j-1])
			}
			i = j
		case c == '\\':
			if i+1 < len(section) {
				lit(string(section[i+1]))
			}
			i += 2
		case c == '_':
			// Reserve the width of the next character; a space is the closest
			// we can get without knowing the font metrics.
			lit(" ")
			i += 2
		case c == '*':
			i += 2
		case c == '[':
			end := strings.IndexByte(section[i:], ']')
			if end < 0 {
				i++
				continue
			}
			inner := section[i+1 : i+end]
			switch {
			case isElapsed(inner):
				push(dtok{kind: 'e', unit: lowerByte(inner[0]), n: len(inner)})
			case strings.HasPrefix(inner, "$"):
				lit(currencySymbol(inner))
			}
			i += end + 1
		case lowerByte(c) == 'y':
			n := runLen(section, i)
			push(dtok{kind: 'y', n: n})
			i += n
		case lowerByte(c) == 'm':
			n := runLen(section, i)
			push(dtok{kind: 'm', n: n})
			i += n
		case lowerByte(c) == 'd':
			n := runLen(section, i)
			push(dtok{kind: 'd', n: n})
			i += n
		case lowerByte(c) == 'h':
			n := runLen(section, i)
			push(dtok{kind: 'h', n: n})
			i += n
		case lowerByte(c) == 's':
			n := runLen(section, i)
			push(dtok{kind: 's', n: n})
			i += n
		case lowerByte(c) == 'a' && (hasPrefixFold(section[i:], "am/pm") || hasPrefixFold(section[i:], "a/p")):
			n := 3
			if hasPrefixFold(section[i:], "am/pm") {
				n = 5
			}
			push(dtok{kind: 'a', lit: section[i : i+n]})
			i += n
		case c == '.' && isDigitRun(section, i+1):
			n := runLenOf(section, i+1, '0')
			push(dtok{kind: 'f', n: n})
			i += 1 + n
		default:
			lit(string(c))
			i++
		}
	}
	return toks
}

// currencySymbol extracts the symbol from a [$SYM-409] style bracket.
func currencySymbol(inner string) string {
	body := strings.TrimPrefix(inner, "$")
	if dash := strings.IndexByte(body, '-'); dash >= 0 {
		body = body[:dash]
	}
	return body
}

func runLen(s string, i int) int { return runLenOf(s, i, lowerByte(s[i])) }

func runLenOf(s string, i int, ch byte) int {
	n := 0
	for i+n < len(s) && lowerByte(s[i+n]) == ch {
		n++
	}
	return n
}

func isDigitRun(s string, i int) bool { return i < len(s) && s[i] == '0' }

// resolveMinutes rewrites 'm' tokens that mean minutes into 'n' tokens.
//
// Excel overloads 'm' for both month and minute and disambiguates by context:
// an 'm' is a minute when it directly follows an hour token or directly
// precedes a seconds token. Three or more 'm's are always a month name, since
// there is no such thing as a minute name.
func resolveMinutes(toks []dtok) {
	meaningful := func(k byte) bool {
		return k == 'y' || k == 'm' || k == 'd' || k == 'h' || k == 's' || k == 'e'
	}
	for i := range toks {
		if toks[i].kind != 'm' || toks[i].n >= 3 {
			continue
		}
		// Look back for the nearest meaningful token.
		for j := i - 1; j >= 0; j-- {
			if !meaningful(toks[j].kind) {
				continue
			}
			if toks[j].kind == 'h' || (toks[j].kind == 'e' && toks[j].unit == 'h') {
				toks[i].kind = 'n'
			}
			break
		}
		if toks[i].kind == 'n' {
			continue
		}
		// Otherwise look forward: an 'm' immediately before seconds is a minute.
		for j := i + 1; j < len(toks); j++ {
			if !meaningful(toks[j].kind) {
				continue
			}
			if toks[j].kind == 's' {
				toks[i].kind = 'n'
			}
			break
		}
	}
}

func pad2(v int) string {
	if v < 10 && v >= 0 {
		return "0" + strconv.Itoa(v)
	}
	return strconv.Itoa(v)
}

func padN(v, n int) string {
	s := strconv.Itoa(v)
	for len(s) < n {
		s = "0" + s
	}
	return s
}

// --- number rendering ---

func formatNumberSection(v float64, section string, addMinus bool) string {
	toks, spec := tokenizeNumber(section)

	value := v
	for i := 0; i < spec.percent; i++ {
		value *= 100
	}
	for i := 0; i < spec.scale; i++ {
		value /= 1000
	}

	var intStr, decStr, expStr string
	if spec.hasExp {
		mant, exp := mantissaExponent(value, maxInt(spec.intPlaces, 1))
		intStr, decStr = splitDigits(mant, spec)
		sign := "-"
		if exp >= 0 {
			sign = "+"
			if !spec.expPlus {
				sign = ""
			}
		}
		expStr = sign + padN(abs(exp), spec.expDigits)
	} else {
		intStr, decStr = splitDigits(value, spec)
	}

	if spec.group {
		intStr = groupThousands(intStr)
	}

	var b strings.Builder
	if addMinus && (value != 0 || decStr != "") {
		b.WriteByte('-')
	}

	intDone, decDone := false, false
	for _, tk := range toks {
		switch tk.kind {
		case 'L':
			b.WriteString(tk.lit)
		case 'd':
			if tk.afterDecimal {
				if !decDone {
					b.WriteString(decStr)
					decDone = true
				}
			} else if !intDone {
				b.WriteString(intStr)
				intDone = true
			}
		case '.':
			if decStr != "" {
				b.WriteByte('.')
			}
		case 'E':
			b.WriteByte('E')
			b.WriteString(expStr)
		}
	}
	return b.String()
}

type numSpec struct {
	minInt int
	// intPlaces counts every integer placeholder, required or optional. It
	// drives engineering notation: "##0.0E+0" asks for exponents in steps of
	// three, and the only thing that says so is the three placeholders.
	intPlaces  int
	minDec     int
	maxDec     int
	group      bool
	scale      int
	percent    int
	hasExp     bool
	expDigits  int
	expPlus    bool
	intPadChar byte
}

// tokenizeNumber walks a numeric section once, collecting the literals in order
// and the placeholder counts that determine how the number itself is rendered.
func tokenizeNumber(section string) ([]ntokPos, numSpec) {
	var (
		toks         []ntokPos
		spec         numSpec
		afterDot     bool
		lastDigitIdx = -1
	)
	spec.intPadChar = '0'

	lit := func(s string) { toks = append(toks, ntokPos{kind: 'L', lit: s}) }

	for i := 0; i < len(section); {
		c := section[i]
		switch {
		case c == '"':
			j := skipQuoted(section, i)
			if j-1 > i+1 {
				lit(section[i+1 : j-1])
			}
			i = j
		case c == '\\':
			if i+1 < len(section) {
				lit(string(section[i+1]))
			}
			i += 2
		case c == '_':
			lit(" ")
			i += 2
		case c == '*':
			i += 2
		case c == '[':
			end := strings.IndexByte(section[i:], ']')
			if end < 0 {
				i++
				continue
			}
			inner := section[i+1 : i+end]
			if strings.HasPrefix(inner, "$") {
				lit(currencySymbol(inner))
			}
			// Colours and conditions contribute nothing to the rendering.
			i += end + 1
		case c == '0' || c == '#' || c == '?':
			toks = append(toks, ntokPos{kind: 'd', ch: c, afterDecimal: afterDot})
			lastDigitIdx = len(toks) - 1
			if afterDot {
				spec.maxDec++
				if c != '#' {
					spec.minDec++
				}
			} else {
				spec.intPlaces++
				if c != '#' {
					spec.minInt++
					if c == '?' {
						spec.intPadChar = ' '
					}
				}
			}
			i++
		case c == '.':
			if afterDot {
				lit(".")
				i++
				continue
			}
			afterDot = true
			toks = append(toks, ntokPos{kind: '.'})
			i++
		case c == ',':
			// A comma between digit placeholders turns on thousands grouping;
			// a comma trailing the integer placeholders scales by 1000 each.
			if lastDigitIdx >= 0 && !afterDot && isTrailingComma(section, i) {
				spec.scale++
			} else if lastDigitIdx >= 0 {
				spec.group = true
			} else {
				lit(",")
			}
			i++
		case c == '%':
			spec.percent++
			lit("%")
			i++
		case (c == 'E' || c == 'e') && i+1 < len(section) && (section[i+1] == '+' || section[i+1] == '-'):
			n := runLenOf(section, i+2, '0')
			if n == 0 {
				lit(string(c))
				i++
				continue
			}
			spec.hasExp = true
			spec.expDigits = n
			spec.expPlus = section[i+1] == '+'
			toks = append(toks, ntokPos{kind: 'E'})
			i += 2 + n
		default:
			lit(string(c))
			i++
		}
	}
	return toks, spec
}

type ntokPos struct {
	kind         byte
	ch           byte
	lit          string
	afterDecimal bool
}

// isTrailingComma reports whether the comma at i is a scaling comma, meaning no
// further digit placeholder follows it before the decimal point or section end.
func isTrailingComma(section string, i int) bool {
	for j := i + 1; j < len(section); j++ {
		switch section[j] {
		case '0', '#', '?':
			return false
		case ',':
			continue
		default:
			return true
		}
	}
	return true
}

// splitDigits rounds the value to the section's decimal count and returns its
// integer and fractional digit strings, padded to the required minimums.
func splitDigits(v float64, spec numSpec) (string, string) {
	a := math.Abs(v)
	s := strconv.FormatFloat(a, 'f', spec.maxDec, 64)

	intPart, decPart, _ := strings.Cut(s, ".")

	for len(intPart) < spec.minInt {
		intPart = string(spec.intPadChar) + intPart
	}
	if spec.minInt == 0 && intPart == "0" {
		// "#.##" style formats drop a lone leading zero.
		intPart = ""
	}

	// Trim optional trailing decimal positions, keeping the required minimum.
	for len(decPart) > spec.minDec && strings.HasSuffix(decPart, "0") {
		decPart = decPart[:len(decPart)-1]
	}
	return intPart, decPart
}

func groupThousands(s string) string {
	// Padding characters may sit to the left of the digits; group only digits.
	start := 0
	for start < len(s) && (s[start] == ' ') {
		start++
	}
	head, digits := s[:start], s[start:]
	if len(digits) <= 3 {
		return s
	}
	var parts []string
	for len(digits) > 3 {
		parts = append([]string{digits[len(digits)-3:]}, parts...)
		digits = digits[:len(digits)-3]
	}
	parts = append([]string{digits}, parts...)
	return head + strings.Join(parts, ",")
}

// mantissaExponent normalises v so its mantissa carries intDigits digits before
// the decimal point, which is what "0.00E+00" and "##0.0E+0" ask for.
func mantissaExponent(v float64, intDigits int) (float64, int) {
	if v == 0 {
		return 0, 0
	}
	exp := int(math.Floor(math.Log10(math.Abs(v))))
	// Round the exponent down to a multiple of intDigits so engineering-style
	// formats such as ##0.0E+0 step in threes.
	if intDigits > 1 {
		exp = int(math.Floor(float64(exp)/float64(intDigits))) * intDigits
	}
	return v / math.Pow(10, float64(exp)), exp
}

func abs(i int) int {
	if i < 0 {
		return -i
	}
	return i
}

func maxInt(a, b int) int {
	if a > b {
		return a
	}
	return b
}
