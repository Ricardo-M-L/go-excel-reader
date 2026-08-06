// Dumps a spreadsheet to the same canonical text form as edrdump, so the two
// can be diffed. See edrdump/Program.cs for the format.
package main

import (
	"bufio"
	"fmt"
	"math"
	"os"
	"strings"
	"time"

	excel "github.com/jloor/go-excel-reader"
)

func main() {
	if len(os.Args) < 2 {
		fmt.Fprintln(os.Stderr, "usage: godump <file> [...]")
		os.Exit(1)
	}
	out := bufio.NewWriter(os.Stdout)
	defer out.Flush()

	for _, path := range os.Args[1:] {
		if err := dump(out, path); err != nil {
			fmt.Fprintf(os.Stderr, "%s: %v\n", path, err)
			os.Exit(1)
		}
	}
}

func dump(out *bufio.Writer, path string) error {
	r, err := excel.OpenFile(path)
	if err != nil {
		return err
	}
	defer r.Close()

	sheetIndex := 0
	for r.NextSheet() {
		name := r.SheetName()
		rowIndex := 0
		for r.Read() {
			for col := 0; col < r.FieldCount(); col++ {
				v := r.GetValue(col)
				if v == nil {
					continue
				}
				tag, text := render(v)
				fmt.Fprintf(out, "%d\t%s\t%d\t%d\t%s\t%s\n",
					sheetIndex, escape(name), rowIndex, col, tag, text)
			}
			rowIndex++
		}
		sheetIndex++
	}
	if err := r.Err(); err != nil {
		return err
	}
	return nil
}

func render(v any) (tag, text string) {
	switch x := v.(type) {
	case string:
		return "S", escape(x)
	case bool:
		if x {
			return "B", "true"
		}
		return "B", "false"
	case time.Time:
		return "D", x.UTC().Format("2006-01-02 15:04:05.000")
	case float64:
		return "N", fmt.Sprintf("%016X", math.Float64bits(x))
	case int64:
		return "N", fmt.Sprintf("%016X", math.Float64bits(float64(x)))
	default:
		return "S", escape(fmt.Sprint(x))
	}
}

var escaper = strings.NewReplacer("\\", "\\\\", "\t", "\\t", "\r", "\\r", "\n", "\\n")

func escape(s string) string { return escaper.Replace(s) }
