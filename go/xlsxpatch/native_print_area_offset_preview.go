package xlsxpatch

import (
	"regexp"
	"strconv"
	"strings"
)

var nativePrintOffsetInteger = regexp.MustCompile(`^[+-]?[0-9]{1,7}$`)

// Closed reference arithmetic, not a workbook calculation engine. Only a single
// OFFSET with literal integers and an absolute same-sheet reference is accepted.
// https://support.microsoft.com/en-us/Excel/functions/offset-function
func parseNativePrintOffset(text, sheetName string) *NativePrintAreaRectV1 {
	// Keep the verbatim source expression safe and bounded in warning transport.
	if len(text) > 2048 || strings.ContainsAny(text, "\r\n\t") || !strings.HasPrefix(text, "OFFSET(") || !strings.HasSuffix(text, ")") {
		return nil
	}
	for _, c := range text {
		if c < 32 || c == 127 {
			return nil
		}
	}
	args := splitNativePrintUnion(text[7:len(text)-1], 5)
	if len(args) < 3 || len(args) > 5 {
		return nil
	}
	base := parseNativePrintAreaRect(strings.TrimSpace(args[0]), sheetName)
	if base == nil {
		return nil
	}
	values := []int{0, 0, base.EndRow - base.Row + 1, base.EndColumn - base.Column + 1}
	for i, arg := range args[1:] {
		arg = strings.Trim(arg, " ")
		if !nativePrintOffsetInteger.MatchString(arg) {
			return nil
		}
		n, err := strconv.Atoi(arg)
		if err != nil || n < -1048576 || n > 1048576 {
			return nil
		}
		values[i] = n
	}
	row, col := base.Row+values[0], base.Column+values[1]
	height, width := values[2], values[3]
	if row < 0 || col < 0 || height <= 0 || width <= 0 || row+height > 1048576 || col+width > 16384 {
		return nil
	}
	return &NativePrintAreaRectV1{Row: row, Column: col, EndRow: row + height - 1, EndColumn: col + width - 1}
}
