package xlsxpatch

import (
	"fmt"
	"regexp"
	"strconv"
	"strings"
)

var structureCellToken = regexp.MustCompile(`^(\$?)([A-Za-z]{1,3})(\$?)([1-9][0-9]*)$`)
var structureColumnToken = regexp.MustCompile(`^(\$?)([A-Za-z]{1,3})$`)
var structureRowToken = regexp.MustCompile(`^(\$?)([1-9][0-9]*)$`)

type structureRef struct {
	row, column       int
	kind              string
	absRow, absColumn string
}

func parseStructureRef(token string) (structureRef, bool) {
	if parts := structureCellToken.FindStringSubmatch(token); parts != nil {
		row, err := strconv.Atoi(parts[4])
		column := 0
		for _, ch := range strings.ToUpper(parts[2]) {
			column = column*26 + int(ch-'A'+1)
		}
		if err != nil || row > excelMaxRows || column > excelMaxColumns {
			return structureRef{}, false
		}
		return structureRef{row - 1, column - 1, "cell", parts[3], parts[1]}, true
	}
	if parts := structureColumnToken.FindStringSubmatch(token); parts != nil {
		column := 0
		for _, ch := range strings.ToUpper(parts[2]) {
			column = column*26 + int(ch-'A'+1)
		}
		if column <= excelMaxColumns {
			return structureRef{0, column - 1, "column", "", parts[1]}, true
		}
	}
	if parts := structureRowToken.FindStringSubmatch(token); parts != nil {
		row, err := strconv.Atoi(parts[2])
		if err == nil && row <= excelMaxRows {
			return structureRef{row - 1, 0, "row", parts[1], ""}, true
		}
	}
	return structureRef{}, false
}
func (r structureRef) text() string {
	name := cellReference(0, r.column)
	name = strings.TrimSuffix(name, "1")
	if r.kind == "column" {
		return r.absColumn + name
	}
	if r.kind == "row" {
		return r.absRow + strconv.Itoa(r.row+1)
	}
	return r.absColumn + name + r.absRow + strconv.Itoa(r.row+1)
}
func structureIdentifier(c byte) bool {
	return c > 127 || c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z' || c >= '0' && c <= '9' || c == '_' || c == '.' || c == '$' || c == '\\'
}
func structureTokenEnd(s string, start int) int {
	for start < len(s) && structureIdentifier(s[start]) {
		start++
	}
	return start
}

// Tokenize strings and qualified references instead of regexp-replacing text:
// quoted literals, function names (LOG10), names and numbers are not addresses.
func shiftStructureFormula(formula, currentSheet, targetSheet string, m StructureMutation) (string, error) {
	var out strings.Builder
	refuse := func() (string, error) {
		return "", fmt.Errorf("xlsxpatch: structure: formula uses unsupported external, structured, 3D, spill, or reference syntax")
	}
	for i := 0; i < len(formula); {
		start := i
		if formula[i] == '"' {
			i++
			closed := false
			for i < len(formula) {
				if formula[i] == '"' {
					i++
					if i < len(formula) && formula[i] == '"' {
						i++
						continue
					}
					closed = true
					break
				}
				i++
			}
			if !closed {
				return refuse()
			}
			out.WriteString(formula[start:i])
			continue
		}
		if formula[i] == '[' || formula[i] == ']' || formula[i] == ':' {
			return refuse()
		}
		if formula[i] == '#' {
			matched := false
			for _, value := range []string{"#REF!", "#VALUE!", "#DIV/0!", "#N/A", "#NAME?", "#NUM!", "#NULL!"} {
				if strings.HasPrefix(formula[i:], value) {
					out.WriteString(value)
					i += len(value)
					matched = true
					break
				}
			}
			if !matched {
				return refuse()
			}
			continue
		}
		sheetName := currentSheet
		prefix := ""
		if formula[i] == '\'' {
			i++
			var sheet strings.Builder
			closed := false
			for i < len(formula) {
				if formula[i] == '\'' {
					i++
					if i < len(formula) && formula[i] == '\'' {
						sheet.WriteByte('\'')
						i++
						continue
					}
					closed = true
					break
				}
				sheet.WriteByte(formula[i])
				i++
			}
			if !closed || i >= len(formula) || formula[i] != '!' {
				return refuse()
			}
			i++
			sheetName = sheet.String()
			if strings.ContainsAny(sheetName, "[]:") {
				return refuse()
			}
			prefix = formula[start:i]
		} else if structureIdentifier(formula[i]) {
			end := structureTokenEnd(formula, i)
			if end < len(formula) && formula[end] == '!' {
				sheetName = formula[i:end]
				i = end + 1
				prefix = formula[start:i]
			}
		} else {
			if formula[i] == '!' {
				return refuse()
			}
			out.WriteByte(formula[i])
			i++
			continue
		}
		tokenStart := i
		end := structureTokenEnd(formula, i)
		if end == i {
			return refuse()
		}
		first, ok := parseStructureRef(formula[i:end])
		last := first
		rangeRef := false
		i = end
		if i < len(formula) && formula[i] == ':' {
			i++
			end = structureTokenEnd(formula, i)
			var lastOK bool
			last, lastOK = parseStructureRef(formula[i:end])
			if !ok || !lastOK || first.kind != last.kind {
				return refuse()
			}
			i = end
			rangeRef = true
			if i < len(formula) && formula[i] == '!' {
				return refuse()
			}
		}
		// Bare column names, row numbers and function identifiers are not references.
		if !ok || (!rangeRef && first.kind != "cell") || (i < len(formula) && formula[i] == '(') {
			out.WriteString(prefix + formula[tokenStart:i])
			continue
		}
		if !strings.EqualFold(sheetName, targetSheet) {
			out.WriteString(prefix + formula[tokenStart:i])
			continue
		}
		axisApplies := first.kind == "cell" || (m.rows() && first.kind == "row") || (!m.rows() && first.kind == "column")
		if !axisApplies {
			out.WriteString(prefix + formula[tokenStart:i])
			continue
		}
		lo, hi := first.column, last.column
		if m.rows() {
			lo, hi = first.row, last.row
		}
		if lo > hi {
			return refuse()
		}
		lo, hi, keep, err := m.shiftInterval(lo, hi)
		if err != nil {
			return "", err
		}
		if !keep {
			out.WriteString("#REF!")
			continue
		}
		if m.rows() {
			first.row, last.row = lo, hi
		} else {
			first.column, last.column = lo, hi
		}
		out.WriteString(prefix + first.text())
		if rangeRef {
			out.WriteByte(':')
			out.WriteString(last.text())
		}
	}
	return out.String(), nil
}
