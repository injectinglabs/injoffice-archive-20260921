package xlsxpatch

import "strings"

// Additive fallback for unions containing at least one qualified OFFSET.
// Literal-only sets retain their existing parser and larger text budget.
func parseNativePrintFormulaUnion(text string, sheet NativeWorkbookSheetV2) ([]NativePrintAreaRectV1, []string) {
	parts := splitNativePrintFormulaUnion(text)
	if len(parts) < 2 {
		return nil, nil
	}
	areas := make([]NativePrintAreaRectV1, 0, len(parts))
	var warnings []string
	hasOffset := false
	for _, part := range parts {
		part = strings.Trim(part, " ")
		area := parseNativePrintAreaRect(part, sheet.Name)
		if area == nil {
			area = parseNativePrintOffset(part, sheet.Name)
			if area == nil {
				var inputs []string
				area, inputs = parseNativePrintSourceCellOffset(part, sheet)
				// Count occurrences, not distinct cells, across the complete union.
				if len(warnings)+len(inputs) > 4 {
					return nil, nil
				}
				warnings = append(warnings, inputs...)
			}
			if area == nil {
				return nil, nil
			}
			hasOffset = true
		}
		for _, previous := range areas {
			if area.Row <= previous.EndRow && previous.Row <= area.EndRow && area.Column <= previous.EndColumn && previous.Column <= area.EndColumn {
				return nil, nil
			}
		}
		areas = append(areas, *area)
	}
	if !hasOffset {
		return nil, nil
	}
	return areas, warnings
}

// Only top-level commas delimit components. A quoted sheet token may contain
// doubled apostrophes, commas and parentheses; unquoted nesting is unsupported.
func splitNativePrintFormulaUnion(text string) []string {
	if len(text) == 0 || len(text) > 2048 {
		return nil
	}
	var parts []string
	quoted, depth, start := false, 0, 0
	for i := 0; i < len(text); i++ {
		c := text[i]
		if c < 32 || c == 127 {
			return nil
		}
		if c == '\'' {
			if quoted && i+1 < len(text) && text[i+1] == '\'' {
				i++
				continue
			}
			quoted = !quoted
			continue
		}
		if quoted {
			continue
		}
		switch c {
		case '(':
			depth++
			if depth > 1 {
				return nil
			}
		case ')':
			depth--
			if depth < 0 {
				return nil
			}
		case ',':
			if depth == 0 {
				if strings.Trim(text[start:i], " ") == "" || len(parts) >= 15 {
					return nil
				}
				parts = append(parts, text[start:i])
				start = i + 1
			}
		}
	}
	if quoted || depth != 0 || strings.Trim(text[start:], " ") == "" {
		return nil
	}
	return append(parts, text[start:])
}
