package xlsxpatch

import (
	"strconv"
	"strings"
)

type nativePrintCountaBudget struct{ dependencies, coordinates int }

// COUNTA has one range argument and no expression operators. The inherited
// OFFSET parser retains all reference arithmetic and integer/grid checks.
func parseNativePrintCountaOffset(text string, sheet *NativeWorkbookSheetV2, source *nativePrintCountaSourceContext, budget *nativePrintCountaBudget) (*NativePrintAreaRectV1, []string) {
	if sheet == nil || source == nil || budget == nil || !source.certified[sheet] || len(text) > 2048 || !strings.HasPrefix(text, "OFFSET(") || !strings.HasSuffix(text, ")") {
		return nil, nil
	}
	for _, c := range text {
		if c < 32 || c == 127 {
			return nil, nil
		}
	}
	args := splitNativePrintUnion(text[7:len(text)-1], 5)
	if len(args) < 3 || len(args) > 5 {
		return nil, nil
	}
	next := *budget
	var warnings []string
	for i := 1; i < len(args); i++ {
		arg := strings.Trim(args[i], " ")
		if !strings.HasPrefix(arg, "COUNTA(") {
			continue
		}
		if !strings.HasSuffix(arg, ")") {
			return nil, nil
		}
		inner := strings.Trim(arg[7:len(arg)-1], " ")
		rect := parseNativePrintAreaRect(inner, sheet.Name)
		if rect == nil {
			return nil, nil
		}
		// Multiplication is safe in int64 even for the complete Excel grid.
		coordinates := int64(rect.EndRow-rect.Row+1) * int64(rect.EndColumn-rect.Column+1)
		if coordinates > int64(nativePrintCountaLimit-next.coordinates) || next.dependencies >= 4 {
			return nil, nil
		}
		n, ok := source.count(sheet, *rect)
		if !ok {
			return nil, nil
		}
		next.dependencies++
		next.coordinates += int(coordinates)
		args[i] = strconv.Itoa(n)
		warnings = append(warnings, "OFFSET saved COUNTA argument: "+arg+" = "+strconv.Itoa(n)+". Counts qualified saved literals only; reinspect the source package after changing this range.")
	}
	if len(warnings) == 0 {
		return nil, nil
	}
	resolved := "OFFSET(" + strings.Join(args, ",") + ")"
	area := parseNativePrintOffset(resolved, sheet.Name)
	if area == nil {
		var scalarWarnings []string
		area, scalarWarnings = parseNativePrintSourceCellOffset(resolved, *sheet)
		if area == nil || next.dependencies+len(scalarWarnings) > 4 {
			return nil, nil
		}
		next.dependencies += len(scalarWarnings)
		warnings = append(warnings, scalarWarnings...)
	}
	if area == nil {
		return nil, nil
	}
	*budget = next
	return area, warnings
}

// This fallback owns only sets containing at least one qualified COUNTA. Old
// literal and OFFSET routes retain their precedence when this is integrated.
func parseNativePrintCountaAreas(text string, sheet *NativeWorkbookSheetV2, source *nativePrintCountaSourceContext) ([]NativePrintAreaRectV1, []string) {
	if sheet == nil || source == nil || !source.certified[sheet] {
		return nil, nil
	}
	parts := splitNativePrintCountaAreas(text)
	if parts == nil {
		return nil, nil
	}
	areas := make([]NativePrintAreaRectV1, 0, len(parts))
	var warnings []string
	budget := nativePrintCountaBudget{}
	hasCounta := false
	for _, part := range parts {
		part = strings.Trim(part, " ")
		area := parseNativePrintAreaRect(part, sheet.Name)
		if area == nil {
			area = parseNativePrintOffset(part, sheet.Name)
		}
		if area == nil {
			var inputs []string
			area, inputs = parseNativePrintSourceCellOffset(part, *sheet)
			if area != nil {
				budget.dependencies += len(inputs)
				if budget.dependencies > 4 {
					return nil, nil
				}
			} else {
				area, inputs = parseNativePrintCountaOffset(part, sheet, source, &budget)
				if area != nil {
					hasCounta = true
				}
			}
			warnings = append(warnings, inputs...)
		}
		if area == nil {
			return nil, nil
		}
		for _, previous := range areas {
			if area.Row <= previous.EndRow && previous.Row <= area.EndRow && area.Column <= previous.EndColumn && previous.Column <= area.EndColumn {
				return nil, nil
			}
		}
		areas = append(areas, *area)
	}
	if !hasCounta {
		return nil, nil
	}
	return areas, warnings
}

// Parenthesis depth only bounds tokenization. Every resulting component must
// still pass a closed literal/OFFSET/COUNTA parser; no nested general syntax is
// admitted merely because its parentheses fit. Quoted sheet punctuation is data.
func splitNativePrintCountaAreas(text string) []string {
	if len(text) == 0 || len(text) > 2048 {
		return nil
	}
	quoted, depth, start := false, 0, 0
	parts := []string{}
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
			if depth > 2 {
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
