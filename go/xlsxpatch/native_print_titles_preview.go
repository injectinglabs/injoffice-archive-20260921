package xlsxpatch

import (
	"regexp"
	"strconv"
)

// NativeSheetPrintTitlesV1 describes read-only worksheet-local repeated headings.
type NativeSheetPrintTitlesV1 struct {
	SheetID   string                  `json:"sheet_id"`
	SheetPart string                  `json:"sheet_part"`
	Status    string                  `json:"status"`
	Warnings  []string                `json:"warnings"`
	Rows      *NativePrintTitleSpanV1 `json:"rows,omitempty"`
	Columns   *NativePrintTitleSpanV1 `json:"columns,omitempty"`
}

// NativePrintTitleSpanV1 coordinates are zero-based and inclusive.
type NativePrintTitleSpanV1 struct {
	Start int `json:"start"`
	End   int `json:"end"`
}

var nativePrintTitleRows = regexp.MustCompile(`^\$([1-9][0-9]{0,6}):\$([1-9][0-9]{0,6})$`)
var nativePrintTitleColumns = regexp.MustCompile(`^\$([A-Z]{1,3}):\$([A-Z]{1,3})$`)

func previewNativePrintTitles(raw []byte, sheets []NativeWorkbookSheetV2) []NativeSheetPrintTitlesV1 {
	count := len(sheets)
	if count > 64 {
		count = 64
	}
	result := make([]NativeSheetPrintTitlesV1, count)
	_, titles, valid := collectNativePrintNames(raw, sheets)
	for i, sheet := range sheets[:count] {
		result[i] = NativeSheetPrintTitlesV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable", Warnings: []string{"Saved print titles unavailable: requires one worksheet-local same-sheet absolute whole-row and/or whole-column range."}}
		defs := titles[sheet.Order]
		if !valid || len(defs) != 1 || !validNativePrintName(defs[0]) {
			continue
		}
		rows, columns := parseNativePrintTitles(defs[0].text, sheet.Name)
		if rows == nil && columns == nil {
			continue
		}
		result[i].Status, result[i].Rows, result[i].Columns = "available", rows, columns
		result[i].Warnings = []string{"Read-only saved print headings. Repetition requires explicit preview selection; layout is approximate and grants no mutation authority."}
	}
	return result
}

func parseNativePrintTitles(text, sheetName string) (rows, columns *NativePrintTitleSpanV1) {
	if len(text) == 0 || len(text) > 4096 {
		return nil, nil
	}
	// Commas inside a quoted sheet name are literals, not union separators.
	parts := []string{}
	quoted, start := false, 0
	for i := 0; i < len(text); i++ {
		if text[i] == '\'' {
			if quoted && i+1 < len(text) && text[i+1] == '\'' {
				i++
				continue
			}
			quoted = !quoted
		} else if text[i] == ',' && !quoted {
			parts = append(parts, text[start:i])
			start = i + 1
		}
	}
	if quoted {
		return nil, nil
	}
	parts = append(parts, text[start:])
	if len(parts) > 2 {
		return nil, nil
	}
	for _, part := range parts {
		ref, ok := nativePrintReference(part, sheetName)
		if !ok {
			return nil, nil
		}
		if m := nativePrintTitleRows.FindStringSubmatch(ref); m != nil {
			if rows != nil {
				return nil, nil
			}
			a, _ := strconv.Atoi(m[1])
			b, _ := strconv.Atoi(m[2])
			if b < a || b > 1048576 {
				return nil, nil
			}
			rows = &NativePrintTitleSpanV1{Start: a - 1, End: b - 1}
		} else if m := nativePrintTitleColumns.FindStringSubmatch(ref); m != nil {
			if columns != nil {
				return nil, nil
			}
			column := func(s string) int {
				v := 0
				for _, c := range s {
					v = v*26 + int(c-'A'+1)
				}
				return v - 1
			}
			a, b := column(m[1]), column(m[2])
			if b < a || b >= 16384 {
				return nil, nil
			}
			columns = &NativePrintTitleSpanV1{Start: a, End: b}
		} else {
			return nil, nil
		}
	}
	return rows, columns
}
