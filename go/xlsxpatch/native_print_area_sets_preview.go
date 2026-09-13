package xlsxpatch

// NativeSheetPrintAreaSetV1 is an ordered, read-only saved print-area projection.
// It grants no mutation authority. Rectangles are inclusive and zero-based.
type NativeSheetPrintAreaSetV1 struct {
	SheetID   string                  `json:"sheet_id"`
	SheetPart string                  `json:"sheet_part"`
	Status    string                  `json:"status"`
	Warnings  []string                `json:"warnings"`
	Areas     []NativePrintAreaRectV1 `json:"areas,omitempty"`
}

func previewNativePrintAreaSets(raw []byte, sheets []NativeWorkbookSheetV2) []NativeSheetPrintAreaSetV1 {
	count := len(sheets)
	if count > 64 {
		count = 64
	}
	result := make([]NativeSheetPrintAreaSetV1, count)
	names, titles, valid := collectNativePrintNames(raw, sheets)
	for i, sheet := range sheets[:count] {
		result[i] = NativeSheetPrintAreaSetV1{SheetID: sheet.ID, SheetPart: sheet.PartName, Status: "unavailable", Warnings: []string{"Saved print areas unavailable: requires one worksheet-local name containing 1 to 16 non-overlapping absolute same-sheet rectangles and supported saved print titles, if present."}}
		defs := names[sheet.Order]
		if !valid || len(defs) != 1 || !validNativePrintName(defs[0]) {
			continue
		}
		if ts := titles[sheet.Order]; len(ts) > 0 {
			if len(ts) != 1 || !validNativePrintName(ts[0]) {
				continue
			}
			rows, columns := parseNativePrintTitles(ts[0].text, sheet.Name)
			if rows == nil && columns == nil {
				continue
			}
		}
		areas := parseNativePrintAreaSet(defs[0].text, sheet.Name)
		if areas == nil {
			continue
		}
		result[i].Status, result[i].Areas = "available", areas
		result[i].Warnings = []string{"Read-only saved print-area rectangles in source order. Each area starts a separate approximate preview sequence; no formula evaluation, printer fidelity or mutation authority. Saved print titles require explicit preview selection."}
	}
	return result
}

func parseNativePrintAreaSet(text, sheetName string) []NativePrintAreaRectV1 {
	parts := splitNativePrintUnion(text, 16)
	if parts == nil {
		return nil
	}
	areas := make([]NativePrintAreaRectV1, 0, len(parts))
	for _, part := range parts {
		area := parseNativePrintAreaRect(part, sheetName)
		if area == nil {
			return nil
		}
		for _, previous := range areas {
			if area.Row <= previous.EndRow && previous.Row <= area.EndRow && area.Column <= previous.EndColumn && previous.Column <= area.EndColumn {
				return nil
			}
		}
		areas = append(areas, *area)
	}
	return areas
}

// Commas inside quoted sheet tokens are literals. Parsing each token remains
// responsible for the closed absolute-reference grammar; no formula is evaluated.
func splitNativePrintUnion(text string, limit int) []string {
	if len(text) == 0 || len(text) > 4096 {
		return nil
	}
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
			if i == start || len(parts) >= limit-1 {
				return nil
			}
			parts = append(parts, text[start:i])
			start = i + 1
		}
	}
	if quoted || start == len(text) {
		return nil
	}
	return append(parts, text[start:])
}
