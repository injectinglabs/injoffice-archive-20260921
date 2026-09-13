package xlsxpatch

import "strings"

// Resolve at most four scalar references from the same extracted worksheet.
// These are saved numeric source values, never formula caches or host state.
func parseNativePrintSourceCellOffset(text string, sheet NativeWorkbookSheetV2) (*NativePrintAreaRectV1, []string) {
	if !sheet.Editable || sheet.RefusalCode != nil {
		return nil, nil
	}
	var warnings []string
	area := parseNativePrintOffsetArguments(text, sheet.Name, func(arg string) (int, bool) {
		ref, ok := nativePrintReference(arg, sheet.Name)
		if !ok || strings.Contains(ref, ":") {
			return 0, false
		}
		rect := parseNativePrintRectReference(ref)
		if rect == nil {
			return 0, false
		}
		// A merged owner is also excluded: no merge reference semantics inferred.
		for _, merged := range sheet.MergedRanges {
			if rect.Row >= merged.Row && rect.Row <= merged.EndRow && rect.Column >= merged.Column && rect.Column <= merged.EndColumn {
				return 0, false
			}
		}
		canonical := cellReference(rect.Row, rect.Column)
		var found *NativeWorkbookCellV2
		for i := range sheet.Cells {
			cell := &sheet.Cells[i]
			if cell.Ref == canonical || (cell.Row == rect.Row && cell.Column == rect.Column) {
				if found != nil || cell.Ref != canonical || cell.Row != rect.Row || cell.Column != rect.Column {
					return 0, false
				}
				found = cell
			}
		}
		// Editable is used only as a conservative known-source-semantics guard.
		// This preview does not grant or change any mutation capability.
		if found == nil || !found.Editable || found.Formula != nil || (found.OOXMLType != nil && *found.OOXMLType != "n") {
			return 0, false
		}
		value := found.Value
		if value == nil || value.Kind != "number" || value.Storage != "number" || value.Lexical == nil || value.Text != nil || value.Rich || len(value.Runs) != 0 {
			return 0, false
		}
		n, ok := nativePrintOffsetIntegerValue(*value.Lexical)
		if !ok {
			return 0, false
		}
		warnings = append(warnings, "OFFSET saved numeric argument: "+arg+" = "+*value.Lexical+". Reinspect the source package after changing this cell.")
		return n, true
	})
	if area == nil || len(warnings) == 0 {
		return nil, nil
	}
	return area, warnings
}
