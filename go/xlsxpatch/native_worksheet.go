package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
	"strings"
)

func (extractor *nativeWorkbookExtractor) extractWorksheet(route nativeWorkbookSheetRoute, order int, data []byte, shared []nativeSharedString, styleCount int, hasStyles bool) (NativeWorkbookSheetV1, error) {
	sheet := NativeWorkbookSheetV1{
		ID: route.id, Name: route.name, Order: order, State: route.state, PartName: route.part,
		Rows: []NativeWorkbookRowDimensionV1{}, Columns: []NativeWorkbookColumnDimensionV1{},
		Cells: []NativeWorkbookCellV1{}, MergedRanges: []NativeWorkbookMergedRangeV1{}, Editable: true,
	}
	if route.refusalCode != "" {
		sheet.Editable = false
		sheet.RefusalCode = nativeWorkbookString(route.refusalCode)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	root, err := nativeReadXMLRoot(decoder)
	if err != nil {
		return sheet, err
	}
	if root.Name != (xml.Name{Space: extractor.namespace, Local: "worksheet"}) {
		return sheet, fmt.Errorf("worksheet root uses an unexpected namespace")
	}
	if !nativeWorksheetRootAttributesAreBenign(root) {
		sheet.Editable = false
		sheet.RefusalCode = nativeWorkbookString("UNSAFE_WORKSHEET_ATTRIBUTES")
		if err := extractor.addUnsupported("WORKSHEET_ATTRIBUTES", "worksheet-features", "sheet:"+route.id, route.part, "", "unmodeled worksheet attributes remain authority-bound to the source package and mutation is refused"); err != nil {
			return sheet, err
		}
	}
	sheetFormatSeen, colsSeen, sheetDataSeen, mergeCellsSeen := false, false, false, false
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return sheet, fmt.Errorf("worksheet has no closing root")
		}
		if err != nil {
			return sheet, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name.Space == extractor.namespace && token.Name.Local == "sheetViews" {
				benign, parseErr := parseNativeBenignSheetViews(decoder, token)
				if parseErr != nil {
					return sheet, parseErr
				}
				if !benign {
					if err := extractor.addUnsupported("SHEET_VIEW_GEOMETRY", "dimensions", "sheet:"+route.id, route.part, "", "unmodeled sheet-view geometry remains authority-bound to the source package"); err != nil {
						return sheet, err
					}
				}
				continue
			}
			if token.Name == (xml.Name{Space: extractor.namespace, Local: "sheetFormatPr"}) {
				if sheetFormatSeen || colsSeen || sheetDataSeen {
					return sheet, fmt.Errorf("sheetFormatPr is duplicated or outside canonical schema order")
				}
				sheetFormatSeen = true
				format, parseErr := extractor.parseNativeSheetFormat(decoder, token, route)
				if parseErr != nil {
					return sheet, parseErr
				}
				sheet.SheetFormat = &format
				continue
			}
			if token.Name.Local == "sheetFormatPr" && (token.Name.Space == spreadsheetMLTransitional || token.Name.Space == spreadsheetMLStrict) {
				return sheet, fmt.Errorf("sheetFormatPr uses the opposing Strict/Transitional SpreadsheetML namespace")
			}
			if token.Name == (xml.Name{Space: extractor.namespace, Local: "mergeCells"}) {
				if mergeCellsSeen || !sheetDataSeen {
					return sheet, fmt.Errorf("mergeCells is duplicated or outside canonical schema order")
				}
				mergeCellsSeen = true
				mergedRanges, parseErr := extractor.parseNativeMergedRanges(decoder, token)
				if parseErr != nil {
					return sheet, parseErr
				}
				sheet.MergedRanges = mergedRanges
				if err := validateAndMarkCellsInNativeMergedRanges(sheet.Cells, mergedRanges); err != nil {
					return sheet, err
				}
				if err := extractor.addUnsupported("MERGED_CELLS", "merges", "sheet:"+route.id, route.part, "", "merged ranges are projected exactly as read-only grid regions; original SpreadsheetML remains source authority until an atomic merge writer is proven"); err != nil {
					return sheet, err
				}
				continue
			}
			if token.Name.Local == "mergeCell" && token.Name.Space == extractor.namespace {
				return sheet, fmt.Errorf("mergeCell appears outside mergeCells")
			}
			if (token.Name.Local == "mergeCells" || token.Name.Local == "mergeCell") &&
				(token.Name.Space == spreadsheetMLTransitional || token.Name.Space == spreadsheetMLStrict) {
				return sheet, fmt.Errorf("merged-cell markup uses the opposing Strict/Transitional SpreadsheetML namespace")
			}
			if token.Name.Space == extractor.namespace && token.Name.Local == "cols" {
				if colsSeen || sheetDataSeen {
					return sheet, fmt.Errorf("cols is duplicated or outside canonical schema order")
				}
				colsSeen = true
				if len(unexpectedSemanticXMLAttributes(token)) != 0 {
					if err := extractor.addUnsupported("COLS_ATTRIBUTES", "dimensions", "sheet:"+route.id, route.part, "", "unmodeled column-container attributes are preserved exactly"); err != nil {
						return sheet, err
					}
				}
				columns, parseErr := extractor.parseNativeColumns(decoder, token, route, styleCount, hasStyles)
				if parseErr != nil {
					return sheet, parseErr
				}
				sheet.Columns = columns
				continue
			}
			if token.Name.Space == extractor.namespace && token.Name.Local == "sheetData" {
				if sheetDataSeen {
					return sheet, fmt.Errorf("worksheet has multiple sheetData elements")
				}
				sheetDataSeen = true
				if len(unexpectedSemanticXMLAttributes(token)) != 0 {
					sheet.Editable = false
					sheet.RefusalCode = nativeWorkbookString("SHEET_DATA_ATTRIBUTES")
					if err := extractor.addUnsupported("SHEET_DATA_ATTRIBUTES", "worksheet-features", "sheet:"+route.id, route.part, "", "unmodeled sheetData attributes remain authority-bound to the source package and mutation is refused"); err != nil {
						return sheet, err
					}
				}
				rows, cells, parseErr := extractor.parseNativeSheetData(decoder, token, route, shared, styleCount, hasStyles)
				if parseErr != nil {
					return sheet, parseErr
				}
				sheet.Rows, sheet.Cells = rows, cells
				continue
			}
			code, capability := nativeWorksheetFeature(token.Name)
			if err := extractor.addUnsupported(code, capability, "sheet:"+route.id, route.part, "", "worksheet feature is preserved exactly outside the v1 native projection"); err != nil {
				return sheet, err
			}
			if token.Name.Space == extractor.namespace && (token.Name.Local == "sheetProtection" || token.Name.Local == "protectedRanges") {
				sheet.Editable = false
				refusal := "SHEET_PROTECTION"
				sheet.RefusalCode = &refusal
			}
			if err := skipNativeXMLElement(decoder, token, 1); err != nil {
				return sheet, err
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return sheet, fmt.Errorf("worksheet has mismatched closing element")
			}
			if !sheetDataSeen {
				return sheet, fmt.Errorf("worksheet has no sheetData")
			}
			if err := extractor.markNativeFormulaGroupFollowers(&sheet); err != nil {
				return sheet, err
			}
			if err := nativeRequireXMLEOF(decoder); err != nil {
				return sheet, err
			}
			return sheet, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return sheet, fmt.Errorf("worksheet contains unsupported direct text")
			}
		case xml.ProcInst:
			return sheet, fmt.Errorf("worksheet contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return sheet, fmt.Errorf("worksheet contains unsupported XML directive")
		}
	}
}

func nativeWorksheetRootAttributesAreBenign(root xml.StartElement) bool {
	const (
		compatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006"
		revisionNamespace      = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
	)
	namespaces := map[string]string{}
	for _, attribute := range root.Attr {
		if attribute.Name.Space == "xmlns" {
			namespaces[attribute.Name.Local] = attribute.Value
		}
	}
	for _, attribute := range root.Attr {
		if isNamespaceDeclaration(attribute) {
			continue
		}
		switch attribute.Name {
		case xml.Name{Space: compatibilityNamespace, Local: "Ignorable"}:
			if attribute.Value != "x14ac xr xr2 xr3" ||
				namespaces["x14ac"] != "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" ||
				namespaces["xr"] != revisionNamespace ||
				namespaces["xr2"] != "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2" ||
				namespaces["xr3"] != "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3" {
				return false
			}
		case xml.Name{Space: revisionNamespace, Local: "uid"}:
			if !validNativeWorksheetUID(attribute.Value) {
				return false
			}
		default:
			return false
		}
	}
	return true
}

func validNativeWorksheetUID(value string) bool {
	if len(value) != 38 || value[0] != '{' || value[37] != '}' {
		return false
	}
	for index := 1; index < 37; index++ {
		if index == 9 || index == 14 || index == 19 || index == 24 {
			if value[index] != '-' {
				return false
			}
			continue
		}
		character := value[index]
		if !((character >= '0' && character <= '9') || (character >= 'a' && character <= 'f') || (character >= 'A' && character <= 'F')) {
			return false
		}
	}
	return true
}

// Workbook selection state is UI metadata, not grid geometry. Only this narrow,
// fully parsed shape is ignored; any view option or unknown markup remains
// source-authoritative as SHEET_VIEW_GEOMETRY.
func parseNativeBenignSheetViews(decoder *xml.Decoder, root xml.StartElement) (bool, error) {
	benign := len(unexpectedSemanticXMLAttributes(root)) == 0
	depth, sheetViews, selections := 0, 0, 0
	for {
		token, err := decoder.Token()
		if err != nil {
			return false, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if depth == 0 && token.Name == (xml.Name{Space: root.Name.Space, Local: "sheetView"}) {
				sheetViews++
				if sheetViews != 1 || !nativeBenignSheetViewAttributes(token) {
					benign = false
				}
			} else if depth == 1 && token.Name == (xml.Name{Space: root.Name.Space, Local: "selection"}) {
				selections++
				if selections != 1 || !nativeBenignSelectionAttributes(token) {
					benign = false
				}
			} else {
				benign = false
			}
			depth++
		case xml.EndElement:
			if depth == 0 {
				if token.Name != root.Name {
					return false, fmt.Errorf("sheetViews has a mismatched closing element")
				}
				return benign && sheetViews == 1, nil
			}
			depth--
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				benign = false
			}
		case xml.ProcInst:
			return false, fmt.Errorf("sheetViews contains a processing instruction")
		case xml.Directive:
			return false, fmt.Errorf("sheetViews contains an XML directive")
		}
	}
}

func nativeBenignSheetViewAttributes(element xml.StartElement) bool {
	if len(unexpectedSemanticXMLAttributes(element, xml.Name{Local: "workbookViewId"}, xml.Name{Local: "tabSelected"}, xml.Name{Local: "showGridLines"})) != 0 {
		return false
	}
	viewID, found, err := unqualifiedXMLAttribute(element, "workbookViewId")
	if err != nil || !found {
		return false
	}
	if _, err = strconv.ParseUint(viewID, 10, 32); err != nil {
		return false
	}
	selected, found, err := unqualifiedXMLAttribute(element, "tabSelected")
	if err != nil || (found && selected != "0" && selected != "1" && selected != "false" && selected != "true") {
		return false
	}
	// Excel default is gridlines on. Native chrome already draws the grid;
	// only an explicit off is view geometry that we do not project.
	grid, found, err := unqualifiedXMLAttribute(element, "showGridLines")
	return err == nil && (!found || grid == "1" || grid == "true")
}

func nativeDefaultOffXMLFlag(element xml.StartElement, local string) bool {
	value, found, err := unqualifiedXMLAttribute(element, local)
	return err == nil && (!found || value == "0" || value == "false")
}

func nativeBenignSelectionAttributes(element xml.StartElement) bool {
	if len(unexpectedSemanticXMLAttributes(element, xml.Name{Local: "activeCell"}, xml.Name{Local: "sqref"})) != 0 {
		return false
	}
	active, activeFound, activeErr := unqualifiedXMLAttribute(element, "activeCell")
	selection, selectionFound, selectionErr := unqualifiedXMLAttribute(element, "sqref")
	if activeErr != nil || selectionErr != nil || !activeFound || !selectionFound || active != selection {
		return false
	}
	row, column, err := parseCellReference(active)
	return err == nil && cellReference(row, column) == active
}

func (extractor *nativeWorkbookExtractor) parseNativeSheetFormat(decoder *xml.Decoder, root xml.StartElement, route nativeWorkbookSheetRoute) (NativeWorkbookSheetFormatV1, error) {
	format := NativeWorkbookSheetFormatV1{}
	if len(unexpectedSemanticXMLAttributes(root,
		xml.Name{Local: "baseColWidth"}, xml.Name{Local: "defaultColWidth"}, xml.Name{Local: "defaultRowHeight"},
		xml.Name{Local: "customHeight"}, xml.Name{Local: "zeroHeight"},
	)) != 0 {
		if err := extractor.addUnsupported("SHEET_FORMAT_EXTRAS", "dimensions", "sheet:"+route.id, route.part, "", "outline levels, thick-edge flags, or unmodeled sheet-format attributes remain authority-bound to the source package"); err != nil {
			return format, err
		}
	}
	rawHeight, found, err := unqualifiedXMLAttribute(root, "defaultRowHeight")
	if err != nil {
		return format, err
	}
	if !found || rawHeight == "" {
		return format, fmt.Errorf("sheetFormatPr requires defaultRowHeight")
	}
	format.DefaultRowHeightPoints, err = finiteNativeFloat(rawHeight, 0, maxRowHeightPoints)
	if err != nil || format.DefaultRowHeightPoints <= 0 {
		return format, fmt.Errorf("sheetFormatPr defaultRowHeight must be finite and within 0..%g points", maxRowHeightPoints)
	}
	if base, present, parseErr := optionalNativeUintAttribute(root, "baseColWidth", uint64(maxColumnWidth)); parseErr != nil {
		return format, fmt.Errorf("sheetFormatPr baseColWidth: %w", parseErr)
	} else if present {
		value := uint32(base)
		format.BaseColumnWidth = &value
	}
	if raw, present, attrErr := unqualifiedXMLAttribute(root, "defaultColWidth"); attrErr != nil {
		return format, attrErr
	} else if present {
		value, parseErr := finiteNativeFloat(raw, 0, maxColumnWidth)
		if parseErr != nil || value <= 0 {
			return format, fmt.Errorf("sheetFormatPr defaultColWidth must be finite and within 0..%g", maxColumnWidth)
		}
		format.DefaultColumnWidth = nativeWorkbookFloat(value)
	}
	if format.CustomHeight, err = nativeBooleanAttribute(root, "customHeight", false); err != nil {
		return format, fmt.Errorf("sheetFormatPr customHeight: %w", err)
	}
	if format.ZeroHeight, err = nativeBooleanAttribute(root, "zeroHeight", false); err != nil {
		return format, fmt.Errorf("sheetFormatPr zeroHeight: %w", err)
	}
	if err := requireNativeEmptyElement(decoder, root); err != nil {
		return format, fmt.Errorf("sheetFormatPr: %w", err)
	}
	return format, nil
}

func (extractor *nativeWorkbookExtractor) parseNativeMergedRanges(decoder *xml.Decoder, root xml.StartElement) ([]NativeWorkbookMergedRangeV1, error) {
	if err := requireOnlySemanticXMLAttributes(root, xml.Name{Local: "count"}); err != nil {
		return nil, fmt.Errorf("mergeCells: %w", err)
	}
	declaredCount, countFound, err := optionalNativeUintAttribute(root, "count", NativeXLSXMaxMergedRanges)
	if err != nil {
		return nil, fmt.Errorf("mergeCells count: %w", err)
	}
	ranges := make([]NativeWorkbookMergedRangeV1, 0)
	seen := make(map[string]bool)
	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name != (xml.Name{Space: extractor.namespace, Local: "mergeCell"}) {
				return nil, fmt.Errorf("mergeCells has unsupported direct child {%s}%s", token.Name.Space, token.Name.Local)
			}
			if err := requireOnlySemanticXMLAttributes(token, xml.Name{Local: "ref"}); err != nil {
				return nil, fmt.Errorf("mergeCell: %w", err)
			}
			rawRef, found, err := unqualifiedXMLAttribute(token, "ref")
			if err != nil {
				return nil, fmt.Errorf("mergeCell ref: %w", err)
			}
			if !found || rawRef == "" {
				return nil, fmt.Errorf("mergeCell requires a non-empty ref")
			}
			canonicalRef, err := canonicalNativeMergedRangeReference(rawRef)
			if err != nil {
				return nil, fmt.Errorf("invalid merged range %q: %w", rawRef, err)
			}
			minimumRow, minimumColumn, maximumRow, maximumColumn, err := parseDimensionReference(canonicalRef)
			if err != nil {
				return nil, fmt.Errorf("invalid merged range %q: %w", rawRef, err)
			}
			if minimumRow == maximumRow && minimumColumn == maximumColumn {
				return nil, fmt.Errorf("merged range %q must span multiple cells", rawRef)
			}
			if seen[canonicalRef] {
				return nil, fmt.Errorf("duplicate merged range %q", canonicalRef)
			}
			seen[canonicalRef] = true
			if err := requireNativeEmptyElement(decoder, token); err != nil {
				return nil, fmt.Errorf("merged range %q: %w", rawRef, err)
			}
			ranges = append(ranges, NativeWorkbookMergedRangeV1{
				Ref: canonicalRef, Row: minimumRow, Column: minimumColumn,
				EndRow: maximumRow, EndColumn: maximumColumn, Editable: false,
			})
			if extractor.mergedRangeCount >= NativeXLSXMaxMergedRanges {
				return nil, fmt.Errorf("workbook mergeCells exceeds %d ranges", NativeXLSXMaxMergedRanges)
			}
			extractor.mergedRangeCount++
		case xml.EndElement:
			if token.Name != root.Name {
				return nil, fmt.Errorf("mergeCells has mismatched closing element")
			}
			if len(ranges) == 0 {
				return nil, fmt.Errorf("mergeCells must contain at least one mergeCell")
			}
			if countFound && declaredCount != uint64(len(ranges)) {
				return nil, fmt.Errorf("mergeCells count=%d does not match %d mergeCell elements", declaredCount, len(ranges))
			}
			if err := rejectOverlappingNativeMergedRanges(ranges); err != nil {
				return nil, err
			}
			sort.Slice(ranges, func(i, j int) bool {
				left, right := ranges[i], ranges[j]
				if left.Row != right.Row {
					return left.Row < right.Row
				}
				if left.Column != right.Column {
					return left.Column < right.Column
				}
				if left.EndRow != right.EndRow {
					return left.EndRow < right.EndRow
				}
				return left.EndColumn < right.EndColumn
			})
			return ranges, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return nil, fmt.Errorf("mergeCells contains unsupported direct text")
			}
		case xml.ProcInst:
			return nil, fmt.Errorf("mergeCells contains processing instruction")
		case xml.Directive:
			return nil, fmt.Errorf("mergeCells contains XML directive")
		}
	}
}

func canonicalNativeMergedRangeReference(ref string) (string, error) {
	parts := strings.Split(ref, ":")
	if len(parts) < 1 || len(parts) > 2 {
		return "", fmt.Errorf("invalid merged-cell reference grammar")
	}
	canonical := make([]string, len(parts))
	for index, part := range parts {
		if part == "" {
			return "", fmt.Errorf("invalid merged-cell reference grammar")
		}
		cursor := 0
		if part[cursor] == '$' {
			cursor++
		}
		columnStart := cursor
		for cursor < len(part) && ((part[cursor] >= 'A' && part[cursor] <= 'Z') || (part[cursor] >= 'a' && part[cursor] <= 'z')) {
			cursor++
		}
		if cursor == columnStart || cursor-columnStart > 3 {
			return "", fmt.Errorf("invalid merged-cell reference grammar")
		}
		columnEnd := cursor
		if cursor < len(part) && part[cursor] == '$' {
			cursor++
		}
		rowStart := cursor
		if rowStart == len(part) || part[rowStart] < '1' || part[rowStart] > '9' {
			return "", fmt.Errorf("invalid merged-cell reference grammar")
		}
		for cursor < len(part) && part[cursor] >= '0' && part[cursor] <= '9' {
			cursor++
		}
		if cursor != len(part) {
			return "", fmt.Errorf("invalid merged-cell reference grammar")
		}
		row, column, err := parseCellReference(part[columnStart:columnEnd] + part[rowStart:])
		if err != nil {
			return "", err
		}
		canonical[index] = cellReference(row, column)
	}
	result := canonical[0]
	if len(canonical) == 2 {
		firstRow, firstColumn, err := parseCellReference(canonical[0])
		if err != nil {
			return "", err
		}
		lastRow, lastColumn, err := parseCellReference(canonical[1])
		if err != nil {
			return "", err
		}
		if firstRow > lastRow || firstColumn > lastColumn {
			return "", fmt.Errorf("bounds are reversed")
		}
		result += ":" + canonical[1]
	}
	return result, nil
}

func rejectOverlappingNativeMergedRanges(ranges []NativeWorkbookMergedRangeV1) error {
	type event struct {
		row, minimumColumn, maximumColumn, delta int
		ref                                      string
	}
	events := make([]event, 0, len(ranges)*2)
	for _, merged := range ranges {
		events = append(events,
			event{row: merged.Row, minimumColumn: merged.Column, maximumColumn: merged.EndColumn, delta: 1, ref: merged.Ref},
			event{row: merged.EndRow + 1, minimumColumn: merged.Column, maximumColumn: merged.EndColumn, delta: -1, ref: merged.Ref},
		)
	}
	sort.Slice(events, func(i, j int) bool {
		if events[i].row != events[j].row {
			return events[i].row < events[j].row
		}
		if events[i].delta != events[j].delta {
			return events[i].delta < events[j].delta
		}
		if events[i].minimumColumn != events[j].minimumColumn {
			return events[i].minimumColumn < events[j].minimumColumn
		}
		if events[i].maximumColumn != events[j].maximumColumn {
			return events[i].maximumColumn < events[j].maximumColumn
		}
		return events[i].ref < events[j].ref
	})
	tree := newNativeColumnRangeTree(excelMaxColumns)
	for _, item := range events {
		if item.delta > 0 && tree.maximum(item.minimumColumn, item.maximumColumn) > 0 {
			return fmt.Errorf("merged range %q overlaps another merged range", item.ref)
		}
		tree.add(item.minimumColumn, item.maximumColumn, item.delta)
	}
	return nil
}

func validateAndMarkCellsInNativeMergedRanges(cells []NativeWorkbookCellV1, ranges []NativeWorkbookMergedRangeV1) error {
	type event struct {
		row, minimumColumn, maximumColumn, delta int
	}
	events := make([]event, 0, len(ranges)*2)
	anchors := make(map[cellKey]bool, len(ranges))
	for _, merged := range ranges {
		anchors[cellKey{row: merged.Row, column: merged.Column}] = true
		events = append(events,
			event{row: merged.Row, minimumColumn: merged.Column, maximumColumn: merged.EndColumn, delta: 1},
			event{row: merged.EndRow + 1, minimumColumn: merged.Column, maximumColumn: merged.EndColumn, delta: -1},
		)
	}
	sort.Slice(events, func(i, j int) bool {
		if events[i].row != events[j].row {
			return events[i].row < events[j].row
		}
		return events[i].delta < events[j].delta
	})
	tree, eventIndex := newNativeColumnRangeTree(excelMaxColumns), 0
	for index := range cells {
		for eventIndex < len(events) && events[eventIndex].row <= cells[index].Row {
			item := events[eventIndex]
			tree.add(item.minimumColumn, item.maximumColumn, item.delta)
			eventIndex++
		}
		if tree.maximum(cells[index].Column, cells[index].Column) > 0 {
			if !anchors[cellKey{row: cells[index].Row, column: cells[index].Column}] && (cells[index].Value != nil || cells[index].Formula != nil) {
				return fmt.Errorf("covered non-anchor cell %s contains a value or formula", cells[index].Ref)
			}
			cells[index].Editable = false
		}
	}
	return nil
}

func nativeWorksheetFeature(name xml.Name) (string, string) {
	if name.Space != spreadsheetMLTransitional && name.Space != spreadsheetMLStrict {
		return "FOREIGN_WORKSHEET_MARKUP", "extensions"
	}
	switch name.Local {
	case "sheetViews":
		return "SHEET_VIEW_GEOMETRY", "dimensions"
	case "mergeCells":
		return "MERGED_CELLS", "merges"
	case "conditionalFormatting":
		return "CONDITIONAL_FORMATTING", "conditional-formatting"
	case "dataValidations":
		return "DATA_VALIDATION", "data-validation"
	case "hyperlinks":
		return "HYPERLINKS", "hyperlinks"
	case "drawing", "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls":
		return "DRAWING_REFERENCE", "drawings"
	case "tableParts":
		return "TABLE_REFERENCE", "tables"
	case "sheetProtection", "protectedRanges":
		return "SHEET_PROTECTION", "protection"
	case "extLst":
		return "WORKSHEET_EXTENSIONS", "extensions"
	default:
		return "UNMODELED_WORKSHEET_FEATURE", "worksheet-features"
	}
}

func (extractor *nativeWorkbookExtractor) parseNativeColumns(decoder *xml.Decoder, root xml.StartElement, route nativeWorkbookSheetRoute, styleCount int, hasStyles bool) ([]NativeWorkbookColumnDimensionV1, error) {
	columns := []NativeWorkbookColumnDimensionV1{}
	lastMaximum := 0
	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name != (xml.Name{Space: extractor.namespace, Local: "col"}) {
				return nil, fmt.Errorf("cols has unsupported direct child {%s}%s", token.Name.Space, token.Name.Local)
			}
			minimum, found, err := optionalNativeUintAttribute(token, "min", excelMaxColumns)
			if err != nil || !found || minimum == 0 {
				return nil, fmt.Errorf("column range requires min within Excel bounds")
			}
			maximum, found, err := optionalNativeUintAttribute(token, "max", excelMaxColumns)
			if err != nil || !found || maximum < minimum {
				return nil, fmt.Errorf("column range requires max >= min within Excel bounds")
			}
			if int(minimum) <= lastMaximum {
				return nil, fmt.Errorf("column ranges overlap or are not strictly ordered at %d:%d", minimum, maximum)
			}
			lastMaximum = int(maximum)
			column := NativeWorkbookColumnDimensionV1{Column: int(minimum) - 1, EndColumn: int(maximum) - 1}
			if len(unexpectedSemanticXMLAttributes(token,
				xml.Name{Local: "min"}, xml.Name{Local: "max"}, xml.Name{Local: "width"}, xml.Name{Local: "hidden"},
				xml.Name{Local: "customWidth"}, xml.Name{Local: "bestFit"}, xml.Name{Local: "style"},
			)) != 0 {
				if err := extractor.addUnsupported("COLUMN_DIMENSION_EXTRAS", "dimensions", "sheet:"+route.id, route.part, "", "outline, collapsed, phonetic, or unmodeled column attributes are preserved but omitted from the bounded dimension projection"); err != nil {
					return nil, err
				}
			}
			if raw, found, attrErr := unqualifiedXMLAttribute(token, "width"); attrErr != nil {
				return nil, attrErr
			} else if found {
				value, parseErr := finiteNativeFloat(raw, 0, maxColumnWidth)
				if parseErr != nil {
					return nil, fmt.Errorf("column %d:%d width: %w", minimum, maximum, parseErr)
				}
				column.Width = nativeWorkbookFloat(value)
			}
			if column.Hidden, err = nativeBooleanAttribute(token, "hidden", false); err != nil {
				return nil, fmt.Errorf("column %d:%d hidden: %w", minimum, maximum, err)
			}
			if column.CustomWidth, err = nativeBooleanAttribute(token, "customWidth", false); err != nil {
				return nil, fmt.Errorf("column %d:%d customWidth: %w", minimum, maximum, err)
			}
			if column.BestFit, err = nativeBooleanAttribute(token, "bestFit", false); err != nil {
				return nil, fmt.Errorf("column %d:%d bestFit: %w", minimum, maximum, err)
			}
			if style, found, styleErr := optionalNativeStyleID(token, "style", styleCount, hasStyles); styleErr != nil {
				return nil, fmt.Errorf("column %d:%d style: %w", minimum, maximum, styleErr)
			} else if found {
				column.StyleID = &style
			}
			if err := requireNativeEmptyElement(decoder, token); err != nil {
				return nil, fmt.Errorf("column %d:%d: %w", minimum, maximum, err)
			}
			columns = append(columns, column)
			if len(columns) > excelMaxColumns {
				return nil, fmt.Errorf("column definitions exceed %d ranges", excelMaxColumns)
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return nil, fmt.Errorf("cols has mismatched closing element")
			}
			return columns, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return nil, fmt.Errorf("cols contains unsupported direct text")
			}
		case xml.ProcInst:
			return nil, fmt.Errorf("cols contains processing instruction")
		case xml.Directive:
			return nil, fmt.Errorf("cols contains XML directive")
		}
	}
}

func (extractor *nativeWorkbookExtractor) parseNativeSheetData(decoder *xml.Decoder, root xml.StartElement, route nativeWorkbookSheetRoute, shared []nativeSharedString, styleCount int, hasStyles bool) ([]NativeWorkbookRowDimensionV1, []NativeWorkbookCellV1, error) {
	rows := []NativeWorkbookRowDimensionV1{}
	cells := []NativeWorkbookCellV1{}
	lastRow := -1
	for {
		token, err := decoder.Token()
		if err != nil {
			return nil, nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name != (xml.Name{Space: extractor.namespace, Local: "row"}) {
				return nil, nil, fmt.Errorf("sheetData has unsupported direct child {%s}%s", token.Name.Space, token.Name.Local)
			}
			rowNumber := lastRow + 1
			if raw, found, attrErr := unqualifiedXMLAttribute(token, "r"); attrErr != nil {
				return nil, nil, attrErr
			} else if found {
				parsed, parseErr := parseNativePositiveIndex(raw, excelMaxRows)
				if parseErr != nil {
					return nil, nil, fmt.Errorf("row r: %w", parseErr)
				}
				rowNumber = parsed - 1
			}
			if rowNumber <= lastRow || rowNumber >= excelMaxRows {
				return nil, nil, fmt.Errorf("rows are duplicated, unordered, or outside Excel bounds at %d", rowNumber+1)
			}
			lastRow = rowNumber
			row, rowCells, parseErr := extractor.parseNativeRow(decoder, token, route, rowNumber, shared, styleCount, hasStyles)
			if parseErr != nil {
				return nil, nil, fmt.Errorf("row %d: %w", rowNumber+1, parseErr)
			}
			if row.HeightPoints != nil || row.Hidden || row.CustomHeight || row.StyleID != nil {
				rows = append(rows, row)
			}
			cells = append(cells, rowCells...)
			extractor.cellCount += len(rowCells)
			if extractor.cellCount > NativeXLSXMaxCells {
				return nil, nil, fmt.Errorf("workbook exceeds %d modeled cells", NativeXLSXMaxCells)
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return nil, nil, fmt.Errorf("sheetData has mismatched closing element")
			}
			return rows, cells, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return nil, nil, fmt.Errorf("sheetData contains unsupported direct text")
			}
		case xml.ProcInst:
			return nil, nil, fmt.Errorf("sheetData contains processing instruction")
		case xml.Directive:
			return nil, nil, fmt.Errorf("sheetData contains XML directive")
		}
	}
}

func (extractor *nativeWorkbookExtractor) parseNativeRow(decoder *xml.Decoder, root xml.StartElement, route nativeWorkbookSheetRoute, rowNumber int, shared []nativeSharedString, styleCount int, hasStyles bool) (NativeWorkbookRowDimensionV1, []NativeWorkbookCellV1, error) {
	row := NativeWorkbookRowDimensionV1{Row: rowNumber}
	if len(unexpectedSemanticXMLAttributes(root,
		xml.Name{Local: "r"}, xml.Name{Local: "ht"}, xml.Name{Local: "hidden"}, xml.Name{Local: "customHeight"}, xml.Name{Local: "s"}, xml.Name{Local: "spans"},
	)) != 0 {
		if err := extractor.addUnsupported("ROW_DIMENSION_EXTRAS", "dimensions", "sheet:"+route.id, route.part, "", "spans, outline, collapsed, custom-format, thick-border, phonetic, or unmodeled row attributes are preserved but omitted from the bounded dimension projection"); err != nil {
			return row, nil, err
		}
	}
	var err error
	if raw, found, attrErr := unqualifiedXMLAttribute(root, "ht"); attrErr != nil {
		return row, nil, attrErr
	} else if found {
		value, parseErr := finiteNativeFloat(raw, 0, maxRowHeightPoints)
		if parseErr != nil {
			return row, nil, fmt.Errorf("height: %w", parseErr)
		}
		row.HeightPoints = nativeWorkbookFloat(value)
	}
	if row.Hidden, err = nativeBooleanAttribute(root, "hidden", false); err != nil {
		return row, nil, fmt.Errorf("hidden: %w", err)
	}
	if row.CustomHeight, err = nativeBooleanAttribute(root, "customHeight", false); err != nil {
		return row, nil, fmt.Errorf("customHeight: %w", err)
	}
	if style, found, styleErr := optionalNativeStyleID(root, "s", styleCount, hasStyles); styleErr != nil {
		return row, nil, fmt.Errorf("style: %w", styleErr)
	} else if found {
		row.StyleID = &style
	}
	cells := []NativeWorkbookCellV1{}
	lastColumn := -1
	for {
		token, err := decoder.Token()
		if err != nil {
			return row, nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name == (xml.Name{Space: extractor.namespace, Local: "c"}) {
				column := lastColumn + 1
				if raw, found, attrErr := unqualifiedXMLAttribute(token, "r"); attrErr != nil {
					return row, nil, attrErr
				} else if found {
					cellRow, cellColumn, parseErr := parseCellReference(raw)
					if parseErr != nil || cellRow != rowNumber {
						return row, nil, fmt.Errorf("cell reference %q does not belong to row %d", raw, rowNumber+1)
					}
					column = cellColumn
				}
				if column <= lastColumn || column >= excelMaxColumns {
					return row, nil, fmt.Errorf("cells are duplicated, unordered, or outside Excel bounds at column %d", column+1)
				}
				lastColumn = column
				cell, parseErr := extractor.parseNativeCell(decoder, token, route, rowNumber, column, shared, styleCount, hasStyles)
				if parseErr != nil {
					return row, nil, fmt.Errorf("cell %s: %w", cellReference(rowNumber, column), parseErr)
				}
				cells = append(cells, cell)
				continue
			}
			if token.Name == (xml.Name{Space: extractor.namespace, Local: "extLst"}) {
				if err := extractor.addUnsupported("ROW_EXTENSIONS", "extensions", "sheet:"+route.id, route.part, "", "row extension markup is preserved exactly"); err != nil {
					return row, nil, err
				}
				if err := skipNativeXMLElement(decoder, token, 1); err != nil {
					return row, nil, err
				}
				continue
			}
			return row, nil, fmt.Errorf("row has unsupported direct child {%s}%s", token.Name.Space, token.Name.Local)
		case xml.EndElement:
			if token.Name != root.Name {
				return row, nil, fmt.Errorf("row has mismatched closing element")
			}
			return row, cells, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return row, nil, fmt.Errorf("row contains unsupported direct text")
			}
		case xml.ProcInst:
			return row, nil, fmt.Errorf("row contains processing instruction")
		case xml.Directive:
			return row, nil, fmt.Errorf("row contains XML directive")
		}
	}
}

type nativeFormulaParse struct {
	formula  NativeWorkbookFormulaV1
	editable bool
}

func (extractor *nativeWorkbookExtractor) parseNativeCell(decoder *xml.Decoder, root xml.StartElement, route nativeWorkbookSheetRoute, row, column int, shared []nativeSharedString, styleCount int, hasStyles bool) (NativeWorkbookCellV1, error) {
	ref := cellReference(row, column)
	cell := NativeWorkbookCellV1{Row: row, Column: column, Ref: ref, Editable: true}
	if len(unexpectedSemanticXMLAttributes(root,
		xml.Name{Local: "r"}, xml.Name{Local: "s"}, xml.Name{Local: "t"}, xml.Name{Local: "cm"}, xml.Name{Local: "vm"},
	)) != 0 {
		cell.Editable = false
		if err := extractor.addUnsupported("CELL_ATTRIBUTES", "cell-markup", "sheet:"+route.id, route.part, ref, "unmodeled cell attributes remain authority-bound to the source package and make the cell mutation-refused"); err != nil {
			return cell, err
		}
	}
	style, found, err := optionalNativeStyleID(root, "s", styleCount, hasStyles)
	if err != nil {
		return cell, err
	}
	if found {
		cell.StyleID = style
	}
	storage, found, err := unqualifiedXMLAttribute(root, "t")
	if err != nil {
		return cell, err
	}
	if found {
		cell.OOXMLType = nativeWorkbookString(storage)
	}
	if !found || storage == "n" {
		storage = "n"
	}
	switch storage {
	case "n", "s", "str", "inlineStr", "b", "e", "d":
	default:
		return cell, fmt.Errorf("unsupported cell type %q", storage)
	}
	for _, metadata := range []string{"cm", "vm"} {
		if _, present, attrErr := unqualifiedXMLAttribute(root, metadata); attrErr != nil {
			return cell, attrErr
		} else if present {
			cell.Editable = false
			if err := extractor.addUnsupported("CELL_METADATA", "cell-metadata", "sheet:"+route.id, route.part, ref, "cell metadata is preserved and makes the cell mutation-refused"); err != nil {
				return cell, err
			}
		}
	}
	var rawValue *string
	var inlineValue *nativeSharedString
	var parsedFormula *nativeFormulaParse
	seenFormula, seenValue, seenInline := false, false, false
	for {
		token, err := decoder.Token()
		if err != nil {
			return cell, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name.Space == extractor.namespace {
				switch token.Name.Local {
				case "f":
					if seenFormula {
						return cell, fmt.Errorf("duplicate formula element")
					}
					formula, parseErr := extractor.parseNativeFormula(decoder, token, route, ref)
					if parseErr != nil {
						return cell, parseErr
					}
					parsedFormula, seenFormula = &formula, true
					continue
				case "v":
					if seenValue {
						return cell, fmt.Errorf("duplicate value element")
					}
					if err := requireOnlySemanticXMLAttributes(token); err != nil {
						return cell, err
					}
					value, parseErr := readNativeRawTextElement(decoder, token)
					if parseErr != nil {
						return cell, parseErr
					}
					rawValue, seenValue = &value, true
					continue
				case "is":
					if seenInline {
						return cell, fmt.Errorf("duplicate inline string element")
					}
					if len(unexpectedSemanticXMLAttributes(token)) != 0 {
						cell.Editable = false
						if err := extractor.addUnsupported("INLINE_STRING_ATTRIBUTES", "rich-text", "sheet:"+route.id, route.part, ref, "unmodeled inline-string attributes remain authority-bound to the source package"); err != nil {
							return cell, err
						}
					}
					value, parseErr := parseNativeStringItem(decoder, token, extractor.namespace)
					if parseErr != nil {
						return cell, parseErr
					}
					if err := extractor.claimNativeText(value.text, maxCellTextLen); err != nil {
						return cell, err
					}
					inlineValue, seenInline = &value, true
					continue
				case "extLst":
					cell.Editable = false
					if err := extractor.addUnsupported("CELL_EXTENSIONS", "extensions", "sheet:"+route.id, route.part, ref, "cell extension markup is preserved exactly"); err != nil {
						return cell, err
					}
					if err := skipNativeXMLElement(decoder, token, 1); err != nil {
						return cell, err
					}
					continue
				}
			}
			cell.Editable = false
			if err := extractor.addUnsupported("OPAQUE_CELL_MARKUP", "cell-markup", "sheet:"+route.id, route.part, ref, "opaque cell markup is preserved and makes the cell mutation-refused"); err != nil {
				return cell, err
			}
			if err := skipNativeXMLElement(decoder, token, 1); err != nil {
				return cell, err
			}
		case xml.EndElement:
			if token.Name != root.Name {
				return cell, fmt.Errorf("cell has mismatched closing element")
			}
			if seenInline && (storage != "inlineStr" || seenValue) {
				return cell, fmt.Errorf("inline string storage has incompatible value children")
			}
			if storage == "inlineStr" && !seenInline {
				if seenValue || parsedFormula != nil {
					return cell, fmt.Errorf("inlineStr cell has no inline string")
				}
				// Excel/openpyxl emit t="inlineStr" with no <is> for a blank cell.
				inlineValue = &nativeSharedString{}
			}
			if storage == "s" && !seenValue {
				return cell, fmt.Errorf("shared-string cell has no index")
			}
			value, valueErr := extractor.nativeCellValue(storage, rawValue, inlineValue, shared)
			if valueErr != nil {
				return cell, valueErr
			}
			if parsedFormula != nil {
				if storage == "inlineStr" || storage == "s" {
					return cell, fmt.Errorf("formula has incompatible %s cell storage", storage)
				}
				parsedFormula.formula.Cached = value
				cell.Formula = &parsedFormula.formula
				cell.Editable = cell.Editable && parsedFormula.editable
			} else {
				cell.Value = value
			}
			if value != nil && value.Rich {
				cell.Editable = false
				if err := extractor.addUnsupported("RICH_CELL_STRING", "rich-text", "sheet:"+route.id, route.part, ref, "rich string text is readable but mutation-refused to prevent formatting loss"); err != nil {
					return cell, err
				}
			}
			return cell, nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return cell, fmt.Errorf("cell contains unsupported direct text")
			}
		case xml.ProcInst:
			return cell, fmt.Errorf("cell contains processing instruction")
		case xml.Directive:
			return cell, fmt.Errorf("cell contains XML directive")
		}
	}
}

func (extractor *nativeWorkbookExtractor) parseNativeFormula(decoder *xml.Decoder, root xml.StartElement, route nativeWorkbookSheetRoute, cellRef string) (nativeFormulaParse, error) {
	formulaType, found, err := unqualifiedXMLAttribute(root, "t")
	if err != nil {
		return nativeFormulaParse{}, err
	}
	if !found {
		formulaType = "normal"
	}
	if formulaType != "normal" && formulaType != "shared" && formulaType != "array" && formulaType != "dataTable" {
		return nativeFormulaParse{}, fmt.Errorf("unsupported formula type %q", formulaType)
	}
	result := nativeFormulaParse{formula: NativeWorkbookFormulaV1{Type: formulaType}, editable: formulaType == "normal"}
	if len(unexpectedSemanticXMLAttributes(root, xml.Name{Local: "t"}, xml.Name{Local: "ref"}, xml.Name{Local: "si"})) != 0 {
		result.editable = false
		if err := extractor.addUnsupported("FORMULA_ATTRIBUTES", "formula-groups", "sheet:"+route.id, route.part, cellRef, "unmodeled formula attributes remain authority-bound to the source package and make the formula mutation-refused"); err != nil {
			return result, err
		}
	}
	if ref, present, attrErr := unqualifiedXMLAttribute(root, "ref"); attrErr != nil {
		return result, attrErr
	} else if present {
		canonicalRef, parseErr := canonicalNativeRangeReference(ref)
		if parseErr != nil {
			return result, fmt.Errorf("invalid formula ref %q: %w", ref, parseErr)
		}
		result.formula.Ref = nativeWorkbookString(canonicalRef)
	}
	if sharedIndex, present, attrErr := optionalNativeUintAttribute(root, "si", ^uint64(0)>>32); attrErr != nil {
		return result, attrErr
	} else if present {
		value := uint32(sharedIndex)
		result.formula.SharedIndex = &value
	}
	if formulaType == "shared" && result.formula.SharedIndex == nil {
		return result, fmt.Errorf("shared formula requires si")
	}
	if (formulaType == "array" || formulaType == "dataTable") && result.formula.Ref == nil {
		return result, fmt.Errorf("%s formula requires ref", formulaType)
	}
	if formulaType != "shared" && result.formula.SharedIndex != nil {
		return result, fmt.Errorf("%s formula cannot carry si", formulaType)
	}
	if formulaType == "normal" && result.formula.Ref != nil {
		return result, fmt.Errorf("normal formula cannot carry ref")
	}
	text, err := readNativeRawTextElement(decoder, root)
	if err != nil {
		return result, err
	}
	text, err = decodeSpreadsheetString(text)
	if err != nil {
		return result, err
	}
	if formulaType == "normal" && text == "" {
		return result, fmt.Errorf("normal formula text is empty")
	}
	if formulaType == "shared" {
		if text == "" && result.formula.Ref != nil {
			return result, fmt.Errorf("shared formula follower cannot carry ref")
		}
		if text != "" && result.formula.Ref == nil {
			return result, fmt.Errorf("shared formula master requires ref")
		}
	}
	if formulaType == "array" && text == "" {
		return result, fmt.Errorf("array formula text is empty")
	}
	if err := extractor.claimNativeText(text, maxFormulaLen-1); err != nil {
		return result, err
	}
	result.formula.Text = text
	if formulaType != "normal" {
		code := "FORMULA_" + strings.ToUpper(formulaType)
		if err := extractor.addUnsupported(code, "formula-groups", "sheet:"+route.id, route.part, cellRef, "shared/array/data-table formulas are surfaced but mutation-refused without group-wide fidelity"); err != nil {
			return result, err
		}
	}
	return result, nil
}

func (extractor *nativeWorkbookExtractor) nativeCellValue(storage string, raw *string, inline *nativeSharedString, shared []nativeSharedString) (*NativeWorkbookValueV1, error) {
	if storage == "inlineStr" {
		if inline == nil {
			return nil, fmt.Errorf("missing inline string")
		}
		return &NativeWorkbookValueV1{Kind: "string", Storage: "inline", Text: nativeWorkbookString(inline.text), Rich: inline.rich, runs: cloneRichRuns(inline.runs)}, nil
	}
	if raw == nil {
		return nil, nil
	}
	if err := extractor.claimNativeText(*raw, maxCellTextLen); err != nil {
		return nil, err
	}
	value := &NativeWorkbookValueV1{Lexical: nativeWorkbookString(*raw), Rich: false}
	switch storage {
	case "n":
		if *raw == "" {
			return nil, nil
		}
		if _, err := finiteNativeFloat(*raw, -1.7976931348623157e308, 1.7976931348623157e308); err != nil {
			return nil, fmt.Errorf("numeric lexical value: %w", err)
		}
		value.Kind, value.Storage = "number", "number"
	case "b":
		if *raw != "0" && *raw != "1" && *raw != "false" && *raw != "true" {
			return nil, fmt.Errorf("boolean lexical value %q is invalid", *raw)
		}
		value.Kind, value.Storage = "boolean", "boolean"
	case "e":
		if *raw == "" {
			return nil, fmt.Errorf("error lexical value is empty")
		}
		value.Kind, value.Storage = "error", "error"
	case "d":
		if !validNativeISODateTime(*raw) {
			return nil, fmt.Errorf("date lexical value %q is not ISO-8601 date/dateTime", *raw)
		}
		value.Kind, value.Storage = "date", "date"
	case "str":
		text, err := decodeSpreadsheetString(*raw)
		if err != nil {
			return nil, err
		}
		value.Kind, value.Storage, value.Text = "string", "formula-string", nativeWorkbookString(text)
	case "s":
		index, err := parseNativePositiveOrZeroIndex(*raw, len(shared))
		if err != nil {
			return nil, fmt.Errorf("shared-string index: %w", err)
		}
		entry := shared[index]
		value.Kind, value.Storage, value.Text, value.Rich, value.runs = "string", "shared", nativeWorkbookString(entry.text), entry.rich, cloneRichRuns(entry.runs)
	default:
		return nil, fmt.Errorf("unsupported cell storage %q", storage)
	}
	return value, nil
}

type nativeFormulaGroupRange struct {
	minimumRow, minimumColumn, maximumRow, maximumColumn int
	ref                                                  string
}

func (extractor *nativeWorkbookExtractor) markNativeFormulaGroupFollowers(sheet *NativeWorkbookSheetV1) error {
	ranges := make([]nativeFormulaGroupRange, 0)
	sharedMasters := map[uint32]nativeFormulaGroupRange{}
	type sharedFollower struct {
		index       uint32
		row, column int
		ref         string
	}
	sharedFollowers := make([]sharedFollower, 0)
	hasGroups := false
	for index := range sheet.Cells {
		cell := &sheet.Cells[index]
		formula := cell.Formula
		if formula == nil || formula.Type == "normal" {
			continue
		}
		hasGroups = true
		cell.Editable = false
		if formula.Ref == nil {
			if formula.Type == "shared" && formula.SharedIndex != nil {
				sharedFollowers = append(sharedFollowers, sharedFollower{index: *formula.SharedIndex, row: cell.Row, column: cell.Column, ref: cell.Ref})
			}
			continue
		}
		minimumRow, minimumColumn, maximumRow, maximumColumn, err := parseDimensionReference(*formula.Ref)
		if err != nil {
			return err
		}
		if cell.Row < minimumRow || cell.Row > maximumRow || cell.Column < minimumColumn || cell.Column > maximumColumn {
			return fmt.Errorf("formula group master %s is outside its range %q", cell.Ref, *formula.Ref)
		}
		group := nativeFormulaGroupRange{minimumRow: minimumRow, minimumColumn: minimumColumn, maximumRow: maximumRow, maximumColumn: maximumColumn, ref: *formula.Ref}
		ranges = append(ranges, group)
		if formula.Type == "shared" && formula.SharedIndex != nil {
			if _, duplicate := sharedMasters[*formula.SharedIndex]; duplicate {
				return fmt.Errorf("shared formula si=%d has multiple masters", *formula.SharedIndex)
			}
			sharedMasters[*formula.SharedIndex] = group
		}
		if len(ranges) > NativeXLSXMaxInventory {
			return fmt.Errorf("formula group inventory exceeds %d ranges", NativeXLSXMaxInventory)
		}
	}
	if !hasGroups {
		return nil
	}
	sheet.Editable = false
	refusal := "FORMULA_GROUPS"
	sheet.RefusalCode = &refusal
	for _, follower := range sharedFollowers {
		master, found := sharedMasters[follower.index]
		if !found {
			return fmt.Errorf("shared formula follower %s references missing master si=%d", follower.ref, follower.index)
		}
		if follower.row < master.minimumRow || follower.row > master.maximumRow || follower.column < master.minimumColumn || follower.column > master.maximumColumn {
			return fmt.Errorf("shared formula follower %s is outside master si=%d range %q", follower.ref, follower.index, master.ref)
		}
	}
	if err := rejectOverlappingNativeFormulaRanges(ranges); err != nil {
		return err
	}
	markCellsInNativeFormulaRanges(sheet.Cells, ranges)
	for _, group := range ranges {
		if err := extractor.addUnsupportedRange("FORMULA_GROUP_RANGE", "formula-groups", "sheet:"+sheet.ID, sheet.PartName, group.ref, "the exact formula-group range is authority-bound to the source package; the complete sheet is mutation-refused"); err != nil {
			return err
		}
	}
	return nil
}

func markCellsInNativeFormulaRanges(cells []NativeWorkbookCellV1, ranges []nativeFormulaGroupRange) {
	type event struct {
		row, minimumColumn, maximumColumn, delta int
	}
	events := make([]event, 0, len(ranges)*2)
	for _, group := range ranges {
		events = append(events,
			event{row: group.minimumRow, minimumColumn: group.minimumColumn, maximumColumn: group.maximumColumn, delta: 1},
			event{row: group.maximumRow + 1, minimumColumn: group.minimumColumn, maximumColumn: group.maximumColumn, delta: -1},
		)
	}
	sort.Slice(events, func(i, j int) bool {
		if events[i].row != events[j].row {
			return events[i].row < events[j].row
		}
		return events[i].delta < events[j].delta
	})
	tree, eventIndex := newNativeColumnRangeTree(excelMaxColumns), 0
	for index := range cells {
		for eventIndex < len(events) && events[eventIndex].row <= cells[index].Row {
			item := events[eventIndex]
			tree.add(item.minimumColumn, item.maximumColumn, item.delta)
			eventIndex++
		}
		if tree.maximum(cells[index].Column, cells[index].Column) > 0 {
			cells[index].Editable = false
		}
	}
}

func rejectOverlappingNativeFormulaRanges(ranges []nativeFormulaGroupRange) error {
	type event struct {
		row, minimumColumn, maximumColumn, delta int
		ref                                      string
	}
	events := make([]event, 0, len(ranges)*2)
	for _, group := range ranges {
		events = append(events,
			event{row: group.minimumRow, minimumColumn: group.minimumColumn, maximumColumn: group.maximumColumn, delta: 1, ref: group.ref},
			event{row: group.maximumRow + 1, minimumColumn: group.minimumColumn, maximumColumn: group.maximumColumn, delta: -1, ref: group.ref},
		)
	}
	sort.Slice(events, func(i, j int) bool {
		if events[i].row != events[j].row {
			return events[i].row < events[j].row
		}
		return events[i].delta < events[j].delta
	})
	tree := newNativeColumnRangeTree(excelMaxColumns)
	for _, item := range events {
		if item.delta > 0 && tree.maximum(item.minimumColumn, item.maximumColumn) > 0 {
			return fmt.Errorf("formula group range %q overlaps another formula group", item.ref)
		}
		tree.add(item.minimumColumn, item.maximumColumn, item.delta)
	}
	return nil
}

type nativeColumnRangeTree struct {
	maximums []int
	lazy     []int
	size     int
}

func newNativeColumnRangeTree(size int) *nativeColumnRangeTree {
	return &nativeColumnRangeTree{maximums: make([]int, size*4), lazy: make([]int, size*4), size: size}
}

func (tree *nativeColumnRangeTree) add(left, right, delta int) {
	tree.addNode(1, 0, tree.size-1, left, right, delta)
}

func (tree *nativeColumnRangeTree) addNode(node, start, end, left, right, delta int) {
	if left <= start && end <= right {
		tree.maximums[node] += delta
		tree.lazy[node] += delta
		return
	}
	middle := (start + end) / 2
	if left <= middle {
		tree.addNode(node*2, start, middle, left, right, delta)
	}
	if right > middle {
		tree.addNode(node*2+1, middle+1, end, left, right, delta)
	}
	tree.maximums[node] = tree.lazy[node] + max(tree.maximums[node*2], tree.maximums[node*2+1])
}

func (tree *nativeColumnRangeTree) maximum(left, right int) int {
	return tree.maximumNode(1, 0, tree.size-1, left, right, 0)
}

func (tree *nativeColumnRangeTree) maximumNode(node, start, end, left, right, inherited int) int {
	inherited += tree.lazy[node]
	if left <= start && end <= right {
		return inherited + tree.maximums[node] - tree.lazy[node]
	}
	middle, result := (start+end)/2, 0
	if left <= middle {
		result = tree.maximumNode(node*2, start, middle, left, right, inherited)
	}
	if right > middle {
		result = max(result, tree.maximumNode(node*2+1, middle+1, end, left, right, inherited))
	}
	return result
}

func readNativeRawTextElement(decoder *xml.Decoder, root xml.StartElement) (string, error) {
	var text strings.Builder
	for {
		token, err := decoder.Token()
		if err != nil {
			return "", err
		}
		switch token := token.(type) {
		case xml.CharData:
			text.Write(token)
		case xml.EndElement:
			if token.Name != root.Name {
				return "", fmt.Errorf("text element has mismatched closing element")
			}
			return text.String(), nil
		case xml.StartElement:
			return "", fmt.Errorf("text element contains nested markup")
		case xml.ProcInst:
			return "", fmt.Errorf("text element contains processing instruction")
		case xml.Directive:
			return "", fmt.Errorf("text element contains XML directive")
		}
	}
}

func requireNativeEmptyElement(decoder *xml.Decoder, root xml.StartElement) error {
	for {
		token, err := decoder.Token()
		if err != nil {
			return err
		}
		switch token := token.(type) {
		case xml.EndElement:
			if token.Name != root.Name {
				return fmt.Errorf("element has mismatched closing tag")
			}
			return nil
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				return fmt.Errorf("element must be empty")
			}
		case xml.Comment:
			continue
		default:
			return fmt.Errorf("element must be empty")
		}
	}
}

func nativeBooleanAttribute(start xml.StartElement, name string, fallback bool) (bool, error) {
	raw, found, err := unqualifiedXMLAttribute(start, name)
	if err != nil || !found {
		return fallback, err
	}
	return ooxmlBoolean(raw, true)
}

func optionalNativeStyleID(start xml.StartElement, name string, styleCount int, hasStyles bool) (uint32, bool, error) {
	value, found, err := optionalNativeUintAttribute(start, name, uint64(^uint32(0)))
	if err != nil || !found {
		return 0, found, err
	}
	if !hasStyles && value != 0 {
		return 0, true, fmt.Errorf("style id %d requires a styles part", value)
	}
	if value >= uint64(styleCount) {
		return 0, true, fmt.Errorf("style id %d is outside cellXfs", value)
	}
	return uint32(value), true, nil
}

func parseNativePositiveIndex(raw string, maximum int) (int, error) {
	value, err := parseNativeUnsigned(raw)
	if err != nil || value == 0 || value > uint64(maximum) {
		return 0, fmt.Errorf("%q is outside 1..%d", raw, maximum)
	}
	return int(value), nil
}

func parseNativePositiveOrZeroIndex(raw string, maximum int) (int, error) {
	value, err := parseNativeUnsigned(raw)
	if err != nil || value >= uint64(maximum) {
		return 0, fmt.Errorf("%q is outside 0..%d", raw, maximum-1)
	}
	return int(value), nil
}

func parseNativeUnsigned(raw string) (uint64, error) {
	if raw == "" {
		return 0, fmt.Errorf("empty unsigned integer")
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, fmt.Errorf("%q is not an unsigned integer", raw)
		}
	}
	return strconv.ParseUint(raw, 10, 64)
}
