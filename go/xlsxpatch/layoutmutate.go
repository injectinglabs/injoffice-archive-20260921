package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

// LayoutMutationKind is the row/column sizing subset of
// @injoffice/sheets protocol v1.
type LayoutMutationKind string

const (
	RowSetHeight   LayoutMutationKind = "row.set_height"
	ColumnSetWidth LayoutMutationKind = "column.set_width"
)

const (
	maxRowHeightPoints = 409.5
	maxColumnWidth     = 255.0
)

// LayoutMutation changes one zero-based row height or column width. A zero
// dimension hides the row/column; a positive dimension explicitly unhides it.
type LayoutMutation struct {
	OperationID  string             `json:"operation_id"`
	SheetID      string             `json:"sheet_id"`
	Kind         LayoutMutationKind `json:"kind"`
	Row          int                `json:"row,omitempty"`
	HeightPoints float64            `json:"height_points,omitempty"`
	Column       int                `json:"column,omitempty"`
	Width        float64            `json:"width,omitempty"`
}

type normalizedLayoutMutation struct {
	LayoutMutation
}

// ApplyLayoutMutations applies row.set_height and column.set_width directly
// to native worksheet XML. The batch is validated and routed atomically;
// unrelated OPC parts are raw-copied and verified by Apply.
func ApplyLayoutMutations(orig []byte, mutations []LayoutMutation) ([]byte, error) {
	normalized, err := validateLayoutMutations(mutations)
	if err != nil {
		return nil, err
	}

	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: layout mutations: open original: %w", err)
	}
	packageIndex, err := newOPCPackageIndex(zr)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: layout mutations: %w", err)
	}
	files := packageIndex.byExact
	metadata := make(map[string]string, 2)
	var metadataReadErr error
	read := func(name string) (string, bool) {
		if value, ok := metadata[name]; ok {
			return value, true
		}
		file, ok := files[name]
		if !ok || metadataReadErr != nil {
			return "", false
		}
		data, err := readZipFile(file)
		if err != nil {
			metadataReadErr = fmt.Errorf("read %q: %w", name, err)
			return "", false
		}
		value := string(data)
		metadata[name] = value
		return value, true
	}
	workbook, err := locateWorkbookPart(packageIndex, read)
	if metadataReadErr != nil {
		return nil, fmt.Errorf("xlsxpatch: layout mutations: %w", metadataReadErr)
	}
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: layout mutations: %w", err)
	}

	byPart := make(map[string][]normalizedLayoutMutation)
	partOrder := make([]string, 0)
	resolvedSheets := make(map[string]string)
	for _, mutation := range normalized {
		part, ok := resolvedSheets[mutation.SheetID]
		if !ok {
			part, err = worksheetPartForID(packageIndex, read, workbook, mutation.SheetID)
			if metadataReadErr != nil {
				return nil, fmt.Errorf("xlsxpatch: layout mutations: operation %q: %w", mutation.OperationID, metadataReadErr)
			}
			if err != nil {
				return nil, fmt.Errorf("xlsxpatch: layout mutations: operation %q: %w", mutation.OperationID, err)
			}
			resolvedSheets[mutation.SheetID] = part
			if _, seen := byPart[part]; !seen {
				partOrder = append(partOrder, part)
			}
		}
		byPart[part] = append(byPart[part], mutation)
	}

	patch := Patch{Replace: make(map[string][]byte, len(byPart))}
	for _, part := range partOrder {
		file, ok := files[part]
		if !ok {
			return nil, fmt.Errorf("xlsxpatch: layout mutations: worksheet part %q is missing", part)
		}
		sheetXML, err := readZipFile(file)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: layout mutations: read worksheet part %q: %w", part, err)
		}
		updated, err := patchWorksheetLayout(sheetXML, byPart[part])
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: layout mutations: worksheet %q: %w", part, err)
		}
		if !bytes.Equal(updated, sheetXML) {
			patch.Replace[part] = updated
		}
	}
	if len(patch.Replace) == 0 {
		return bytes.Clone(orig), nil
	}
	return Apply(orig, patch)
}

func validateLayoutMutations(mutations []LayoutMutation) ([]normalizedLayoutMutation, error) {
	if len(mutations) == 0 {
		return nil, fmt.Errorf("xlsxpatch: layout mutations: empty batch")
	}
	if len(mutations) > maxCellMutations {
		return nil, fmt.Errorf("xlsxpatch: layout mutations: batch exceeds %d operations", maxCellMutations)
	}
	seenIDs := make(map[string]bool, len(mutations))
	normalized := make([]normalizedLayoutMutation, 0, len(mutations))
	for index, mutation := range mutations {
		prefix := fmt.Sprintf("xlsxpatch: layout mutations: operation %d", index)
		if mutation.OperationID == "" || len(mutation.OperationID) > maxOperationIDLen || !restrictedIDPattern.MatchString(mutation.OperationID) {
			return nil, fmt.Errorf("%s: invalid operation_id %q", prefix, mutation.OperationID)
		}
		if seenIDs[mutation.OperationID] {
			return nil, fmt.Errorf("%s: duplicate operation_id %q", prefix, mutation.OperationID)
		}
		seenIDs[mutation.OperationID] = true
		if mutation.SheetID == "" || !utf8.ValidString(mutation.SheetID) || utf16Length(mutation.SheetID) > maxStableSheetIDLen || strings.TrimSpace(mutation.SheetID) != mutation.SheetID {
			return nil, fmt.Errorf("%s: invalid sheet_id", prefix)
		}
		switch mutation.Kind {
		case RowSetHeight:
			if mutation.Column != 0 || mutation.Width != 0 || math.Signbit(mutation.Width) {
				return nil, fmt.Errorf("%s: row.set_height must not carry column or width", prefix)
			}
			if mutation.Row < 0 || mutation.Row >= excelMaxRows {
				return nil, fmt.Errorf("%s: row %d is outside Excel limits", prefix, mutation.Row)
			}
			if math.IsNaN(mutation.HeightPoints) || math.IsInf(mutation.HeightPoints, 0) || mutation.HeightPoints < 0 || mutation.HeightPoints > maxRowHeightPoints || (mutation.HeightPoints == 0 && math.Signbit(mutation.HeightPoints)) {
				return nil, fmt.Errorf("%s: height_points must be finite and within 0..%g", prefix, maxRowHeightPoints)
			}
		case ColumnSetWidth:
			if mutation.Row != 0 || mutation.HeightPoints != 0 || math.Signbit(mutation.HeightPoints) {
				return nil, fmt.Errorf("%s: column.set_width must not carry row or height_points", prefix)
			}
			if mutation.Column < 0 || mutation.Column >= excelMaxColumns {
				return nil, fmt.Errorf("%s: column %d is outside Excel limits", prefix, mutation.Column)
			}
			if math.IsNaN(mutation.Width) || math.IsInf(mutation.Width, 0) || mutation.Width < 0 || mutation.Width > maxColumnWidth || (mutation.Width == 0 && math.Signbit(mutation.Width)) {
				return nil, fmt.Errorf("%s: width must be finite and within 0..%g", prefix, maxColumnWidth)
			}
		default:
			return nil, fmt.Errorf("%s: unsupported kind %q", prefix, mutation.Kind)
		}
		normalized = append(normalized, normalizedLayoutMutation{LayoutMutation: mutation})
	}
	return normalized, nil
}

type layoutTextEdit struct {
	start, end int
	data       []byte
}

type layoutRow struct {
	xmlSpan
	row int
}

type layoutColumn struct {
	xmlSpan
	min, max int
}

type layoutColumns struct {
	xmlSpan
	columns []layoutColumn
}

type worksheetLayoutIndex struct {
	root      xmlSpan
	namespace string
	sheetData xmlSpan
	rows      []layoutRow
	cols      *layoutColumns
}

func patchWorksheetLayout(data []byte, mutations []normalizedLayoutMutation) ([]byte, error) {
	rowHeights := make(map[int]float64)
	columnWidths := make(map[int]float64)
	for _, mutation := range mutations {
		switch mutation.Kind {
		case RowSetHeight:
			rowHeights[mutation.Row] = mutation.HeightPoints
		case ColumnSetWidth:
			columnWidths[mutation.Column] = mutation.Width
		}
	}

	index, err := indexWorksheetLayout(data)
	if err != nil {
		return nil, err
	}
	edits := make([]layoutTextEdit, 0, len(rowHeights)+1)
	rowEdits, err := buildRowLayoutEdits(data, index, rowHeights)
	if err != nil {
		return nil, err
	}
	edits = append(edits, rowEdits...)
	if len(columnWidths) > 0 {
		columnEdit, err := buildColumnLayoutEdit(data, index, columnWidths)
		if err != nil {
			return nil, err
		}
		edits = append(edits, columnEdit)
	}
	if len(edits) == 0 {
		return bytes.Clone(data), nil
	}

	sort.Slice(edits, func(i, j int) bool {
		if edits[i].start == edits[j].start {
			return edits[i].end > edits[j].end
		}
		return edits[i].start > edits[j].start
	})
	updated := bytes.Clone(data)
	for _, edit := range edits {
		updated = append(updated[:edit.start], append(edit.data, updated[edit.end:]...)...)
	}
	if err := validateWorksheetXML(updated); err != nil {
		return nil, fmt.Errorf("produced invalid worksheet XML: %w", err)
	}
	return updated, nil
}

func indexWorksheetLayout(data []byte) (worksheetLayoutIndex, error) {
	index := worksheetLayoutIndex{}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootClosed := false
	sheetDataDepth, colsDepth := 0, 0
	lastRow, lastColumnMax := 0, 0
	var activeRow *layoutRow
	var activeColumn *layoutColumn
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return worksheetLayoutIndex{}, fmt.Errorf("parse worksheet XML: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if token.Name.Local != "worksheet" || !isSpreadsheetMLNamespace(token.Name.Space) {
					return worksheetLayoutIndex{}, fmt.Errorf("root element is not a supported SpreadsheetML worksheet")
				}
				index.namespace = token.Name.Space
				index.root = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				continue
			}
			if activeColumn != nil {
				return worksheetLayoutIndex{}, fmt.Errorf("column %d:%d has unsupported nested markup", activeColumn.min, activeColumn.max)
			}
			if depth == 2 && token.Name.Space == index.namespace {
				switch token.Name.Local {
				case "cols":
					if index.cols != nil {
						return worksheetLayoutIndex{}, fmt.Errorf("multiple cols elements")
					}
					if index.sheetData.end != 0 || sheetDataDepth != 0 {
						return worksheetLayoutIndex{}, fmt.Errorf("cols element appears after sheetData")
					}
					index.cols = &layoutColumns{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}}
					colsDepth = depth
					continue
				case "sheetData":
					if index.sheetData.startTagEnd != 0 {
						return worksheetLayoutIndex{}, fmt.Errorf("multiple sheetData elements")
					}
					index.sheetData = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
					sheetDataDepth = depth
					continue
				case "col":
					return worksheetLayoutIndex{}, fmt.Errorf("col element appears outside cols")
				}
			}
			if colsDepth != 0 && depth == colsDepth+1 {
				if token.Name.Space != index.namespace || token.Name.Local != "col" {
					return worksheetLayoutIndex{}, fmt.Errorf("cols has unsupported direct child %q", rawQName(data, before, after))
				}
				minimum, maximum, err := parseCanonicalColumnRange(data[before:after])
				if err != nil {
					return worksheetLayoutIndex{}, err
				}
				if minimum <= lastColumnMax {
					return worksheetLayoutIndex{}, fmt.Errorf("column definitions overlap or are not strictly ordered at %d:%d", minimum, maximum)
				}
				lastColumnMax = maximum
				activeColumn = &layoutColumn{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, min: minimum, max: maximum}
				continue
			}
			if sheetDataDepth != 0 && depth == sheetDataDepth+1 {
				if token.Name.Space != index.namespace || token.Name.Local != "row" {
					return worksheetLayoutIndex{}, fmt.Errorf("sheetData has unsupported direct child %q", rawQName(data, before, after))
				}
				rowNumber, found, err := optionalPositiveIndexAttribute(token, "r", excelMaxRows)
				if err != nil {
					return worksheetLayoutIndex{}, fmt.Errorf("row: %w", err)
				}
				if !found {
					rowNumber = lastRow + 1
					if rowNumber > excelMaxRows {
						return worksheetLayoutIndex{}, fmt.Errorf("inferred row number exceeds Excel limits")
					}
				}
				if rowNumber <= lastRow {
					return worksheetLayoutIndex{}, fmt.Errorf("rows are not strictly ordered at row %d", rowNumber)
				}
				lastRow = rowNumber
				activeRow = &layoutRow{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, row: rowNumber - 1}
			}
		case xml.EndElement:
			if activeColumn != nil && depth == colsDepth+1 && token.Name.Space == index.namespace && token.Name.Local == "col" {
				activeColumn.endStart, activeColumn.end = before, after
				index.cols.columns = append(index.cols.columns, *activeColumn)
				activeColumn = nil
			}
			if activeRow != nil && depth == sheetDataDepth+1 && token.Name.Space == index.namespace && token.Name.Local == "row" {
				activeRow.endStart, activeRow.end = before, after
				index.rows = append(index.rows, *activeRow)
				activeRow = nil
			}
			if colsDepth != 0 && depth == colsDepth && token.Name.Space == index.namespace && token.Name.Local == "cols" {
				index.cols.endStart, index.cols.end = before, after
				colsDepth = 0
			}
			if sheetDataDepth != 0 && depth == sheetDataDepth && token.Name.Space == index.namespace && token.Name.Local == "sheetData" {
				index.sheetData.endStart, index.sheetData.end = before, after
				sheetDataDepth = 0
			}
			if depth == 1 && token.Name.Space == index.namespace && token.Name.Local == "worksheet" {
				index.root.endStart, index.root.end = before, after
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if len(bytes.TrimSpace(token)) == 0 {
				continue
			}
			if activeColumn != nil {
				return worksheetLayoutIndex{}, fmt.Errorf("column %d:%d has unsupported text content", activeColumn.min, activeColumn.max)
			}
			if colsDepth != 0 && depth == colsDepth {
				return worksheetLayoutIndex{}, fmt.Errorf("cols has unsupported text content")
			}
			if sheetDataDepth != 0 && depth == sheetDataDepth {
				return worksheetLayoutIndex{}, fmt.Errorf("sheetData has unsupported text content")
			}
		}
	}
	if depth != 0 || !rootClosed || index.root.end == 0 {
		return worksheetLayoutIndex{}, fmt.Errorf("missing complete worksheet root")
	}
	if index.sheetData.end == 0 {
		return worksheetLayoutIndex{}, fmt.Errorf("missing complete sheetData element")
	}
	if index.cols != nil && index.cols.end == 0 {
		return worksheetLayoutIndex{}, fmt.Errorf("missing complete cols element")
	}
	return index, nil
}

func optionalPositiveIndexAttribute(start xml.StartElement, name string, maximum int) (int, bool, error) {
	value, found := "", false
	for _, attribute := range start.Attr {
		if attribute.Name.Space != "" || attribute.Name.Local != name {
			continue
		}
		if found {
			return 0, false, fmt.Errorf("duplicate unqualified attribute %q", name)
		}
		value, found = attribute.Value, true
	}
	if !found {
		return 0, false, nil
	}
	parsed, err := strconv.Atoi(value)
	if err != nil || parsed < 1 || parsed > maximum {
		return 0, false, fmt.Errorf("attribute %s=%q is outside 1..%d", name, value, maximum)
	}
	return parsed, true, nil
}

func parseCanonicalColumnRange(tag []byte) (int, int, error) {
	minimum, err := canonicalColumnIndexAttribute(tag, "min")
	if err != nil {
		return 0, 0, fmt.Errorf("column min: %w", err)
	}
	maximum, err := canonicalColumnIndexAttribute(tag, "max")
	if err != nil {
		return 0, 0, fmt.Errorf("column %d max: %w", minimum, err)
	}
	if minimum > maximum {
		return 0, 0, fmt.Errorf("column definition %d:%d has reversed bounds", minimum, maximum)
	}
	return minimum, maximum, nil
}

func canonicalColumnIndexAttribute(tag []byte, name string) (int, error) {
	attribute, found, err := rawUnqualifiedAttribute(tag, name)
	if err != nil {
		return 0, err
	}
	if !found {
		return 0, fmt.Errorf("missing unqualified attribute %q", name)
	}
	raw := string(tag[attribute.valueStart:attribute.valueEnd])
	if raw == "" || (len(raw) > 1 && raw[0] == '0') {
		return 0, fmt.Errorf("attribute %s=%q is not a canonical positive integer", name, raw)
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, fmt.Errorf("attribute %s=%q is not a canonical positive integer", name, raw)
		}
	}
	value, err := strconv.Atoi(raw)
	if err != nil || value < 1 || value > excelMaxColumns {
		return 0, fmt.Errorf("attribute %s=%q is outside 1..%d", name, raw, excelMaxColumns)
	}
	return value, nil
}

func buildRowLayoutEdits(data []byte, index worksheetLayoutIndex, heights map[int]float64) ([]layoutTextEdit, error) {
	if len(heights) == 0 {
		return nil, nil
	}
	targets := sortedIntKeys(heights)
	existing := make(map[int]layoutRow, len(index.rows))
	for _, row := range index.rows {
		existing[row.row] = row
	}
	edits := make([]layoutTextEdit, 0, len(targets))
	insertions := make(map[int][]int)
	for _, target := range targets {
		if row, ok := existing[target]; ok {
			matches, err := sizedStartTagMatches(data[row.start:row.startTagEnd], "ht", heights[target])
			if err != nil {
				return nil, fmt.Errorf("row %d: %w", target+1, err)
			}
			if matches {
				continue
			}
			updated, err := rewriteSizedStartTag(data[row.start:row.startTagEnd], "ht", heights[target])
			if err != nil {
				return nil, fmt.Errorf("row %d: %w", target+1, err)
			}
			if !bytes.Equal(updated, data[row.start:row.startTagEnd]) {
				edits = append(edits, layoutTextEdit{start: row.start, end: row.startTagEnd, data: updated})
			}
			continue
		}
		insertAt := index.sheetData.endStart
		for _, row := range index.rows {
			if row.row > target {
				insertAt = row.start
				break
			}
		}
		insertions[insertAt] = append(insertions[insertAt], target)
	}

	if index.sheetData.selfClosing(data) {
		var rows bytes.Buffer
		for _, target := range targets {
			built, err := buildNewSizedRow(prefixedLocal(index.sheetData.qname, "row"), target, heights[target])
			if err != nil {
				return nil, err
			}
			rows.Write(built)
		}
		start := openTag(data[index.sheetData.start:index.sheetData.startTagEnd])
		replacement := append(start, rows.Bytes()...)
		replacement = append(replacement, []byte("</"+index.sheetData.qname+">")...)
		return []layoutTextEdit{{start: index.sheetData.start, end: index.sheetData.end, data: replacement}}, nil
	}

	for offset, rows := range insertions {
		sort.Ints(rows)
		var builtRows bytes.Buffer
		for _, row := range rows {
			built, err := buildNewSizedRow(prefixedLocal(index.sheetData.qname, "row"), row, heights[row])
			if err != nil {
				return nil, err
			}
			builtRows.Write(built)
		}
		edits = append(edits, layoutTextEdit{start: offset, end: offset, data: builtRows.Bytes()})
	}
	return edits, nil
}

func rewriteSizedStartTag(tag []byte, dimensionName string, value float64) ([]byte, error) {
	formatted := formatLayoutDimension(value)
	updated, err := rewriteUnqualifiedAttribute(tag, dimensionName, &formatted)
	if err != nil {
		return nil, fmt.Errorf("rewrite %s: %w", dimensionName, err)
	}
	one := "1"
	customName := "customWidth"
	if dimensionName == "ht" {
		customName = "customHeight"
	}
	updated, err = rewriteUnqualifiedAttribute(updated, customName, &one)
	if err != nil {
		return nil, fmt.Errorf("rewrite %s: %w", customName, err)
	}
	if value == 0 {
		updated, err = rewriteUnqualifiedAttribute(updated, "hidden", &one)
	} else {
		updated, err = rewriteUnqualifiedAttribute(updated, "hidden", nil)
	}
	if err != nil {
		return nil, fmt.Errorf("rewrite hidden: %w", err)
	}
	return updated, nil
}

func sizedStartTagMatches(tag []byte, dimensionName string, value float64) (bool, error) {
	start, err := decodeStartElement(tag)
	if err != nil {
		return false, fmt.Errorf("parse sizing attributes: %w", err)
	}
	dimension, found, err := unqualifiedXMLAttribute(start, dimensionName)
	if err != nil {
		return false, err
	}
	if !found {
		return false, nil
	}
	parsed, err := strconv.ParseFloat(dimension, 64)
	if err != nil || math.IsNaN(parsed) || math.IsInf(parsed, 0) {
		return false, fmt.Errorf("invalid %s=%q", dimensionName, dimension)
	}
	if parsed != value {
		return false, nil
	}
	customName := "customWidth"
	if dimensionName == "ht" {
		customName = "customHeight"
	}
	custom, customFound, err := unqualifiedXMLAttribute(start, customName)
	if err != nil {
		return false, err
	}
	customEnabled, err := ooxmlBoolean(custom, customFound)
	if err != nil {
		return false, fmt.Errorf("invalid %s=%q", customName, custom)
	}
	if !customEnabled {
		return false, nil
	}
	hidden, hiddenFound, err := unqualifiedXMLAttribute(start, "hidden")
	if err != nil {
		return false, err
	}
	hiddenEnabled, err := ooxmlBoolean(hidden, hiddenFound)
	if err != nil {
		return false, fmt.Errorf("invalid hidden=%q", hidden)
	}
	return hiddenEnabled == (value == 0), nil
}

func ooxmlBoolean(value string, found bool) (bool, error) {
	if !found {
		return false, nil
	}
	switch value {
	case "1", "true":
		return true, nil
	case "0", "false":
		return false, nil
	default:
		return false, fmt.Errorf("expected 0, 1, false, or true")
	}
}

func buildNewSizedRow(qname string, row int, height float64) ([]byte, error) {
	tag := []byte(fmt.Sprintf("<%s r=\"%d\"/>", qname, row+1))
	return rewriteSizedStartTag(tag, "ht", height)
}

func buildColumnLayoutEdit(data []byte, index worksheetLayoutIndex, widths map[int]float64) (layoutTextEdit, error) {
	targets := sortedIntKeys(widths)
	if index.cols == nil {
		colsQName := prefixedLocal(index.root.qname, "cols")
		colQName := prefixedLocal(colsQName, "col")
		var out bytes.Buffer
		out.WriteString("<" + colsQName + ">")
		for _, target := range targets {
			built, err := buildNewSizedColumn(colQName, target+1, widths[target])
			if err != nil {
				return layoutTextEdit{}, err
			}
			out.Write(built)
		}
		out.WriteString("</" + colsQName + ">")
		return layoutTextEdit{start: index.sheetData.start, end: index.sheetData.start, data: out.Bytes()}, nil
	}

	cols := *index.cols
	colQName := prefixedLocal(cols.qname, "col")
	var out bytes.Buffer
	out.Write(data[cols.start:cols.startTagEnd])
	cursor, targetCursor := cols.startTagEnd, 0
	for _, column := range cols.columns {
		out.Write(data[cursor:column.start])
		for targetCursor < len(targets) && targets[targetCursor]+1 < column.min {
			target := targets[targetCursor]
			built, err := buildNewSizedColumn(colQName, target+1, widths[target])
			if err != nil {
				return layoutTextEdit{}, err
			}
			out.Write(built)
			targetCursor++
		}
		firstTarget := targetCursor
		for targetCursor < len(targets) && targets[targetCursor]+1 <= column.max {
			targetCursor++
		}
		effectiveTargets := make([]int, 0, targetCursor-firstTarget)
		for _, target := range targets[firstTarget:targetCursor] {
			matches, err := sizedStartTagMatches(data[column.start:column.startTagEnd], "width", widths[target])
			if err != nil {
				return layoutTextEdit{}, fmt.Errorf("column %d:%d: %w", column.min, column.max, err)
			}
			if !matches {
				effectiveTargets = append(effectiveTargets, target)
			}
		}
		if len(effectiveTargets) == 0 {
			out.Write(data[column.start:column.end])
		} else {
			rebuilt, err := splitSizedColumn(data[column.start:column.startTagEnd], column, effectiveTargets, widths)
			if err != nil {
				return layoutTextEdit{}, err
			}
			out.Write(rebuilt)
		}
		cursor = column.end
	}
	out.Write(data[cursor:cols.endStart])
	for targetCursor < len(targets) {
		target := targets[targetCursor]
		built, err := buildNewSizedColumn(colQName, target+1, widths[target])
		if err != nil {
			return layoutTextEdit{}, err
		}
		out.Write(built)
		targetCursor++
	}
	out.Write(data[cols.endStart:cols.end])
	if cols.selfClosing(data) {
		out.Reset()
		out.Write(openTag(data[cols.start:cols.startTagEnd]))
		for _, target := range targets {
			built, err := buildNewSizedColumn(colQName, target+1, widths[target])
			if err != nil {
				return layoutTextEdit{}, err
			}
			out.Write(built)
		}
		out.WriteString("</" + cols.qname + ">")
	}
	return layoutTextEdit{start: cols.start, end: cols.end, data: out.Bytes()}, nil
}

func splitSizedColumn(tag []byte, column layoutColumn, targets []int, widths map[int]float64) ([]byte, error) {
	var out bytes.Buffer
	cursor := column.min
	for _, zeroBased := range targets {
		target := zeroBased + 1
		if cursor < target {
			segment, err := rewriteColumnRange(tag, cursor, target-1, nil)
			if err != nil {
				return nil, err
			}
			out.Write(segment)
		}
		width := widths[zeroBased]
		segment, err := rewriteColumnRange(tag, target, target, &width)
		if err != nil {
			return nil, err
		}
		out.Write(segment)
		cursor = target + 1
	}
	if cursor <= column.max {
		segment, err := rewriteColumnRange(tag, cursor, column.max, nil)
		if err != nil {
			return nil, err
		}
		out.Write(segment)
	}
	return out.Bytes(), nil
}

func rewriteColumnRange(tag []byte, minimum, maximum int, width *float64) ([]byte, error) {
	minValue, maxValue := strconv.Itoa(minimum), strconv.Itoa(maximum)
	updated, err := rewriteUnqualifiedAttribute(tag, "min", &minValue)
	if err != nil {
		return nil, fmt.Errorf("rewrite column min: %w", err)
	}
	updated, err = rewriteUnqualifiedAttribute(updated, "max", &maxValue)
	if err != nil {
		return nil, fmt.Errorf("rewrite column max: %w", err)
	}
	updated, err = canonicalEmptyElementTag(updated)
	if err != nil {
		return nil, err
	}
	if width == nil {
		return updated, nil
	}
	return rewriteSizedStartTag(updated, "width", *width)
}

func canonicalEmptyElementTag(tag []byte) ([]byte, error) {
	trimmed := bytes.TrimSpace(tag)
	if bytes.HasSuffix(trimmed, []byte("/>")) {
		return bytes.Clone(tag), nil
	}
	closeAt := bytes.LastIndexByte(tag, '>')
	if closeAt < 0 {
		return nil, fmt.Errorf("column start tag has no closer")
	}
	return append(bytes.Clone(tag[:closeAt]), append([]byte{'/'}, tag[closeAt:]...)...), nil
}

func buildNewSizedColumn(qname string, column int, width float64) ([]byte, error) {
	tag := []byte(fmt.Sprintf("<%s min=\"%d\" max=\"%d\"/>", qname, column, column))
	return rewriteSizedStartTag(tag, "width", width)
}

func formatLayoutDimension(value float64) string {
	if value == 0 {
		return "0"
	}
	return strconv.FormatFloat(value, 'g', -1, 64)
}

func sortedIntKeys[T any](values map[int]T) []int {
	keys := make([]int, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	sort.Ints(keys)
	return keys
}
