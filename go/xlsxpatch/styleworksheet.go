package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strconv"
)

type styleWorksheetCell struct {
	xmlSpan
	column    int
	style     int
	styleSeen bool
}

type styleWorksheetRow struct {
	xmlSpan
	row   int
	cells []styleWorksheetCell
}

type styleWorksheetIndex struct {
	namespace    string
	sheetData    xmlSpan
	dimension    *xmlSpan
	dimensionRef string
	rows         []styleWorksheetRow
	byRow        map[int]styleWorksheetRow
}

func patchWorksheetStyles(data []byte, deltas map[cellKey]StyleDelta, registry *styleRegistry) ([]byte, error) {
	index, err := indexStyleWorksheet(data)
	if err != nil {
		return nil, err
	}
	byRow := make(map[int]map[int]int)
	material := make(map[cellKey]bool)
	keys := sortedStyleKeys(deltas)
	for _, key := range keys {
		source, exists := styleWorksheetSource(index, key)
		desired, err := registry.resolveStyle(source, deltas[key])
		if err != nil {
			return nil, fmt.Errorf("cell %s: %w", cellReference(key.row, key.column), err)
		}
		if exists && desired == source {
			continue
		}
		if !exists && desired == 0 {
			continue
		}
		if byRow[key.row] == nil {
			byRow[key.row] = make(map[int]int)
		}
		byRow[key.row][key.column] = desired
		if !exists && desired != 0 {
			material[key] = true
		}
	}
	if len(byRow) == 0 {
		return bytes.Clone(data), nil
	}

	type textEdit struct {
		start, end int
		data       []byte
	}
	edits := make([]textEdit, 0, len(byRow)+1)
	insertRows := make(map[int][]int)
	rowNumbers := make([]int, 0, len(byRow))
	for row := range byRow {
		rowNumbers = append(rowNumbers, row)
	}
	sort.Ints(rowNumbers)
	for _, rowNumber := range rowNumbers {
		if row, exists := index.byRow[rowNumber]; exists {
			updated, err := rebuildStyleRow(data, row, byRow[rowNumber])
			if err != nil {
				return nil, fmt.Errorf("row %d: %w", rowNumber+1, err)
			}
			if !bytes.Equal(updated, data[row.start:row.end]) {
				edits = append(edits, textEdit{start: row.start, end: row.end, data: updated})
			}
			continue
		}
		insertAt := index.sheetData.endStart
		for _, row := range index.rows {
			if row.row > rowNumber {
				insertAt = row.start
				break
			}
		}
		insertRows[insertAt] = append(insertRows[insertAt], rowNumber)
	}
	for offset, rows := range insertRows {
		sort.Ints(rows)
		var built bytes.Buffer
		for _, row := range rows {
			built.Write(buildNewStyleRow(index.sheetData.qname, row, byRow[row]))
		}
		edits = append(edits, textEdit{start: offset, end: offset, data: built.Bytes()})
	}
	if index.sheetData.selfClosing(data) {
		var rows bytes.Buffer
		for _, row := range rowNumbers {
			rows.Write(buildNewStyleRow(index.sheetData.qname, row, byRow[row]))
		}
		start := openTag(data[index.sheetData.start:index.sheetData.startTagEnd])
		replacement := append(start, rows.Bytes()...)
		replacement = append(replacement, []byte("</"+index.sheetData.qname+">")...)
		edits = []textEdit{{start: index.sheetData.start, end: index.sheetData.end, data: replacement}}
	}
	if index.dimension != nil && len(material) > 0 {
		expanded, err := expandStyleDimension(data[index.dimension.start:index.dimension.startTagEnd], index.dimensionRef, material)
		if err != nil {
			return nil, err
		}
		if !bytes.Equal(expanded, data[index.dimension.start:index.dimension.startTagEnd]) {
			edits = append(edits, textEdit{start: index.dimension.start, end: index.dimension.startTagEnd, data: expanded})
		}
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

func indexStyleWorksheet(data []byte) (styleWorksheetIndex, error) {
	index := styleWorksheetIndex{byRow: make(map[int]styleWorksheetRow)}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootClosed, sheetDataSeen := false, false
	sheetDepth, lastRow := 0, 0
	var activeRow *styleWorksheetRow
	rowDepth, lastColumn := 0, 0
	var activeCell *styleWorksheetCell
	cellDepth := 0
	rowSawNonCell := false
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return styleWorksheetIndex{}, fmt.Errorf("parse worksheet XML: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if index.namespace != "" {
					return styleWorksheetIndex{}, fmt.Errorf("worksheet XML has multiple root elements")
				}
				if token.Name.Local != "worksheet" || (token.Name.Space != spreadsheetMLTransitional && token.Name.Space != spreadsheetMLStrict) {
					return styleWorksheetIndex{}, fmt.Errorf("root element is not a supported SpreadsheetML worksheet")
				}
				index.namespace = token.Name.Space
				continue
			}
			if depth == 2 && token.Name.Space == index.namespace && token.Name.Local == "dimension" {
				if index.dimension != nil {
					return styleWorksheetIndex{}, fmt.Errorf("multiple dimension elements")
				}
				span := xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				index.dimension = &span
				index.dimensionRef = attribute(token, "ref")
				continue
			}
			if depth == 2 && token.Name.Space == index.namespace && token.Name.Local == "sheetData" {
				if sheetDataSeen {
					return styleWorksheetIndex{}, fmt.Errorf("multiple sheetData elements")
				}
				sheetDataSeen, sheetDepth = true, depth
				index.sheetData = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				continue
			}
			if sheetDepth != 0 && depth == sheetDepth+1 {
				if token.Name.Space != index.namespace || token.Name.Local != "row" {
					return styleWorksheetIndex{}, fmt.Errorf("sheetData has unsupported direct child %q", rawQName(data, before, after))
				}
				rowNumber, found, err := optionalPositiveIndexAttribute(token, "r", excelMaxRows)
				if err != nil {
					return styleWorksheetIndex{}, fmt.Errorf("row: %w", err)
				}
				if !found {
					rowNumber = lastRow + 1
				}
				if rowNumber <= lastRow || rowNumber > excelMaxRows {
					return styleWorksheetIndex{}, fmt.Errorf("rows are not strictly ordered at row %d", rowNumber)
				}
				lastRow = rowNumber
				activeRow = &styleWorksheetRow{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, row: rowNumber - 1}
				rowDepth, lastColumn, rowSawNonCell = depth, 0, false
				continue
			}
			if activeRow != nil && depth == rowDepth+1 {
				if token.Name.Space == index.namespace && token.Name.Local == "c" {
					if rowSawNonCell {
						return styleWorksheetIndex{}, fmt.Errorf("row %d has a cell after non-cell markup", activeRow.row+1)
					}
					column, found, err := optionalStyleCellReference(token, activeRow.row)
					if err != nil {
						return styleWorksheetIndex{}, fmt.Errorf("row %d: %w", activeRow.row+1, err)
					}
					if !found {
						column = lastColumn
					}
					if column < lastColumn || column >= excelMaxColumns {
						return styleWorksheetIndex{}, fmt.Errorf("cells are not strictly ordered in row %d", activeRow.row+1)
					}
					if len(activeRow.cells) > 0 && column == activeRow.cells[len(activeRow.cells)-1].column {
						return styleWorksheetIndex{}, fmt.Errorf("duplicate cell column %d in row %d", column+1, activeRow.row+1)
					}
					style, styleSeen, err := styleIndexAttribute(data[before:after])
					if err != nil {
						return styleWorksheetIndex{}, fmt.Errorf("cell %s: %w", cellReference(activeRow.row, column), err)
					}
					activeCell = &styleWorksheetCell{xmlSpan: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, column: column, style: style, styleSeen: styleSeen}
					cellDepth, lastColumn = depth, column+1
					continue
				}
				if token.Name.Space != index.namespace || token.Name.Local != "extLst" {
					return styleWorksheetIndex{}, fmt.Errorf("row %d has unsupported direct child %q", activeRow.row+1, rawQName(data, before, after))
				}
				rowSawNonCell = true
			}
		case xml.EndElement:
			if index.dimension != nil && index.dimension.end == 0 && depth == 2 && token.Name.Space == index.namespace && token.Name.Local == "dimension" {
				index.dimension.endStart, index.dimension.end = before, after
			}
			if activeCell != nil && depth == cellDepth && token.Name.Space == index.namespace && token.Name.Local == "c" {
				activeCell.endStart, activeCell.end = before, after
				activeRow.cells = append(activeRow.cells, *activeCell)
				activeCell = nil
			}
			if activeRow != nil && depth == rowDepth && token.Name.Space == index.namespace && token.Name.Local == "row" {
				activeRow.endStart, activeRow.end = before, after
				index.rows = append(index.rows, *activeRow)
				index.byRow[activeRow.row] = *activeRow
				activeRow = nil
			}
			if sheetDepth != 0 && depth == sheetDepth && token.Name.Space == index.namespace && token.Name.Local == "sheetData" {
				index.sheetData.endStart, index.sheetData.end = before, after
				sheetDepth = 0
			}
			if depth == 1 && token.Name.Space == index.namespace && token.Name.Local == "worksheet" {
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if sheetDepth != 0 && depth == sheetDepth && len(bytes.TrimSpace(token)) != 0 {
				return styleWorksheetIndex{}, fmt.Errorf("sheetData has unsupported text")
			}
		case xml.ProcInst:
			if depth == 0 && !rootClosed && token.Target == "xml" {
				continue
			}
			return styleWorksheetIndex{}, fmt.Errorf("worksheet has unsupported processing instruction")
		case xml.Directive:
			return styleWorksheetIndex{}, fmt.Errorf("worksheet has unsupported directive")
		}
	}
	if depth != 0 || !rootClosed || !sheetDataSeen || index.sheetData.end == 0 {
		return styleWorksheetIndex{}, fmt.Errorf("missing complete worksheet/sheetData root")
	}
	return index, nil
}

func optionalStyleCellReference(start xml.StartElement, expectedRow int) (int, bool, error) {
	raw, found, err := unqualifiedXMLAttribute(start, "r")
	if err != nil || !found {
		return 0, found, err
	}
	row, column, err := parseCellReference(raw)
	if err != nil {
		return 0, true, err
	}
	if row != expectedRow {
		return 0, true, fmt.Errorf("cell %q is stored in row %d", raw, expectedRow+1)
	}
	return column, true, nil
}

func styleIndexAttribute(tag []byte) (int, bool, error) {
	span, found, err := rawUnqualifiedAttribute(tag, "s")
	if err != nil || !found {
		return 0, found, err
	}
	raw := string(tag[span.valueStart:span.valueEnd])
	if raw == "" || (len(raw) > 1 && raw[0] == '0') {
		return 0, true, fmt.Errorf("s=%q is not a canonical style index", raw)
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, true, fmt.Errorf("s=%q is not a style index", raw)
		}
	}
	value, err := strconv.Atoi(raw)
	if err != nil {
		return 0, true, fmt.Errorf("s=%q is outside platform limits", raw)
	}
	return value, true, nil
}

func styleWorksheetSource(index styleWorksheetIndex, key cellKey) (int, bool) {
	row, found := index.byRow[key.row]
	if !found {
		return 0, false
	}
	position := sort.Search(len(row.cells), func(i int) bool { return row.cells[i].column >= key.column })
	if position < len(row.cells) && row.cells[position].column == key.column {
		return row.cells[position].style, true
	}
	return 0, false
}

func rebuildStyleRow(data []byte, row styleWorksheetRow, desired map[int]int) ([]byte, error) {
	existing := make(map[int]bool, len(row.cells))
	for _, cell := range row.cells {
		existing[cell.column] = true
	}
	columns := make([]int, 0, len(desired))
	inserting := false
	for column := range desired {
		columns = append(columns, column)
		if !existing[column] {
			inserting = true
		}
	}
	sort.Ints(columns)
	rowStart := bytes.Clone(data[row.start:row.startTagEnd])
	var err error
	if inserting {
		rowStart, err = rewriteUnqualifiedAttribute(rowStart, "spans", nil)
		if err != nil {
			return nil, err
		}
	}
	var out bytes.Buffer
	out.Write(openTag(rowStart))
	cursor, next := row.startTagEnd, 0
	for _, cell := range row.cells {
		for next < len(columns) && columns[next] < cell.column {
			column := columns[next]
			if !existing[column] {
				out.Write(buildStyledCell(prefixedLocal(row.qname, "c"), row.row, column, desired[column]))
			}
			next++
		}
		out.Write(data[cursor:cell.start])
		if style, found := desired[cell.column]; found {
			start := bytes.Clone(data[cell.start:cell.startTagEnd])
			if style == 0 {
				start, err = rewriteUnqualifiedAttribute(start, "s", nil)
			} else {
				value := strconv.Itoa(style)
				start, err = rewriteUnqualifiedAttribute(start, "s", &value)
			}
			if err != nil {
				return nil, err
			}
			out.Write(start)
			out.Write(data[cell.startTagEnd:cell.end])
		} else {
			out.Write(data[cell.start:cell.end])
		}
		cursor = cell.end
		for next < len(columns) && columns[next] == cell.column {
			next++
		}
	}
	for next < len(columns) {
		column := columns[next]
		if !existing[column] {
			out.Write(buildStyledCell(prefixedLocal(row.qname, "c"), row.row, column, desired[column]))
		}
		next++
	}
	if !row.selfClosing(data) {
		out.Write(data[cursor:row.endStart])
	}
	out.WriteString("</" + row.qname + ">")
	return out.Bytes(), nil
}

func buildNewStyleRow(sheetDataQName string, row int, desired map[int]int) []byte {
	columns := make([]int, 0, len(desired))
	for column, style := range desired {
		if style != 0 {
			columns = append(columns, column)
		}
	}
	sort.Ints(columns)
	rowQName := prefixedLocal(sheetDataQName, "row")
	cellQName := prefixedLocal(sheetDataQName, "c")
	var out bytes.Buffer
	fmt.Fprintf(&out, `<%s r="%d">`, rowQName, row+1)
	for _, column := range columns {
		out.Write(buildStyledCell(cellQName, row, column, desired[column]))
	}
	out.WriteString("</" + rowQName + ">")
	return out.Bytes()
}

func buildStyledCell(qname string, row, column, style int) []byte {
	return []byte(fmt.Sprintf(`<%s r="%s" s="%d"/>`, qname, cellReference(row, column), style))
}

func expandStyleDimension(tag []byte, ref string, cells map[cellKey]bool) ([]byte, error) {
	if ref == "" {
		return nil, fmt.Errorf("dimension element is missing ref")
	}
	minRow, minColumn, maxRow, maxColumn, err := parseDimensionReference(ref)
	if err != nil {
		return nil, err
	}
	for key := range cells {
		minRow = min(minRow, key.row)
		minColumn = min(minColumn, key.column)
		maxRow = max(maxRow, key.row)
		maxColumn = max(maxColumn, key.column)
	}
	expanded := cellReference(minRow, minColumn)
	if minRow != maxRow || minColumn != maxColumn {
		expanded += ":" + cellReference(maxRow, maxColumn)
	}
	if expanded == ref {
		return bytes.Clone(tag), nil
	}
	return rewriteUnqualifiedAttribute(tag, "ref", &expanded)
}
