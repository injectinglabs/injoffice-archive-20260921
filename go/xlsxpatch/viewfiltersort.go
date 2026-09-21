package xlsxpatch

import (
	"bytes"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

const (
	SheetFreeze  = "sheet.freeze"
	SheetFilter  = "sheet.filter"
	RangeSort    = "range.sort"
	maxSortRows  = 10_000
	maxSortCells = 100_000
)

type ViewMutation struct {
	OperationID string `json:"operation_id"`
	SheetID     string `json:"sheet_id"`
	Kind        string `json:"kind"`
	Rows        int    `json:"rows"`
	Columns     int    `json:"columns"`
}

type FilterMutation struct {
	OperationID string           `json:"operation_id"`
	SheetID     string           `json:"sheet_id"`
	Kind        string           `json:"kind"`
	Filter      *SheetFilterSpec `json:"filter"`
}

type SheetFilterSpec struct {
	Ref    string   `json:"ref"`
	Column int      `json:"column"`
	Values []string `json:"values"`
	Blank  bool     `json:"blank"`
}

type SortMutation struct {
	OperationID string     `json:"operation_id"`
	SheetID     string     `json:"sheet_id"`
	Kind        string     `json:"kind"`
	Range       StyleRange `json:"range"`
	KeyColumn   int        `json:"key_column"`
	Descending  bool       `json:"descending"`
	Header      bool       `json:"header"`
}

func applyNativeViewTransaction(original []byte, before *NativeWorkbookV1, mutations []ViewMutation) (*NativeWorkbookMutationResultV1, error) {
	if err := validateViewMutations(mutations); err != nil {
		return nil, err
	}
	last := map[string]ViewMutation{}
	for _, mutation := range mutations {
		last[mutation.SheetID] = mutation
	}
	ids := make([]string, 0, len(last))
	for id := range last {
		ids = append(ids, id)
	}
	return patchExclusiveSheets(original, before, "freeze", ids, func(id string, sheet *NativeWorkbookSheetV1, data []byte) ([]byte, error) {
		return rewriteWorksheetFreeze(data, sheet, last[id])
	})
}

func applyNativeFilterTransaction(original []byte, before *NativeWorkbookV1, mutations []FilterMutation) (*NativeWorkbookMutationResultV1, error) {
	normalized, err := validateFilterMutations(mutations)
	if err != nil {
		return nil, err
	}
	last := map[string]FilterMutation{}
	for _, mutation := range normalized {
		last[mutation.SheetID] = mutation
	}
	ids := make([]string, 0, len(last))
	for id := range last {
		ids = append(ids, id)
	}
	return patchExclusiveSheets(original, before, "filter", ids, func(id string, sheet *NativeWorkbookSheetV1, data []byte) ([]byte, error) {
		return rewriteWorksheetFilter(data, sheet, last[id])
	})
}

func applyNativeSortTransaction(original []byte, before *NativeWorkbookV1, mutations []SortMutation) (*NativeWorkbookMutationResultV1, error) {
	normalized, err := validateSortMutations(mutations)
	if err != nil {
		return nil, err
	}
	bySheet := map[string][]SortMutation{}
	for _, mutation := range normalized {
		bySheet[mutation.SheetID] = append(bySheet[mutation.SheetID], mutation)
	}
	ids := make([]string, 0, len(bySheet))
	for id := range bySheet {
		ids = append(ids, id)
	}
	return patchExclusiveSheets(original, before, "sort", ids, func(id string, sheet *NativeWorkbookSheetV1, data []byte) ([]byte, error) {
		updated := data
		for _, mutation := range bySheet[id] {
			next, err := rewriteWorksheetSort(updated, sheet, mutation)
			if err != nil {
				return nil, err
			}
			updated = next
		}
		return updated, nil
	})
}

func patchExclusiveSheets(original []byte, before *NativeWorkbookV1, label string, ids []string, rewrite func(string, *NativeWorkbookSheetV1, []byte) ([]byte, error)) (*NativeWorkbookMutationResultV1, error) {
	pkg, err := openNativeWorkbookPackage(original)
	if err != nil {
		return nil, err
	}
	sheets := indexNativeMutationSheets(before)
	patch := Patch{Replace: map[string][]byte{}}
	sort.Strings(ids)
	for _, id := range ids {
		indexed, ok := sheets[id]
		if !ok || !indexed.sheet.Editable {
			return nil, fmt.Errorf("xlsxpatch: %s: missing or read-only sheet %q", label, id)
		}
		updated, err := rewrite(id, indexed.sheet, pkg.files[indexed.sheet.PartName])
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: %s: %w", label, err)
		}
		if !bytes.Equal(updated, pkg.files[indexed.sheet.PartName]) {
			patch.Replace[indexed.sheet.PartName] = updated
		}
	}
	if len(patch.Replace) == 0 {
		return nil, fmt.Errorf("xlsxpatch: %s: semantic no-op produced no authoritative package change", label)
	}
	produced, err := Apply(original, patch)
	if err != nil {
		return nil, err
	}
	after, err := ExtractNativeWorkbookV1WithOptions(produced, NativeWorkbookExtractionOptions{Previous: before})
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: %s: reopen/extract result: %w", label, err)
	}
	if issues := ValidateNativeWorkbookV1(after); len(issues) != 0 {
		return nil, fmt.Errorf("xlsxpatch: %s: validate result: %w", label, &NativeWorkbookValidationError{Issues: issues})
	}
	return &NativeWorkbookMutationResultV1{Package: produced, Workbook: after}, nil
}

func requireMutationID(operationID, sheetID string, seen map[string]bool) error {
	if operationID == "" || len(operationID) > maxOperationIDLen || !restrictedIDPattern.MatchString(operationID) || seen[operationID] {
		return fmt.Errorf("xlsxpatch: invalid operation_id")
	}
	seen[operationID] = true
	if sheetID == "" || utf16Length(sheetID) > maxStableSheetIDLen || strings.TrimSpace(sheetID) != sheetID {
		return fmt.Errorf("xlsxpatch: invalid sheet_id")
	}
	return nil
}

func validateViewMutations(mutations []ViewMutation) error {
	if len(mutations) == 0 || len(mutations) > maxCellMutations {
		return fmt.Errorf("xlsxpatch: freeze: empty or oversized batch")
	}
	seen := map[string]bool{}
	for _, mutation := range mutations {
		if err := requireMutationID(mutation.OperationID, mutation.SheetID, seen); err != nil {
			return err
		}
		if mutation.Kind != SheetFreeze || mutation.Rows < 0 || mutation.Rows > excelMaxRows || mutation.Columns < 0 || mutation.Columns > excelMaxColumns {
			return fmt.Errorf("xlsxpatch: freeze: invalid freeze")
		}
	}
	return nil
}

func validateFilterMutations(mutations []FilterMutation) ([]FilterMutation, error) {
	if len(mutations) == 0 || len(mutations) > maxCellMutations {
		return nil, fmt.Errorf("xlsxpatch: filter: empty or oversized batch")
	}
	seen := map[string]bool{}
	for _, mutation := range mutations {
		if err := requireMutationID(mutation.OperationID, mutation.SheetID, seen); err != nil {
			return nil, err
		}
		if mutation.Kind != SheetFilter {
			return nil, fmt.Errorf("xlsxpatch: filter: unsupported kind")
		}
		if mutation.Filter == nil {
			continue
		}
		parsed := parseCanonicalA1Range(mutation.Filter.Ref)
		if parsed == nil || mutation.Filter.Column < parsed.column || mutation.Filter.Column > parsed.endColumn || parsed.endRow <= parsed.row {
			return nil, fmt.Errorf("xlsxpatch: filter: AutoFilter range must include a header row")
		}
		if (len(mutation.Filter.Values) == 0 && !mutation.Filter.Blank) || len(mutation.Filter.Values) > 256 {
			return nil, fmt.Errorf("xlsxpatch: filter: invalid filter values")
		}
		for _, value := range mutation.Filter.Values {
			if value == "" || utf16Length(value) > maxNumberFormatLength || strings.TrimSpace(value) != value {
				return nil, fmt.Errorf("xlsxpatch: filter: invalid filter value")
			}
		}
	}
	return mutations, nil
}

func validateSortMutations(mutations []SortMutation) ([]SortMutation, error) {
	if len(mutations) == 0 || len(mutations) > maxCellMutations {
		return nil, fmt.Errorf("xlsxpatch: sort: empty or oversized batch")
	}
	seen := map[string]bool{}
	for _, mutation := range mutations {
		if err := requireMutationID(mutation.OperationID, mutation.SheetID, seen); err != nil {
			return nil, err
		}
		area, dataRows := mutation.Range, mutation.Range.EndRow-mutation.Range.Row+1
		if mutation.Header {
			dataRows--
		}
		if mutation.Kind != RangeSort || area.Row < 0 || area.EndRow < area.Row || area.EndRow >= excelMaxRows || area.Column < 0 || area.EndColumn < area.Column || area.EndColumn >= excelMaxColumns || mutation.KeyColumn < area.Column || mutation.KeyColumn > area.EndColumn || dataRows < 1 || dataRows > maxSortRows || dataRows*(area.EndColumn-area.Column+1) > maxSortCells {
			return nil, fmt.Errorf("xlsxpatch: sort: invalid sort")
		}
	}
	return mutations, nil
}

func rewriteWorksheetFreeze(data []byte, sheet *NativeWorkbookSheetV1, mutation ViewMutation) ([]byte, error) {
	if sheet.SheetView != nil && sheet.SheetView.PaneState == "split" {
		return nil, fmt.Errorf("existing sheet view is not a qualified freeze target")
	}
	if mutation.Rows != 0 || mutation.Columns != 0 {
		for _, merged := range sheet.MergedRanges {
			if (mutation.Rows > 0 && merged.Row < mutation.Rows && merged.EndRow >= mutation.Rows) || (mutation.Columns > 0 && merged.Column < mutation.Columns && merged.EndColumn >= mutation.Columns) {
				return nil, fmt.Errorf("freeze boundary crosses merged range %s", merged.Ref)
			}
		}
	}
	views, root, err := directChildElements(data, "worksheet", "sheetViews")
	if err != nil {
		return nil, err
	}
	if len(views) > 1 {
		return nil, fmt.Errorf("duplicate sheetViews")
	}
	if mutation.Rows == 0 && mutation.Columns == 0 && len(views) == 0 {
		return append([]byte(nil), data...), nil
	}
	return spliceWorksheetChild(data, root, views, buildSheetViewsXML(root.qname, mutation.Rows, mutation.Columns), []string{"dimension"})
}

func buildSheetViewsXML(rootQName string, rows, columns int) []byte {
	views, view := prefixedLocal(rootQName, "sheetViews"), prefixedLocal(rootQName, "sheetView")
	if rows == 0 && columns == 0 {
		return []byte(`<` + views + `><` + view + ` workbookViewId="0"/></` + views + `>`)
	}
	active, topLeft := "bottomRight", cellReference(rows, columns)
	if rows == 0 {
		active = "topRight"
	} else if columns == 0 {
		active = "bottomLeft"
	}
	pane := `<` + prefixedLocal(rootQName, "pane")
	if columns > 0 {
		pane += ` xSplit="` + strconv.Itoa(columns) + `"`
	}
	if rows > 0 {
		pane += ` ySplit="` + strconv.Itoa(rows) + `"`
	}
	pane += ` topLeftCell="` + topLeft + `" activePane="` + active + `" state="frozen"/>`
	sel := `<` + prefixedLocal(rootQName, "selection") + ` pane="` + active + `" activeCell="` + topLeft + `" sqref="` + topLeft + `"/>`
	return []byte(`<` + views + `><` + view + ` workbookViewId="0">` + pane + sel + `</` + view + `></` + views + `>`)
}

func rewriteWorksheetFilter(data []byte, sheet *NativeWorkbookSheetV1, mutation FilterMutation) ([]byte, error) {
	existing, root, err := directChildElements(data, "worksheet", "autoFilter")
	if err != nil {
		return nil, err
	}
	if len(existing) > 1 {
		return nil, fmt.Errorf("duplicate autoFilter")
	}
	hide, unhide := map[int]bool{}, map[int]bool{}
	var body []byte
	if mutation.Filter == nil {
		if len(existing) == 0 {
			return append([]byte(nil), data...), nil
		}
		parsed := (*struct{ row, column, endRow, endColumn int })(nil)
		if sheet.AutoFilter != nil {
			parsed = parseCanonicalA1Range(sheet.AutoFilter.Ref)
		}
		if parsed != nil {
			for row := parsed.row + 1; row <= parsed.endRow; row++ {
				unhide[row] = true
			}
		}
	} else {
		parsed := parseCanonicalA1Range(mutation.Filter.Ref)
		keep := map[string]bool{}
		for _, value := range mutation.Filter.Values {
			keep[strings.ToLower(value)] = true
		}
		hidden := map[int]bool{}
		for _, row := range sheet.Rows {
			hidden[row.Row] = row.Hidden
		}
		for row := parsed.row + 1; row <= parsed.endRow; row++ {
			if hidden[row] {
				continue
			}
			text, kind := nativeFilterCell(sheet, row, mutation.Filter.Column)
			if kind == "other" || (kind == "empty" && mutation.Filter.Blank) || (kind == "text" && keep[strings.ToLower(text)]) {
				continue
			}
			hide[row] = true
		}
		built, err := buildAutoFilterXML(root.qname, mutation, parsed.column)
		if err != nil {
			return nil, err
		}
		body = built
	}
	updated, err := spliceWorksheetChild(data, root, existing, body, []string{"sheetData"})
	if err != nil {
		return nil, err
	}
	return applyRowHidden(updated, hide, unhide)
}

func nativeFilterCell(sheet *NativeWorkbookSheetV1, row, column int) (string, string) {
	for _, cell := range sheet.Cells {
		if cell.Row != row || cell.Column != column {
			continue
		}
		if cell.Formula != nil {
			return "", "other"
		}
		if cell.Value == nil {
			return "", "empty"
		}
		if cell.Value.Kind == "string" && cell.Value.Text != nil {
			return *cell.Value.Text, "text"
		}
		if cell.Value.Kind == "string" {
			return "", "empty"
		}
		return "", "other"
	}
	return "", "empty"
}

func buildAutoFilterXML(rootQName string, mutation FilterMutation, rangeColumn int) ([]byte, error) {
	qname, col, filters, filter := prefixedLocal(rootQName, "autoFilter"), prefixedLocal(rootQName, "filterColumn"), prefixedLocal(rootQName, "filters"), prefixedLocal(rootQName, "filter")
	var body strings.Builder
	body.WriteString(`<` + qname + ` ref="` + mutation.Filter.Ref + `"><` + col + ` colId="` + strconv.Itoa(mutation.Filter.Column-rangeColumn) + `">`)
	if mutation.Filter.Blank {
		body.WriteString(`<` + filters + ` blank="1">`)
	} else {
		body.WriteString(`<` + filters + `>`)
	}
	for _, value := range mutation.Filter.Values {
		escaped, err := escapeXMLAttribute(value)
		if err != nil {
			return nil, err
		}
		body.WriteString(`<` + filter + ` val="` + escaped + `"/>`)
	}
	body.WriteString(`</` + filters + `></` + col + `></` + qname + `>`)
	return []byte(body.String()), nil
}

func applyRowHidden(data []byte, hide, unhide map[int]bool) ([]byte, error) {
	if len(hide) == 0 && len(unhide) == 0 {
		return data, nil
	}
	layout, err := indexWorksheetLayout(data)
	if err != nil {
		return nil, err
	}
	existing := map[int]layoutRow{}
	for _, row := range layout.rows {
		existing[row.row] = row
	}
	edits := make([]byteEdit, 0, len(hide)+len(unhide))
	one := "1"
	for row := range hide {
		current, ok := existing[row]
		if !ok || rowStartIsHidden(data[current.start:current.startTagEnd]) {
			continue
		}
		updated, err := rewriteUnqualifiedAttribute(data[current.start:current.startTagEnd], "hidden", &one)
		if err != nil {
			return nil, err
		}
		edits = append(edits, byteEdit{current.start, current.startTagEnd, updated})
	}
	for row := range unhide {
		current, ok := existing[row]
		if !ok || !rowStartIsHidden(data[current.start:current.startTagEnd]) {
			continue
		}
		updated, err := rewriteUnqualifiedAttribute(data[current.start:current.startTagEnd], "hidden", nil)
		if err != nil {
			return nil, err
		}
		edits = append(edits, byteEdit{current.start, current.startTagEnd, updated})
	}
	return applyByteEdits(data, edits)
}

func rowStartIsHidden(tag []byte) bool {
	start, err := decodeStartElement(tag)
	if err != nil {
		return false
	}
	raw, found, err := unqualifiedXMLAttribute(start, "hidden")
	if err != nil || !found {
		return false
	}
	hidden, err := ooxmlBoolean(raw, true)
	return err == nil && hidden
}

func rewriteWorksheetSort(data []byte, sheet *NativeWorkbookSheetV1, mutation SortMutation) ([]byte, error) {
	first := mutation.Range.Row
	if mutation.Header {
		first++
	}
	for _, row := range sheet.Rows {
		if row.Row >= first && row.Row <= mutation.Range.EndRow && row.Hidden {
			return nil, fmt.Errorf("sort refuses hidden rows")
		}
	}
	for _, merged := range sheet.MergedRanges {
		if merged.EndRow >= mutation.Range.Row && merged.Row <= mutation.Range.EndRow && merged.EndColumn >= mutation.Range.Column && merged.Column <= mutation.Range.EndColumn {
			return nil, fmt.Errorf("sort refuses merged records")
		}
	}
	for _, cell := range sheet.Cells {
		if cell.Formula != nil && cell.Row >= first && cell.Row <= mutation.Range.EndRow {
			return nil, fmt.Errorf("sort refuses formula records")
		}
	}
	targetRows := make([]int, 0, mutation.Range.EndRow-first+1)
	for row := first; row <= mutation.Range.EndRow; row++ {
		targetRows = append(targetRows, row)
	}
	index, err := indexWorksheet(data, targetRows)
	if err != nil {
		return nil, err
	}
	type sortKey struct {
		empty   bool
		rank    int
		number  float64
		text    string
		boolean bool
	}
	keys := make([]sortKey, len(targetRows))
	for i, row := range targetRows {
		key := sortKey{empty: true}
		for _, cell := range sheet.Cells {
			if cell.Row != row || cell.Column != mutation.KeyColumn || cell.Value == nil {
				continue
			}
			switch cell.Value.Kind {
			case "date", "error":
				return nil, fmt.Errorf("sort refuses date or error sort keys")
			case "boolean":
				key = sortKey{rank: 2, boolean: cell.Value.Lexical != nil && (*cell.Value.Lexical == "1" || strings.EqualFold(*cell.Value.Lexical, "true"))}
			case "string":
				text := ""
				if cell.Value.Text != nil {
					text = *cell.Value.Text
				}
				if text != "" {
					key = sortKey{rank: 1, text: strings.ToLower(text)}
				}
			case "number":
				if cell.Value.Lexical == nil {
					break
				}
				number, err := strconv.ParseFloat(*cell.Value.Lexical, 64)
				if err != nil {
					return nil, fmt.Errorf("sort refuses non-numeric keys")
				}
				key = sortKey{rank: 0, number: number}
			default:
				return nil, fmt.Errorf("sort refuses unreadable keys")
			}
		}
		keys[i] = key
	}
	order := make([]int, len(targetRows))
	for i := range order {
		order[i] = i
	}
	sort.SliceStable(order, func(i, j int) bool {
		left, right := keys[order[i]], keys[order[j]]
		if left.empty || right.empty {
			return !left.empty && right.empty
		}
		if left.rank != right.rank {
			if mutation.Descending {
				return left.rank > right.rank
			}
			return left.rank < right.rank
		}
		cmp := left.number < right.number || left.text < right.text || !left.boolean && right.boolean
		if left.rank == 0 {
			cmp = left.number < right.number
		} else if left.rank == 1 {
			cmp = left.text < right.text
		} else {
			cmp = !left.boolean && right.boolean
		}
		if mutation.Descending {
			return !cmp && (left.number != right.number || left.text != right.text || left.boolean != right.boolean)
		}
		return cmp
	})
	edits := make([]byteEdit, 0, len(targetRows))
	for destIndex, destRow := range targetRows {
		sourceRow := targetRows[order[destIndex]]
		if sourceRow == destRow {
			continue
		}
		dest, destOK := index.existing[destRow]
		source, sourceOK := index.existing[sourceRow]
		if !destOK {
			return nil, fmt.Errorf("sort refuses sparse rows")
		}
		type piece struct {
			column int
			xml    []byte
		}
		combined := make([]piece, 0, len(dest.cells)+len(source.cells))
		for _, cell := range dest.cells {
			if cell.column < mutation.Range.Column || cell.column > mutation.Range.EndColumn {
				combined = append(combined, piece{cell.column, append([]byte(nil), data[cell.start:cell.end]...)})
			}
		}
		if sourceOK {
			for _, cell := range source.cells {
				if cell.column < mutation.Range.Column || cell.column > mutation.Range.EndColumn {
					continue
				}
				updated := append([]byte(nil), data[cell.start:cell.end]...)
				ref := cellReference(destRow, cell.column)
				closeAt := bytes.IndexByte(updated, '>')
				if closeAt < 0 {
					return nil, fmt.Errorf("cell is missing a start tag")
				}
				rewritten, err := rewriteUnqualifiedAttribute(updated[:closeAt+1], "r", &ref)
				if err != nil {
					return nil, err
				}
				combined = append(combined, piece{cell.column, append(rewritten, updated[closeAt+1:]...)})
			}
		}
		sort.Slice(combined, func(i, j int) bool { return combined[i].column < combined[j].column })
		var body bytes.Buffer
		body.Write(data[dest.start:dest.startTagEnd])
		for _, item := range combined {
			body.Write(item.xml)
		}
		if dest.endStart > dest.startTagEnd {
			body.Write(data[dest.endStart:dest.end])
		}
		updated := body.Bytes()
		if !bytes.Equal(updated, data[dest.start:dest.end]) {
			edits = append(edits, byteEdit{dest.start, dest.end, updated})
		}
	}
	return applyByteEdits(data, edits)
}

type byteEdit struct {
	start, end int
	data       []byte
}

func applyByteEdits(data []byte, edits []byteEdit) ([]byte, error) {
	if len(edits) == 0 {
		return append([]byte(nil), data...), nil
	}
	sort.Slice(edits, func(i, j int) bool {
		if edits[i].start == edits[j].start {
			return edits[i].end > edits[j].end
		}
		return edits[i].start > edits[j].start
	})
	updated := append([]byte(nil), data...)
	for _, item := range edits {
		updated = append(updated[:item.start], append(item.data, updated[item.end:]...)...)
	}
	if err := validateWorksheetXML(updated); err != nil {
		return nil, err
	}
	return updated, nil
}

func spliceWorksheetChild(data []byte, root xmlSpan, existing []directChildElement, body []byte, predecessors []string) ([]byte, error) {
	start, end := 0, 0
	if len(existing) == 1 {
		start, end = existing[0].span.start, existing[0].span.end
	} else {
		if len(body) == 0 {
			return append([]byte(nil), data...), nil
		}
		insertAt := root.startTagEnd
		for _, name := range predecessors {
			parts, _, err := directChildElements(data, "worksheet", name)
			if err != nil {
				return nil, err
			}
			for _, part := range parts {
				if part.span.end > insertAt {
					insertAt = part.span.end
				}
			}
		}
		start, end = insertAt, insertAt
	}
	if bytes.Equal(body, data[start:end]) {
		return append([]byte(nil), data...), nil
	}
	out := append(append([]byte(nil), data[:start]...), body...)
	return append(out, data[end:]...), nil
}
