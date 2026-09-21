package xlsxpatch

import (
	"bytes"
	"fmt"
	"reflect"
	"sort"
	"strings"
)

// MergeMutation changes merge metadata and clears covered empty strings.
// Nonempty values and formulas are never discarded.
// Merge batches are isolated from cell/style/layout batches to retain wire order.
type MergeMutation struct {
	OperationID string     `json:"operation_id"`
	SheetID     string     `json:"sheet_id"`
	Kind        string     `json:"kind"`
	Range       StyleRange `json:"range"`
}

func validateMergeMutations(merges []MergeMutation) error {
	if len(merges) > 100 {
		return fmt.Errorf("xlsxpatch: merge: at most 100 operations per transaction")
	}
	seen := map[string]bool{}
	for _, m := range merges {
		if !restrictedIDPattern.MatchString(m.OperationID) || len(m.OperationID) > maxOperationIDLen || seen[m.OperationID] {
			return fmt.Errorf("xlsxpatch: merge: invalid or duplicate operation_id")
		}
		seen[m.OperationID] = true
		if m.SheetID == "" || strings.TrimSpace(m.SheetID) != m.SheetID || utf16Length(m.SheetID) > maxStableSheetIDLen {
			return fmt.Errorf("xlsxpatch: merge: invalid sheet_id")
		}
		if m.Kind != "range.merge" && m.Kind != "range.unmerge" {
			return fmt.Errorf("xlsxpatch: merge: unsupported kind %q", m.Kind)
		}
		r := m.Range
		if r.Row < 0 || r.Column < 0 || r.EndRow < r.Row || r.EndColumn < r.Column || r.EndRow >= excelMaxRows || r.EndColumn >= excelMaxColumns || (r.Row == r.EndRow && r.Column == r.EndColumn) {
			return fmt.Errorf("xlsxpatch: merge: invalid range")
		}
	}
	return nil
}

func applyNativeMergeTransaction(original []byte, before *NativeWorkbookV1, mutations []MergeMutation) (*NativeWorkbookMutationResultV1, error) {
	pkg, err := openNativeWorkbookPackage(original)
	if err != nil {
		return nil, err
	}
	expected := map[string][]NativeWorkbookMergedRangeV1{}
	emptyCells := map[string]CellMutation{}
	sheets := indexNativeMutationSheets(before)
	for _, m := range mutations {
		indexed, ok := sheets[m.SheetID]
		if !ok || !indexed.sheet.Editable {
			return nil, fmt.Errorf("xlsxpatch: merge: missing or read-only sheet %q", m.SheetID)
		}
		// Tables and unmodeled worksheet regions may forbid merging even empty cells.
		for _, u := range before.Unsupported {
			if u.ScopeID == "sheet:"+m.SheetID && (u.Code == "TABLE_REFERENCE" || u.Code == "UNMODELED_WORKSHEET_FEATURE" || u.Code == "WORKSHEET_EXTENSIONS") {
				return nil, fmt.Errorf("xlsxpatch: merge: sheet has unsupported %s", u.Code)
			}
		}
		ranges, exists := expected[m.SheetID]
		if !exists {
			ranges = append([]NativeWorkbookMergedRangeV1{}, indexed.sheet.MergedRanges...)
		}
		r := m.Range
		next := make([]NativeWorkbookMergedRangeV1, 0, len(ranges)+1)
		removed := false
		for _, v := range ranges {
			overlaps := r.Row <= v.EndRow && r.EndRow >= v.Row && r.Column <= v.EndColumn && r.EndColumn >= v.Column
			if !overlaps {
				next = append(next, v)
				continue
			}
			if m.Kind == "range.merge" {
				return nil, fmt.Errorf("xlsxpatch: merge: range overlaps existing merge %s; unmerge first", v.Ref)
			}
			if r.Row > v.Row || r.Column > v.Column || r.EndRow < v.EndRow || r.EndColumn < v.EndColumn {
				return nil, fmt.Errorf("xlsxpatch: unmerge: select the complete merged range %s", v.Ref)
			}
			removed = true
		}
		if m.Kind == "range.merge" {
			next = append(next, NativeWorkbookMergedRangeV1{Ref: cellReference(r.Row, r.Column) + ":" + cellReference(r.EndRow, r.EndColumn), Row: r.Row, Column: r.Column, EndRow: r.EndRow, EndColumn: r.EndColumn, Editable: false})
			cells := append([]NativeWorkbookCellV1{}, indexed.sheet.Cells...)
			for i, c := range cells {
				if c.Row >= r.Row && c.Row <= r.EndRow && c.Column >= r.Column && c.Column <= r.EndColumn && (c.Row != r.Row || c.Column != r.Column) && c.Formula == nil && c.Value != nil && c.Value.Kind == "string" && c.Value.Text != nil && *c.Value.Text == "" {
					if !c.Editable {
						return nil, fmt.Errorf("xlsxpatch: merge: covered empty cell is read-only")
					}
					key := m.SheetID + ":" + c.Ref
					if _, exists := emptyCells[key]; !exists {
						emptyCells[key] = CellMutation{OperationID: fmt.Sprintf("merge-empty-%d", len(emptyCells)), SheetID: m.SheetID, Kind: CellClearValue, Cell: CellRef{Row: c.Row, Column: c.Column}}
					}
					cells[i].Value = nil
				}
			}

			if err := validateAndMarkCellsInNativeMergedRanges(cells, next); err != nil {
				return nil, fmt.Errorf("xlsxpatch: merge: %w", err)
			}
		} else if !removed {
			return nil, fmt.Errorf("xlsxpatch: unmerge: selection contains no merged ranges")
		}
		if len(next) > NativeXLSXMaxMergedRanges {
			return nil, fmt.Errorf("xlsxpatch: merge: too many merged ranges")
		}
		sort.Slice(next, func(i, j int) bool {
			if next[i].Row != next[j].Row {
				return next[i].Row < next[j].Row
			}
			return next[i].Column < next[j].Column
		})
		expected[m.SheetID] = next
	}
	if len(emptyCells) > 0 {
		clears := make([]CellMutation, 0, len(emptyCells))
		keys := make([]string, 0, len(emptyCells))
		for key := range emptyCells {
			keys = append(keys, key)
		}
		sort.Strings(keys)
		for _, key := range keys {
			clears = append(clears, emptyCells[key])
		}
		original, err = ApplyCellMutations(original, clears)
		if err != nil {
			return nil, err
		}
		pkg, err = openNativeWorkbookPackage(original)
		if err != nil {
			return nil, err
		}
	}
	patch := Patch{Replace: map[string][]byte{}}
	for id, ranges := range expected {
		part := sheets[id].sheet.PartName
		updated, err := rewriteWorksheetMerges(pkg.files[part], ranges)
		if err != nil {
			return nil, err
		}
		patch.Replace[part] = updated
	}
	produced, err := Apply(original, patch)
	if err != nil {
		return nil, err
	}
	after, err := ExtractNativeWorkbookV1WithOptions(produced, NativeWorkbookExtractionOptions{Previous: before})
	if err != nil {
		return nil, err
	}
	if issues := ValidateNativeWorkbookV1(after); len(issues) > 0 {
		return nil, &NativeWorkbookValidationError{Issues: issues}
	}
	// After normalizing covered empty strings, the writer splices only the mergeCells span; Apply verifies raw preservation
	// of every other ZIP member. Re-extraction validates the complete package.
	for _, sheet := range after.Sheets {
		want, changed := expected[sheet.ID]
		if !changed {
			want = sheets[sheet.ID].sheet.MergedRanges
		}
		if !reflect.DeepEqual(want, sheet.MergedRanges) {
			return nil, fmt.Errorf("xlsxpatch: merge: readback mismatch")
		}
	}
	return &NativeWorkbookMutationResultV1{Package: produced, Workbook: after}, nil
}

func rewriteWorksheetMerges(data []byte, ranges []NativeWorkbookMergedRangeV1) ([]byte, error) {
	entries, root, err := directChildElements(data, "worksheet", "mergeCells")
	if err != nil {
		return nil, err
	}
	if len(entries) > 1 {
		return nil, fmt.Errorf("xlsxpatch: merge: duplicate mergeCells")
	}
	qname := prefixedLocal(root.qname, "mergeCells")
	var out strings.Builder
	if len(ranges) > 0 {
		fmt.Fprintf(&out, `<%s count="%d">`, qname, len(ranges))
		for _, r := range ranges {
			fmt.Fprintf(&out, `<%s ref="%s"/>`, prefixedLocal(root.qname, "mergeCell"), r.Ref)
		}
		out.WriteString("</" + qname + ">")
	}
	start, end := 0, 0
	if len(entries) == 1 {
		start, end = entries[0].span.start, entries[0].span.end
	} else {
		// All legal predecessors are considered so a saved autoFilter/sortState
		// remains ahead of mergeCells in SpreadsheetML schema order.
		for _, name := range []string{"sheetData", "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter", "sortState", "dataConsolidate", "customSheetViews"} {
			parts, _, err := directChildElements(data, "worksheet", name)
			if err != nil {
				return nil, err
			}
			for _, p := range parts {
				if p.span.end > start {
					start = p.span.end
				}
			}
		}
		end = start
		if start == 0 {
			return nil, fmt.Errorf("xlsxpatch: merge: missing sheetData")
		}
	}
	result := append(bytes.Clone(data[:start]), []byte(out.String())...)
	result = append(result, data[end:]...)
	return result, nil
}
