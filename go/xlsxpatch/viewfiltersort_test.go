package xlsxpatch

import (
	"strings"
	"testing"
)

func TestApplyNativeWorkbookMutationTransactionV1_FreezeAndUnfreezeRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			entries := nativeMutationFixture(strict)
			original := buildZip(t, entries)
			before, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			frozen, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
				ExpectedRevision: before.Revision,
				View:             []ViewMutation{{OperationID: "freeze-top", SheetID: "7", Kind: SheetFreeze, Rows: 1, Columns: 0}},
			})
			if err != nil {
				t.Fatal(err)
			}
			sheet := mustSheet(t, frozen.Workbook, "7")
			if sheet.SheetView == nil || sheet.SheetView.PaneState != "frozen" || sheet.SheetView.FrozenRows != 1 || sheet.SheetView.FrozenColumns != 0 {
				t.Fatalf("freeze did not read back: %#v", sheet.SheetView)
			}
			xml := readEntry(t, frozen.Package, "Sheets/s1.xml")
			if !strings.Contains(xml, `ySplit="1"`) || !strings.Contains(xml, `state="frozen"`) {
				t.Fatalf("worksheet is missing frozen pane: %s", xml)
			}
			reopened, err := ExtractNativeWorkbookV1(frozen.Package)
			if err != nil || reopened.Sheets[0].SheetView == nil || reopened.Sheets[0].SheetView.FrozenRows != 1 {
				t.Fatalf("frozen package did not reopen: err=%v view=%#v", err, reopened)
			}
			cleared, err := ApplyNativeWorkbookMutationTransactionV1(frozen.Package, NativeWorkbookMutationTransactionV1{
				ExpectedRevision: frozen.Workbook.Revision,
				View:             []ViewMutation{{OperationID: "unfreeze", SheetID: "7", Kind: SheetFreeze, Rows: 0, Columns: 0}},
			})
			if err != nil {
				t.Fatal(err)
			}
			if view := mustSheet(t, cleared.Workbook, "7").SheetView; view != nil && (view.FrozenRows != 0 || view.FrozenColumns != 0) {
				t.Fatalf("unfreeze left a frozen pane: %#v", view)
			}
		})
	}
}

func TestApplyNativeViewTransaction_RefusesSplitAndMergedBoundary(t *testing.T) {
	entries := nativeMutationFixture(false)
	ssNS := spreadsheetMLTransitional
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:C2"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="2000" ySplit="1500" topLeftCell="C4" activePane="bottomRight"/></sheetView></sheetViews><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>`
	original := buildZip(t, entries)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	_, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		View:             []ViewMutation{{OperationID: "freeze", SheetID: "7", Kind: SheetFreeze, Rows: 1, Columns: 0}},
	})
	if err == nil || !strings.Contains(err.Error(), "split") && !strings.Contains(err.Error(), "qualified freeze") {
		t.Fatalf("split pane should refuse freeze: %v", err)
	}

	entries = nativeMutationFixture(false)
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells></worksheet>`
	original = buildZip(t, entries)
	before, err = ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	_, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		View:             []ViewMutation{{OperationID: "freeze", SheetID: "7", Kind: SheetFreeze, Rows: 1, Columns: 0}},
	})
	if err == nil || !strings.Contains(err.Error(), "merged") {
		t.Fatalf("merged freeze boundary should refuse: %v", err)
	}
}

func TestApplyNativeWorkbookMutationTransactionV1_AutoFilterSetAndClearRoundTrip(t *testing.T) {
	entries := nativeMutationFixture(false)
	ssNS := spreadsheetMLTransitional
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:A4"/><sheetData>` +
		`<row r="1"><c r="A1" t="inlineStr"><is><t>Color</t></is></c></row>` +
		`<row r="2"><c r="A2" t="inlineStr"><is><t>Green</t></is></c></row>` +
		`<row r="3"><c r="A3" t="inlineStr"><is><t>Blue</t></is></c></row>` +
		`<row r="4"><c r="A4" t="inlineStr"><is><t>Green</t></is></c></row>` +
		`</sheetData></worksheet>`
	original := buildZip(t, entries)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	filtered, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Filters: []FilterMutation{{
			OperationID: "filter-green", SheetID: "7", Kind: SheetFilter,
			Filter: &SheetFilterSpec{Ref: "A1:A4", Column: 0, Values: []string{"Green"}, Blank: false},
		}},
	})
	if err != nil {
		t.Fatalf("apply: %v\nbefore rows=%#v", err, mustSheet(t, before, "7").Rows)
	}
	sheet := mustSheet(t, filtered.Workbook, "7")
	if sheet.AutoFilter == nil || sheet.AutoFilter.Ref != "A1:A4" {
		t.Fatalf("AutoFilter did not read back: %#v", sheet.AutoFilter)
	}
	hidden := map[int]bool{}
	for _, row := range sheet.Rows {
		if row.Hidden {
			hidden[row.Row] = true
		}
	}
	if !hidden[2] || hidden[1] || hidden[3] {
		t.Fatalf("filter hid the wrong rows: %#v rows=%#v", hidden, sheet.Rows)
	}
	xml := readEntry(t, filtered.Package, "Sheets/s1.xml")
	if !strings.Contains(xml, `<autoFilter ref="A1:A4">`) || !strings.Contains(xml, `val="Green"`) {
		t.Fatalf("AutoFilter XML missing: %s", xml)
	}
	reopened, err := ExtractNativeWorkbookV1(filtered.Package)
	if err != nil || reopened == nil || mustSheet(t, reopened, "7").AutoFilter == nil {
		t.Fatalf("filtered package did not reopen: %v", err)
	}
	cleared, err := ApplyNativeWorkbookMutationTransactionV1(filtered.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: filtered.Workbook.Revision,
		Filters:          []FilterMutation{{OperationID: "clear-filter", SheetID: "7", Kind: SheetFilter, Filter: nil}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if mustSheet(t, cleared.Workbook, "7").AutoFilter != nil {
		t.Fatalf("clear left AutoFilter: %#v", mustSheet(t, cleared.Workbook, "7").AutoFilter)
	}
	for _, row := range mustSheet(t, cleared.Workbook, "7").Rows {
		if row.Hidden {
			t.Fatalf("clear left hidden row %#v", row)
		}
	}
}

func TestValidateFilterMutations_RefusesSingleRowRange(t *testing.T) {
	_, err := validateFilterMutations([]FilterMutation{{
		OperationID: "bad", SheetID: "7", Kind: SheetFilter,
		Filter: &SheetFilterSpec{Ref: "A1:B1", Column: 0, Values: []string{"Green"}, Blank: false},
	}})
	if err == nil || !strings.Contains(err.Error(), "header") {
		t.Fatalf("single-row AutoFilter should refuse: %v", err)
	}
}

func TestApplyNativeWorkbookMutationTransactionV1_SortAscendingDescendingRoundTrip(t *testing.T) {
	entries := nativeMutationFixture(false)
	ssNS := spreadsheetMLTransitional
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:B4"/><sheetData>` +
		`<row r="1"><c r="A1" t="inlineStr"><is><t>N</t></is></c><c r="B1" t="inlineStr"><is><t>Keep</t></is></c></row>` +
		`<row r="2"><c r="A2"><v>3</v></c><c r="B2" t="inlineStr"><is><t>c</t></is></c></row>` +
		`<row r="3"><c r="A3"><v>1</v></c><c r="B3" t="inlineStr"><is><t>a</t></is></c></row>` +
		`<row r="4"><c r="A4"><v>2</v></c><c r="B4" t="inlineStr"><is><t>b</t></is></c></row>` +
		`</sheetData></worksheet>`
	original := buildZip(t, entries)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	sorted, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Sorts: []SortMutation{{
			OperationID: "sort-asc", SheetID: "7", Kind: RangeSort,
			Range: StyleRange{Row: 0, Column: 0, EndRow: 3, EndColumn: 1}, KeyColumn: 0, Descending: false, Header: true,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	got := columnValues(t, sorted.Workbook, "7", 0)
	if strings.Join(got, ",") != "N,1,2,3" {
		t.Fatalf("ascending sort order: %v", got)
	}
	labels := columnValues(t, sorted.Workbook, "7", 1)
	if strings.Join(labels, ",") != "Keep,a,b,c" {
		t.Fatalf("record movement lost companion column: %v", labels)
	}
	reopened, err := ExtractNativeWorkbookV1(sorted.Package)
	if err != nil {
		t.Fatal(err)
	}
	if strings.Join(columnValues(t, reopened, "7", 0), ",") != "N,1,2,3" {
		t.Fatalf("sorted package did not reopen in order")
	}
	desc, err := ApplyNativeWorkbookMutationTransactionV1(sorted.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: sorted.Workbook.Revision,
		Sorts: []SortMutation{{
			OperationID: "sort-desc", SheetID: "7", Kind: RangeSort,
			Range: StyleRange{Row: 0, Column: 0, EndRow: 3, EndColumn: 1}, KeyColumn: 0, Descending: true, Header: true,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if strings.Join(columnValues(t, desc.Workbook, "7", 0), ",") != "N,3,2,1" {
		t.Fatalf("descending sort order: %v", columnValues(t, desc.Workbook, "7", 0))
	}
}

func TestApplyNativeSortTransaction_RefusesMergedAndFormulaKeys(t *testing.T) {
	entries := nativeMutationFixture(false)
	ssNS := spreadsheetMLTransitional
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:A2"/><sheetData>` +
		`<row r="1"><c r="A1"><v>2</v></c></row><row r="2"/>` +
		`</sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells></worksheet>`
	original := buildZip(t, entries)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	_, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Sorts: []SortMutation{{
			OperationID: "merged", SheetID: "7", Kind: RangeSort,
			Range: StyleRange{Row: 0, Column: 0, EndRow: 1, EndColumn: 0}, KeyColumn: 0,
		}},
	})
	if err == nil || !strings.Contains(err.Error(), "merged") {
		t.Fatalf("merged sort should refuse: %v", err)
	}

	entries = nativeMutationFixture(false)
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:A2"/><sheetData>` +
		`<row r="1"><c r="A1"><f>SUM(1)</f><v>1</v></c></row><row r="2"><c r="A2"><v>2</v></c></row>` +
		`</sheetData></worksheet>`
	original = buildZip(t, entries)
	before, err = ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	_, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Sorts: []SortMutation{{
			OperationID: "formula", SheetID: "7", Kind: RangeSort,
			Range: StyleRange{Row: 0, Column: 0, EndRow: 1, EndColumn: 0}, KeyColumn: 0,
		}},
	})
	if err == nil || !strings.Contains(err.Error(), "formula") {
		t.Fatalf("formula sort should refuse: %v", err)
	}
}

func mustSheet(t *testing.T, workbook *NativeWorkbookV1, id string) NativeWorkbookSheetV1 {
	t.Helper()
	for _, sheet := range workbook.Sheets {
		if sheet.ID == id {
			return sheet
		}
	}
	t.Fatalf("sheet %s not found", id)
	return NativeWorkbookSheetV1{}
}

func columnValues(t *testing.T, workbook *NativeWorkbookV1, id string, column int) []string {
	t.Helper()
	sheet := mustSheet(t, workbook, id)
	max := -1
	byRow := map[int]string{}
	for _, cell := range sheet.Cells {
		if cell.Column != column || cell.Value == nil {
			continue
		}
		if cell.Row > max {
			max = cell.Row
		}
		switch {
		case cell.Value.Kind == "string" && cell.Value.Text != nil:
			byRow[cell.Row] = *cell.Value.Text
		case cell.Value.Lexical != nil:
			byRow[cell.Row] = *cell.Value.Lexical
		}
	}
	out := make([]string, 0, max+1)
	for row := 0; row <= max; row++ {
		out = append(out, byRow[row])
	}
	return out
}
