package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"strings"
	"testing"
)

func nativeChartSourceWorkbook(t *testing.T) []byte {
	t.Helper()
	raw, _, err := BuildNativeGetCorpusFile("pass-agent-dejavu")
	if err != nil {
		t.Fatal(err)
	}
	return raw
}

func nativeChartInsert(sheetID, typ, title string) ChartMutation {
	return ChartMutation{
		OperationID: "insert-" + typ,
		SheetID:     sheetID,
		Kind:        chartInsert,
		ChartType:   typ,
		Title:       title,
		Range:       StyleRange{Row: 0, Column: 0, EndRow: 4, EndColumn: 2},
		Anchor:      NativeChartAnchorV1{FromRow: 0, FromColumn: 4, ToRow: 18, ToColumn: 12},
	}
}

func TestNativeChartInsertUpdateDeleteRoundTrip(t *testing.T) {
	original := nativeChartSourceWorkbook(t)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	sheetID := before.Sheets[0].ID
	inserted, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Charts:           []ChartMutation{nativeChartInsert(sheetID, "bar", "Revenue")},
	})
	if err != nil {
		t.Fatal(err)
	}
	if issues := ValidateNativeWorkbookV1(inserted.Workbook); len(issues) != 0 {
		t.Fatalf("inserted workbook is invalid: %v", issues)
	}
	reopened, err := ExtractNativeWorkbookV1WithOptions(inserted.Package, NativeWorkbookExtractionOptions{Previous: inserted.Workbook})
	if err != nil || reopened.Revision != inserted.Workbook.Revision {
		t.Fatalf("inserted package did not reopen: %v", err)
	}
	charts, err := ReadCharts(inserted.Package)
	if err != nil || len(charts) != 1 || charts[0].Type != "bar" || charts[0].Title != "Revenue" || charts[0].Identity == nil {
		t.Fatalf("inserted chart: %+v err=%v", charts, err)
	}
	if charts[0].Series[0].CategoriesRef != "Data!$A$2:$A$5" || charts[0].Series[0].ValuesRef != "Data!$B$2:$B$5" || charts[0].Series[1].ValuesRef != "Data!$C$2:$C$5" {
		t.Fatalf("series refs: %+v", charts[0].Series)
	}
	sheet := readEntry(t, inserted.Package, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, "<drawing ") {
		t.Fatalf("worksheet missing drawing: %s", sheet)
	}
	if !strings.Contains(readEntry(t, inserted.Package, "xl/worksheets/_rels/sheet1.xml.rels"), "../drawings/drawing1.xml") {
		t.Fatal("sheet rels missing drawing")
	}
	if !strings.Contains(readEntry(t, inserted.Package, "xl/drawings/_rels/drawing1.xml.rels"), "../charts/chart1.xml") {
		t.Fatal("drawing rels missing chart")
	}
	ct := readEntry(t, inserted.Package, "[Content_Types].xml")
	if !strings.Contains(ct, `PartName="/xl/charts/chart1.xml"`) || !strings.Contains(ct, `PartName="/xl/drawings/drawing1.xml"`) {
		t.Fatalf("content types missing chart/drawing: %s", ct)
	}
	editable := projectNativeEditableChartsV1(inserted.Package)
	if len(editable) != 1 || !editable[0].Editable || editable[0].ChartType != "bar" || editable[0].SheetID != sheetID {
		t.Fatalf("editable projection: %+v", editable)
	}
	if len(editable[0].Categories) != 4 || editable[0].Categories[0] != "Q1" || len(editable[0].Series) != 2 || editable[0].Series[0].Values[0] != "10" {
		t.Fatalf("literal series: %+v", editable[0])
	}

	updated, err := ApplyNativeWorkbookMutationTransactionV1(inserted.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: inserted.Workbook.Revision,
		Charts: []ChartMutation{{
			OperationID:               "to-line",
			SheetID:                   sheetID,
			Kind:                      chartUpdate,
			ChartType:                 "line",
			Title:                     "Revenue",
			Range:                     StyleRange{Row: 0, Column: 0, EndRow: 4, EndColumn: 2},
			Anchor:                    NativeChartAnchorV1{FromRow: 0, FromColumn: 4, ToRow: 18, ToColumn: 12},
			Identity:                  charts[0].Identity,
			ExpectedFingerprintSHA256: editable[0].FingerprintSHA256,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	line, err := ReadCharts(updated.Package)
	if err != nil || len(line) != 1 || line[0].Type != "line" || line[0].Identity == nil || *line[0].Identity != *charts[0].Identity {
		t.Fatalf("updated chart lost identity: %+v err=%v", line, err)
	}

	pieSource := projectNativeEditableChartsV1(updated.Package)
	pie, err := ApplyNativeWorkbookMutationTransactionV1(updated.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: updated.Workbook.Revision,
		Charts: []ChartMutation{{
			OperationID:               "to-pie",
			SheetID:                   sheetID,
			Kind:                      chartUpdate,
			ChartType:                 "pie",
			Title:                     "Share",
			Range:                     StyleRange{Row: 0, Column: 0, EndRow: 4, EndColumn: 1},
			Anchor:                    NativeChartAnchorV1{FromRow: 0, FromColumn: 4, ToRow: 18, ToColumn: 12},
			Identity:                  line[0].Identity,
			ExpectedFingerprintSHA256: pieSource[0].FingerprintSHA256,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	pieCharts, err := ReadCharts(pie.Package)
	if err != nil || len(pieCharts) != 1 || pieCharts[0].Type != "pie" || pieCharts[0].Title != "Share" {
		t.Fatalf("pie chart: %+v err=%v", pieCharts, err)
	}

	deleted, err := ApplyNativeWorkbookMutationTransactionV1(pie.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: pie.Workbook.Revision,
		Charts: []ChartMutation{{
			OperationID:               "delete-1",
			SheetID:                   sheetID,
			Kind:                      chartDelete,
			Identity:                  pieCharts[0].Identity,
			ExpectedFingerprintSHA256: projectNativeEditableChartsV1(pie.Package)[0].FingerprintSHA256,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	gone, err := ReadCharts(deleted.Package)
	if err != nil || len(gone) != 0 {
		t.Fatalf("deleted charts: %+v err=%v", gone, err)
	}
	zr, err := zip.NewReader(bytes.NewReader(deleted.Package), int64(len(deleted.Package)))
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range zr.File {
		if file.Name == "xl/charts/chart1.xml" {
			t.Fatal("chart part survived delete")
		}
	}
}

func TestNativeChartInsertRefusesFormulasAndStaleFingerprints(t *testing.T) {
	original := nativeChartSourceWorkbook(t)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	sheetID := before.Sheets[0].ID
	formula, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells: []CellMutation{{
			OperationID: "set-formula", SheetID: sheetID, Kind: CellSetFormula,
			Cell: CellRef{Row: 1, Column: 1}, Formula: "=B3",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ApplyNativeWorkbookMutationTransactionV1(formula.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: formula.Workbook.Revision,
		Charts:           []ChartMutation{nativeChartInsert(sheetID, "column", "")},
	}); err == nil || !strings.Contains(err.Error(), "literal") {
		t.Fatalf("formula source must be refused: %v", err)
	}

	inserted, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Charts:           []ChartMutation{nativeChartInsert(sheetID, "column", "A")},
	})
	if err != nil {
		t.Fatal(err)
	}
	charts, _ := ReadCharts(inserted.Package)
	if _, err := ApplyNativeWorkbookMutationTransactionV1(inserted.Package, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: inserted.Workbook.Revision,
		Charts: []ChartMutation{{
			OperationID:               "stale",
			SheetID:                   sheetID,
			Kind:                      chartDelete,
			Identity:                  charts[0].Identity,
			ExpectedFingerprintSHA256: "sha256:" + strings.Repeat("0", 64),
		}},
	}); err == nil || !strings.Contains(err.Error(), "stale chart fingerprint") {
		t.Fatalf("stale fingerprint must be refused: %v", err)
	}
}

func TestNativeChartInspectObjectsIncludesEditableCharts(t *testing.T) {
	original := nativeChartSourceWorkbook(t)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	inserted, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Charts:           []ChartMutation{nativeChartInsert(before.Sheets[0].ID, "bar", "Inspect")},
	})
	if err != nil {
		t.Fatal(err)
	}
	objects, err := InspectNativeWorkbookObjectsV1(inserted.Package)
	if err != nil {
		t.Fatal(err)
	}
	if len(objects.EditableCharts) != 1 || !objects.EditableCharts[0].Editable || objects.EditableCharts[0].Title != "Inspect" {
		t.Fatalf("inspect editable charts: %+v", objects.EditableCharts)
	}
}
