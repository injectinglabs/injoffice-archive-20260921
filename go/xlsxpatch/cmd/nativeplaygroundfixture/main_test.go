package main

import (
	"bytes"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func TestPlaygroundFixtureIsMeaningfulDeterministicAndEditable(t *testing.T) {
	first, err := buildPlaygroundFixture()
	if err != nil {
		t.Fatal(err)
	}
	second, err := buildPlaygroundFixture()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, second) {
		t.Fatal("playground XLSX generator is not deterministic")
	}

	workbook, err := xlsxpatch.ExtractNativeWorkbookV2(first)
	if err != nil {
		t.Fatal(err)
	}
	if len(workbook.Sheets) != 1 || workbook.Sheets[0].Name != "Launch Readiness" {
		t.Fatalf("unexpected workbook sheets: %#v", workbook.Sheets)
	}
	if got := len(workbook.Sheets[0].Cells); got != 48 {
		t.Fatalf("fixture cells = %d, want 48", got)
	}
	if len(workbook.Styles) < 6 {
		t.Fatalf("fixture styles = %d, want at least 6", len(workbook.Styles))
	}

	target := workbook.Sheets[0].Cells[8] // C2, the first workstream status.
	if target.Ref != "C2" || !target.Editable || target.Value == nil || target.Value.Text == nil || *target.Value.Text != "On track" {
		t.Fatalf("expected meaningful editable C2 target, got %#v", target)
	}
	result, err := xlsxpatch.ApplyNativeWorkbookMutationTransactionV1(first, xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: workbook.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: "fixture-status-edit",
			SheetID:     workbook.Sheets[0].ID,
			Kind:        xlsxpatch.CellSetValue,
			Cell:        xlsxpatch.CellRef{Row: 1, Column: 2},
			Value:       "Ready",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Equal(first, result.Package) {
		t.Fatal("guarded fixture mutation returned the source bytes")
	}
	cell := result.Workbook.Sheets[0].Cells[8]
	if cell.Value == nil || cell.Value.Text == nil || *cell.Value.Text != "Ready" {
		t.Fatalf("reopened C2 = %#v, want Ready", cell.Value)
	}
}
