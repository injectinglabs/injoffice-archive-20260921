package xlsxpatch

import (
	"bytes"
	"testing"
)

func TestNativeMergeUnmergeRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := buildZip(t, nativeMutationFixture(strict))
		before, err := ExtractNativeWorkbookV1(original)
		if err != nil {
			t.Fatal(err)
		}
		m := MergeMutation{OperationID: "merge", SheetID: "7", Kind: "range.merge", Range: StyleRange{Row: 0, Column: 1, EndRow: 2, EndColumn: 1}}
		merged, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Merges: []MergeMutation{m}})
		if err != nil {
			t.Fatal(err)
		}
		if len(merged.Workbook.Sheets[0].MergedRanges) != 1 || merged.Workbook.Sheets[0].MergedRanges[0].Ref != "B1:B3" {
			t.Fatal("merge not extracted")
		}
		for _, c := range merged.Workbook.Sheets[0].Cells {
			if c.Ref == "B1" && c.Editable {
				t.Fatal("merged cell unexpectedly editable")
			}
		}
		m.Kind = "range.unmerge"
		unmerged, err := ApplyNativeWorkbookMutationTransactionV1(merged.Package, NativeWorkbookMutationTransactionV1{ExpectedRevision: merged.Workbook.Revision, Merges: []MergeMutation{m}})
		if err != nil {
			t.Fatal(err)
		}
		if len(unmerged.Workbook.Sheets[0].MergedRanges) != 0 {
			t.Fatal("merge survived unmerge")
		}
		if got := readEntry(t, unmerged.Package, "Sheets/s1.xml"); got != readEntry(t, original, "Sheets/s1.xml") {
			t.Fatal("unmerge did not restore exact worksheet XML")
		}
		if got := readEntry(t, unmerged.Package, "Meta/Styles.style"); got != readEntry(t, original, "Meta/Styles.style") {
			t.Fatal("styles changed")
		}
		if bytes.Equal(original, merged.Package) {
			t.Fatal("merge produced no bytes")
		}
	}
}

func TestNativeMergeRefusesLossOverlapPartialAndMixed(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	m := MergeMutation{OperationID: "merge", SheetID: "7", Kind: "range.merge", Range: StyleRange{Row: 0, Column: 1, EndRow: 2, EndColumn: 2}}
	if result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Merges: []MergeMutation{m}}); err == nil || result != nil {
		t.Fatal("merge discarded C1")
	}
	m.Range.EndColumn = 1
	merged, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Merges: []MergeMutation{m}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ApplyNativeWorkbookMutationTransactionV1(merged.Package, NativeWorkbookMutationTransactionV1{ExpectedRevision: merged.Workbook.Revision, Merges: []MergeMutation{m}}); err == nil {
		t.Fatal("overlap accepted")
	}
	m.Kind = "range.unmerge"
	m.Range.EndRow = 1
	if _, err := ApplyNativeWorkbookMutationTransactionV1(merged.Package, NativeWorkbookMutationTransactionV1{ExpectedRevision: merged.Workbook.Revision, Merges: []MergeMutation{m}}); err == nil {
		t.Fatal("partial unmerge accepted")
	}
	if _, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Merges: []MergeMutation{m}, Layout: []LayoutMutation{{OperationID: "row", SheetID: "7", Kind: RowSetHeight, HeightPoints: 25}}}); err == nil {
		t.Fatal("mixed batch accepted")
	}
}

func TestNativeMergeCoveredEmptyStrings(t *testing.T) {
	entries := nativeMutationFixture(false)
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Keep</t></is></c><c r="B1" t="inlineStr"><is><t></t></is></c></row></sheetData></worksheet>`
	original := buildZip(t, entries)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Merges: []MergeMutation{{OperationID: "empty", SheetID: "7", Kind: "range.merge", Range: StyleRange{Row: 0, Column: 0, EndRow: 0, EndColumn: 1}}}})
	if err != nil {
		t.Fatal(err)
	}
	for _, c := range result.Workbook.Sheets[0].Cells {
		if c.Ref == "A1" && (c.Value == nil || c.Value.Text == nil || *c.Value.Text != "Keep") {
			t.Fatal("anchor value lost")
		}
		if c.Ref == "B1" && c.Value != nil {
			t.Fatal("empty string still covers merge")
		}
	}
}
