package xlsxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func TestNativeChartDataMutationExactRecalculationInventory(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeMutationFixture(strict)
		parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `<calcPr calcId="191029"/>`, "", 1)
		original := buildZip(t, parts)
		before, err := ExtractNativeWorkbookV1(original)
		if err != nil {
			t.Fatal(err)
		}
		if hasNativeWorkbookUnsupported(before, "UNMODELED_WORKBOOK_FEATURE") {
			t.Fatal("source must have no generic workbook record")
		}
		transaction := NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Cells: []CellMutation{{OperationID: "metadata-edit", SheetID: "7", Kind: CellSetValue, Cell: CellRef{Row: 0, Column: 1}, Value: "updated"}}}
		result, err := ApplyNativeWorkbookMutationTransactionV1(original, transaction)
		if err != nil {
			t.Fatal(err)
		}
		expected, err := nativeExpectedUnsupportedAfterCells(original, result.Package, before, transaction)
		if err != nil {
			t.Fatal(err)
		}
		if len(expected) != len(before.Unsupported)+1 {
			t.Fatal("expected exactly one new generic record")
		}
		if err := verifyNativeUnsupportedInventoryPreserved(expected, result.Workbook.Unsupported); err != nil {
			t.Fatal(err)
		}
		if err := verifyNativeUnsupportedInventoryPreserved(before.Unsupported, result.Workbook.Unsupported); err == nil {
			t.Fatal("strict default comparator widened")
		}
		for _, tx := range []NativeWorkbookMutationTransactionV1{{}, {Styles: []StylePatchMutation{{}}}, {Layout: []LayoutMutation{{}}}} {
			unchanged, err := nativeExpectedUnsupportedAfterCells(original, result.Package, before, tx)
			if err != nil {
				t.Fatal(err)
			}
			if err := verifyNativeUnsupportedInventoryPreserved(unchanged, result.Workbook.Unsupported); err == nil {
				t.Fatal("non-cell transaction admitted generated record")
			}
		}
		for _, edit := range []func(map[string]string){
			func(p map[string]string) {
				p["Book/Workbook.xml"] = strings.Replace(p["Book/Workbook.xml"], "</workbook>", `<foreign xmlns="urn:test"/></workbook>`, 1)
			},
			func(p map[string]string) {
				p["Book/Workbook.xml"] = strings.Replace(p["Book/Workbook.xml"], "</workbook>", `<unknown/></workbook>`, 1)
			},
			func(p map[string]string) {
				p["Book/Workbook.xml"] = strings.Replace(p["Book/Workbook.xml"], `name="Hidden"`, `name="Changed"`, 1)
			},
		} {
			modified := map[string]string{}
			for name, data := range independentOPCEntries(t, result.Package) {
				modified[name] = string(data)
			}
			edit(modified)
			if _, err := nativeExpectedUnsupportedAfterCells(original, buildZip(t, modified), before, transaction); err == nil {
				t.Fatal("unrelated workbook edit admitted")
			}
		}
		for _, change := range []func([]NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1{
			func(items []NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1 {
				items[len(items)-1].ID = "spoof"
				return items
			},
			func(items []NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1 {
				items[len(items)-1].PartName = nativeWorkbookString("wrong.xml")
				return items
			},
			func(items []NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1 {
				items[len(items)-1].Message = "spoof"
				return items
			},
			func(items []NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1 {
				items[0].Message = "changed prior record"
				return items
			},
			func(items []NativeWorkbookUnsupportedV1) []NativeWorkbookUnsupportedV1 {
				return append(items, items[len(items)-1])
			},
		} {
			encoded, _ := json.Marshal(expected)
			var changed []NativeWorkbookUnsupportedV1
			if err := json.Unmarshal(encoded, &changed); err != nil {
				t.Fatal(err)
			}
			if err := verifyNativeUnsupportedInventoryPreserved(expected, change(changed)); err == nil {
				t.Fatal("spoofed inventory admitted")
			}
		}
	}
}
