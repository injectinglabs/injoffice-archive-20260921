package xlsxpatch

import (
	"strings"
	"testing"
)

func structureFixture(strict bool) map[string]string {
	entries := nativeMutationFixture(strict)
	delete(entries, "Sheets/_rels/s1.xml.rels")
	delete(entries, "Charts/chart1.xml")
	delete(entries, "Custom/data.bin")
	entries["[Content_Types].xml"] = strings.Replace(entries["[Content_Types].xml"], `<Override PartName="/Charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`, "", 1)
	entries["Meta/Strings.xml"] = strings.Replace(entries["Meta/Strings.xml"], `<r><rPr><b/></rPr><t>Rich </t></r><r><t>Text</t></r>`, `<t>Plain</t>`, 1)
	ns := spreadsheetMLTransitional
	if strict {
		ns = spreadsheetMLStrict
	}
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ns + `"><dimension ref="A1:D4"/><cols><col min="1" max="2" width="17" customWidth="1"/></cols><sheetData><row r="1" ht="30" customHeight="1"><c r="A1" s="1"><v>10</v></c><c r="B1"><f>SUM($A$1:A3)</f><v>30</v></c></row><row r="3"><c r="A3"><v>20</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="C3:D4"/></mergeCells></worksheet>`
	entries["Sheets/s2.xml"] = `<worksheet xmlns="` + ns + `"><sheetData><row r="1"><c r="A1"><f>'Data Set'!$A$1+SUM('Data Set'!A1:A3)+LOG10(100)</f><v>42</v></c></row></sheetData></worksheet>`
	return entries
}
func TestNativeStructureRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, axis := range []string{"row", "column"} {
			original := buildZip(t, structureFixture(strict))
			before, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			apply := func(kind string, index, count int) *NativeWorkbookMutationResultV1 {
				t.Helper()
				result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Structure: []StructureMutation{{OperationID: "axis", SheetID: "7", Kind: kind, Index: index, Count: count}}})
				if err != nil {
					t.Fatal(err)
				}
				original, before = result.Package, result.Workbook
				return result
			}
			result := apply(axis+".insert", 0, 2)
			moved := "A3"
			formula := "B3"
			want := "SUM($A$3:A5)"
			merge := "C5:D6"
			if axis == "column" {
				moved = "C1"
				formula = "D1"
				want = "SUM($C$1:C3)"
				merge = "E3:F4"
			}
			if findNativeCell(t, result.Workbook, "7", moved).StyleID != 1 {
				t.Fatal("style not moved")
			}
			if got := findNativeCell(t, result.Workbook, "7", formula).Formula; got == nil || got.Text != want || got.Cached != nil {
				t.Fatalf("formula: %+v", got)
			}
			if result.Workbook.Sheets[0].MergedRanges[0].Ref != merge {
				t.Fatal("merge not shifted")
			}
			other := findNativeCell(t, result.Workbook, "9", "A1").Formula
			if other == nil || !strings.Contains(other.Text, strings.Split(want, "(")[1][:4]) {
				t.Fatalf("cross-sheet formula not shifted: %+v", other)
			}
			result = apply(axis+".delete", 0, 2)
			if findNativeCell(t, result.Workbook, "7", "B1").Formula.Text != "SUM($A$1:A3)" {
				t.Fatal("inverse deletion did not restore references")
			}
			if result.Workbook.Sheets[0].MergedRanges[0].Ref != "C3:D4" {
				t.Fatal("inverse deletion did not restore merge")
			}
			if result.Workbook.Sheets[0].Rows[0].HeightPoints == nil || *result.Workbook.Sheets[0].Rows[0].HeightPoints != 30 {
				t.Fatal("row height lost")
			}
			if result.Workbook.Sheets[0].Columns[0].Width == nil || *result.Workbook.Sheets[0].Columns[0].Width != 17 {
				t.Fatal("column width lost")
			}
		}
	}
}
func TestStructureFormulaReferences(t *testing.T) {
	cases := []struct {
		input, want, kind string
		index, count      int
	}{
		{`SUM(A1:A5)+$B$2+LOG10(100)+"A1"`, `SUM(A1:A3)+#REF!+LOG10(100)+"A1"`, "row.delete", 1, 2},
		{`'Other'!A1+'Data Set'!$A$2+A:A+1:3`, `'Other'!A1+'Data Set'!$A$4+A:A+1:5`, "row.insert", 1, 2},
		{`SUM($A:$C)+A1+$C$5`, `SUM($A:$B)+A1+$B$5`, "column.delete", 1, 1},
		{`A1+B1+C1`, `A1+#REF!+B1`, "column.delete", 1, 1},
	}
	for _, c := range cases {
		got, err := shiftStructureFormula(c.input, "Data Set", "Data Set", StructureMutation{Kind: c.kind, Index: c.index, Count: c.count})
		if err != nil || got != c.want {
			t.Errorf("%s => %s, %v; want %s", c.input, got, err, c.want)
		}
	}
	for _, formula := range []string{`Table1[Column]`, `Sheet1:Sheet3!A1`, `'[Other.xlsx]Sheet1'!A1`, `A1#`} {
		if _, err := shiftStructureFormula(formula, "Data Set", "Data Set", StructureMutation{Kind: "row.insert", Index: 0, Count: 1}); err == nil {
			t.Errorf("unsafe formula accepted: %s", formula)
		}
	}
}
func TestStructureRefusesBoundaryAndUnmodeledReferences(t *testing.T) {
	entries := structureFixture(false)
	entries["Book/Workbook.xml"] = strings.Replace(entries["Book/Workbook.xml"], "</workbook>", `<definedNames><definedName name="Total">'Data Set'!A1</definedName></definedNames></workbook>`, 1)
	original := buildZip(t, entries)
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Structure: []StructureMutation{{OperationID: "axis", SheetID: "7", Kind: "row.insert", Index: 0, Count: 1}}}); err == nil {
		t.Fatal("named references silently left stale")
	}
	if _, _, _, err := (StructureMutation{Kind: "row.insert", Index: 0, Count: 1}).shiftCell(excelMaxRows-1, 0); err == nil {
		t.Fatal("overflow accepted")
	}
}

func TestStructureDeletionRemovesValuesAndRepairsReferences(t *testing.T) {
	for _, axis := range []string{"row", "column"} {
		original := buildZip(t, structureFixture(false))
		before, err := ExtractNativeWorkbookV1(original)
		if err != nil {
			t.Fatal(err)
		}
		result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Structure: []StructureMutation{{OperationID: "delete", SheetID: "7", Kind: axis + ".delete", Index: 0, Count: 1}}})
		if err != nil {
			t.Fatal(err)
		}
		want := `#REF!+SUM('Data Set'!A1:A2)+LOG10(100)`
		if axis == "column" {
			want = `#REF!+SUM(#REF!)+LOG10(100)`
			if got := findNativeCell(t, result.Workbook, "7", "A1").Formula; got == nil || got.Text != "SUM(#REF!)" {
				t.Fatalf("moved formula: %+v", got)
			}
		} else {
			if got := findNativeCell(t, result.Workbook, "7", "A2").Value; got == nil || got.Lexical == nil || *got.Lexical != "20" {
				t.Fatal("surviving value lost")
			}
		}
		if got := findNativeCell(t, result.Workbook, "9", "A1").Formula; got == nil || got.Text != want || got.Cached != nil {
			t.Fatalf("deletion formula: %+v; want %s", got, want)
		}
	}
}

func TestStructureWorkbookMetadataRefusals(t *testing.T) {
	for _, feature := range []string{`<calcPr ref="A1"/>`, `<calcPr refMode="R1C1"/>`, `<calcPr><extLst/></calcPr>`, `<bookViews><workbookView><extLst/></workbookView></bookViews>`, `<workbookPr><foreign xmlns="urn:test" ref="A1"/></workbookPr>`} {
		data := []byte(`<workbook xmlns="` + spreadsheetMLTransitional + `"><sheets/>` + feature + `</workbook>`)
		if err := validateStructureWorkbook(data); err == nil {
			t.Fatalf("unsafe metadata accepted: %s", feature)
		}
	}
}
