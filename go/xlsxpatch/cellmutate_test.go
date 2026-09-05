package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"math"
	"strings"
	"testing"
)

func hasZipEntry(t *testing.T, data []byte, name string) bool {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range zr.File {
		if file.Name == name {
			return true
		}
	}
	return false
}

func cellMutationFixture() map[string]string {
	return map[string]string{
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml": `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>` +
			`<sheet name="Data" sheetId="7" r:id="rIdData"/><sheet name="Other" sheetId="9" r:id="rIdOther"/>` +
			`</sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>` +
			`<Relationship Id="rIdOther" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>` +
			`</Relationships>`,
		"xl/worksheets/sheet1.xml": `<?xml version="1.0"?><x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:dimension ref="A1:D3"/><x:sheetData>` +
			`<x:row r="1" spans="1:4"><x:c r="A1" s="7" t="s"><x:v>0</x:v><x:extLst><x:ext uri="keep"><x:payload/></x:ext></x:extLst></x:c><x:c r="D1" s="3"><x:f>OLD()</x:f><x:v>99</x:v></x:c></x:row>` +
			`<x:row r="3"><x:c r="B3" s="2" t="inlineStr"><x:is><x:t>old</x:t></x:is></x:c></x:row>` +
			`</x:sheetData><x:extLst><x:ext uri="worksheet-keep"/></x:extLst></x:worksheet>`,
		"xl/worksheets/sheet2.xml": `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`,
		"xl/charts/chart1.xml":     `<chartSpace><native-content>keep exactly</native-content></chartSpace>`,
		"customXml/item1.xml":      `<root attr="preserve">opaque</root>`,
	}
}

func mutation(id, sheet string, kind CellMutationKind, row, column int) CellMutation {
	return CellMutation{OperationID: id, SheetID: sheet, Kind: kind, Cell: CellRef{Row: row, Column: column}}
}

func TestApplyCellMutations_WritesNativeCellsAndPreservesOpaqueParts(t *testing.T) {
	entries := cellMutationFixture()
	original := buildZip(t, entries)
	operations := []CellMutation{
		func() CellMutation { op := mutation("empty", "7", CellSetValue, 0, 0); op.Value = ""; return op }(),
		func() CellMutation { op := mutation("zero", "7", CellSetValue, 0, 1); op.Value = 0; return op }(),
		func() CellMutation { op := mutation("false", "7", CellSetValue, 0, 2); op.Value = false; return op }(),
		mutation("clear-old-formula", "7", CellClearFormula, 0, 3),
		func() CellMutation {
			op := mutation("formula", "7", CellSetFormula, 0, 4)
			op.Formula = `=SUM(B1,2)&"<x>"`
			return op
		}(),
		func() CellMutation {
			op := mutation("spaced", "7", CellSetValue, 1, 2)
			op.Value = "  spaced  "
			return op
		}(),
		mutation("clear-old-value", "7", CellClearValue, 2, 1),
		func() CellMutation { op := mutation("new-row", "7", CellSetValue, 3, 0); op.Value = 42.5; return op }(),
		func() CellMutation {
			op := mutation("other-sheet", "9", CellSetValue, 0, 0)
			op.Value = "native"
			return op
		}(),
	}

	output, err := ApplyCellMutations(original, operations)
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	checks := []string{
		`<x:dimension ref="A1:E4"/>`,
		`<x:row r="1">`,
		`<x:c r="A1" s="7" t="inlineStr"><x:is><x:t></x:t></x:is><x:extLst><x:ext uri="keep"><x:payload/></x:ext></x:extLst></x:c>`,
		`<x:c r="B1"><x:v>0</x:v></x:c>`,
		`<x:c r="C1" t="b"><x:v>0</x:v></x:c>`,
		`<x:c r="D1" s="3"/>`,
		`<x:c r="E1"><x:f>SUM(B1,2)&amp;&#34;&lt;x&gt;&#34;</x:f></x:c>`,
		`<x:row r="2"><x:c r="C2" t="inlineStr"><x:is><x:t xml:space="preserve">  spaced  </x:t></x:is></x:c></x:row>`,
		`<x:row r="3"><x:c r="B3" s="2"/></x:row>`,
		`<x:row r="4"><x:c r="A4"><x:v>42.5</x:v></x:c></x:row>`,
		`<x:extLst><x:ext uri="worksheet-keep"/></x:extLst>`,
	}
	for _, want := range checks {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet is missing %q\n%s", want, sheet)
		}
	}
	if strings.Index(sheet, `r="B1"`) > strings.Index(sheet, `r="C1"`) || strings.Index(sheet, `r="C1"`) > strings.Index(sheet, `r="D1"`) {
		t.Errorf("inserted cells are not ordered: %s", sheet)
	}
	other := readEntry(t, output, "xl/worksheets/sheet2.xml")
	if !strings.Contains(other, `<c r="A1" t="inlineStr"><is><t>native</t></is></c>`) {
		t.Errorf("second sheet was not resolved by stable sheet id: %s", other)
	}
	for _, name := range []string{"xl/charts/chart1.xml", "customXml/item1.xml", "xl/_rels/workbook.xml.rels"} {
		if got := readEntry(t, output, name); got != entries[name] {
			t.Errorf("untouched part %q changed\n got: %s\nwant: %s", name, got, entries[name])
		}
	}
	workbook := readEntry(t, output, "xl/workbook.xml")
	if !strings.Contains(workbook, `<calcPr calcId="0" fullCalcOnLoad="1" forceFullCalc="1" calcCompleted="0"/>`) {
		t.Errorf("workbook was not marked for deterministic full recalculation: %s", workbook)
	}
}

func TestApplyCellMutations_RefusesUnsafeWorksheetRouting(t *testing.T) {
	tests := []struct {
		name      string
		mutate    func(map[string]string)
		wantError string
	}{
		{
			name: "duplicate stable sheet id",
			mutate: func(entries map[string]string) {
				entries["xl/workbook.xml"] = strings.Replace(entries["xl/workbook.xml"], `sheetId="9"`, `sheetId="7"`, 1)
			},
			wantError: `stable sheet id "7" is duplicated`,
		},
		{
			name: "relationship id in foreign namespace",
			mutate: func(entries map[string]string) {
				entries["xl/workbook.xml"] = strings.Replace(entries["xl/workbook.xml"], `r:id="rIdData"`, `xmlns:local="urn:local" local:id="rIdData"`, 1)
			},
			wantError: "no relationship id in a supported namespace",
		},
		{
			name: "nested sheet extension collision",
			mutate: func(entries map[string]string) {
				entries["xl/workbook.xml"] = strings.Replace(entries["xl/workbook.xml"], `sheetId="7"`, `sheetId="8"`, 1)
				entries["xl/workbook.xml"] = strings.Replace(entries["xl/workbook.xml"], `</workbook>`, `<extLst><ext><sheet name="Injected" sheetId="7" r:id="rIdData"/></ext></extLst></workbook>`, 1)
			},
			wantError: `stable sheet id "7" not found`,
		},
		{
			name: "duplicate relationship id",
			mutate: func(entries map[string]string) {
				duplicate := `<Relationship Id="rIdData" Type="` + relTypeWorksheetTransitional + `" Target="worksheets/sheet2.xml"/>`
				entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], `</Relationships>`, duplicate+`</Relationships>`, 1)
			},
			wantError: `relationship "rIdData" is duplicated`,
		},
		{
			name: "external worksheet relationship",
			mutate: func(entries map[string]string) {
				entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], `Target="worksheets/sheet1.xml"`, `Target="https://example.test/sheet1.xml" TargetMode="External"`, 1)
			},
			wantError: `relationship "rIdData" has an external target`,
		},
		{
			name: "unknown target mode",
			mutate: func(entries map[string]string) {
				entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], `Target="worksheets/sheet1.xml"`, `Target="worksheets/sheet1.xml" TargetMode="Sideways"`, 1)
			},
			wantError: `unsupported target mode "Sideways"`,
		},
		{
			name: "non worksheet relationship type",
			mutate: func(entries map[string]string) {
				entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], relTypeWorksheetTransitional, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", 1)
			},
			wantError: "has unsupported type",
		},
		{
			name: "relationship traversal above package root",
			mutate: func(entries map[string]string) {
				entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], `Target="worksheets/sheet1.xml"`, `Target="../../customXml/item1.xml"`, 1)
			},
			wantError: "traverses above the package root",
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			entries := cellMutationFixture()
			test.mutate(entries)
			op := mutation("safe-routing", "7", CellSetValue, 0, 0)
			op.Value = "must-not-route"
			output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
			if err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("expected error containing %q, output=%d err=%v", test.wantError, len(output), err)
			}
			if output != nil {
				t.Fatal("refused worksheet routing returned partial output")
			}
		})
	}
}

func TestApplyCellMutations_AcceptsStrictWorksheetRouting(t *testing.T) {
	entries := cellMutationFixture()
	entries["_rels/.rels"] = strings.ReplaceAll(entries["_rels/.rels"], relTypeOfficeDocumentTransitional, relTypeOfficeDocumentStrict)
	entries["xl/workbook.xml"] = strings.ReplaceAll(entries["xl/workbook.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	entries["xl/workbook.xml"] = strings.ReplaceAll(entries["xl/workbook.xml"], officeRelNamespaceTransitional, officeRelNamespaceStrict)
	entries["xl/_rels/workbook.xml.rels"] = strings.ReplaceAll(entries["xl/_rels/workbook.xml.rels"], relTypeWorksheetTransitional, relTypeWorksheetStrict)
	entries["xl/worksheets/sheet1.xml"] = strings.ReplaceAll(entries["xl/worksheets/sheet1.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	op := mutation("strict-routing", "7", CellSetValue, 0, 0)
	op.Value = "strict-native"

	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if sheet := readEntry(t, output, "xl/worksheets/sheet1.xml"); !strings.Contains(sheet, `<x:t>strict-native</x:t>`) {
		t.Fatalf("Strict worksheet was not routed: %s", sheet)
	}
}

func TestApplyCellMutations_RefusesMalformedRootRoutingRelationships(t *testing.T) {
	office := `<Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `" Target="xl/workbook.xml"/>`
	tests := map[string]string{
		"spoofed root namespace":  `<Relationships xmlns="urn:spoof">` + office + `</Relationships>`,
		"spoofed child namespace": `<Relationships xmlns="` + packageRelationshipsNamespace + `" xmlns:s="urn:spoof"><s:Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `" Target="xl/workbook.xml"/></Relationships>`,
		"unexpected direct child": `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Unexpected/>` + office + `</Relationships>`,
		"missing Id":              `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Type="` + relTypeOfficeDocumentTransitional + `" Target="xl/workbook.xml"/></Relationships>`,
		"missing Type":            `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Target="xl/workbook.xml"/></Relationships>`,
		"missing Target":          `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `"/></Relationships>`,
		"duplicate attribute":     `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Id="again" Type="` + relTypeOfficeDocumentTransitional + `" Target="xl/workbook.xml"/></Relationships>`,
		"processing instruction":  `<Relationships xmlns="` + packageRelationshipsNamespace + `><?unsafe value?>` + office + `</Relationships>`,
		"XML directive":           `<!DOCTYPE Relationships><Relationships xmlns="` + packageRelationshipsNamespace + `">` + office + `</Relationships>`,
	}
	for name, relationships := range tests {
		t.Run(name, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["_rels/.rels"] = relationships
			op := mutation("root-routing", "7", CellSetValue, 0, 0)
			op.Value = "must-not-route"
			output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
			if err == nil || output != nil {
				t.Fatalf("expected root routing refusal, output=%d err=%v", len(output), err)
			}
		})
	}
}

func TestApplyCellMutations_RefusesMalformedWorkbookRoutingRelationships(t *testing.T) {
	worksheet := `<Relationship Id="rIdData" Type="` + relTypeWorksheetTransitional + `" Target="worksheets/sheet1.xml"/>`
	tests := map[string]string{
		"spoofed root namespace":  `<Relationships xmlns="urn:spoof">` + worksheet + `</Relationships>`,
		"spoofed child namespace": `<Relationships xmlns="` + packageRelationshipsNamespace + `" xmlns:s="urn:spoof"><s:Relationship Id="rIdData" Type="` + relTypeWorksheetTransitional + `" Target="worksheets/sheet1.xml"/></Relationships>`,
		"unexpected direct child": `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Unexpected/>` + worksheet + `</Relationships>`,
		"missing Id":              `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Type="` + relTypeWorksheetTransitional + `" Target="worksheets/sheet1.xml"/></Relationships>`,
		"missing Type":            `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdData" Target="worksheets/sheet1.xml"/></Relationships>`,
		"missing Target":          `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdData" Type="` + relTypeWorksheetTransitional + `"/></Relationships>`,
		"duplicate attribute":     `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdData" Id="again" Type="` + relTypeWorksheetTransitional + `" Target="worksheets/sheet1.xml"/></Relationships>`,
		"processing instruction":  `<Relationships xmlns="` + packageRelationshipsNamespace + `><?unsafe value?>` + worksheet + `</Relationships>`,
		"XML directive":           `<!DOCTYPE Relationships><Relationships xmlns="` + packageRelationshipsNamespace + `">` + worksheet + `</Relationships>`,
	}
	for name, relationships := range tests {
		t.Run(name, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["xl/_rels/workbook.xml.rels"] = relationships
			op := mutation("workbook-routing", "7", CellSetValue, 0, 0)
			op.Value = "must-not-route"
			output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
			if err == nil || output != nil {
				t.Fatalf("expected workbook routing refusal, output=%d err=%v", len(output), err)
			}
		})
	}
}

func TestApplyCellMutations_RoutesRelocatedPartsAndRemovesCaseVariantCalcChain(t *testing.T) {
	entries := cellMutationFixture()
	workbook := entries["xl/workbook.xml"]
	workbookRels := entries["xl/_rels/workbook.xml.rels"]
	sheet := entries["xl/worksheets/sheet1.xml"]
	delete(entries, "xl/workbook.xml")
	delete(entries, "xl/_rels/workbook.xml.rels")
	delete(entries, "xl/worksheets/sheet1.xml")
	delete(entries, "[Content_Types].xml")
	entries["Custom/Office/Book.XML"] = workbook
	workbookRels = strings.Replace(workbookRels, `Target="worksheets/sheet1.xml"`, `Target="../../CUSTOM/SHEETS/DATA.XML"`, 1)
	workbookRels = strings.Replace(workbookRels, `</Relationships>`, `<Relationship Id="rIdCalc" Type="`+relTypeCalcChainTransitional+`" Target="../../xl/CALCCHAIN.xml"/></Relationships>`, 1)
	entries["Custom/Office/_rels/Book.XML.rels"] = workbookRels
	entries["Custom/Sheets/Data.XML"] = sheet
	entries["XL/CalcChain.XML"] = `<calcChain xmlns="` + spreadsheetMLTransitional + `"><c r="A1"/></calcChain>`
	entries["[CONTENT_TYPES].XML"] = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/calcchain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`
	entries["_rels/.rels"] = `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `" Target="/custom/office/book.xml"/></Relationships>`
	op := mutation("relocated-cell", "7", CellSetValue, 0, 0)
	op.Value = "native-relocated"

	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if relocated := readEntry(t, output, "Custom/Sheets/Data.XML"); !strings.Contains(relocated, "native-relocated") {
		t.Fatalf("relocated worksheet was not updated: %s", relocated)
	}
	if relocatedWorkbook := readEntry(t, output, "Custom/Office/Book.XML"); !strings.Contains(relocatedWorkbook, `fullCalcOnLoad="1"`) {
		t.Fatalf("relocated workbook was not marked for recalculation: %s", relocatedWorkbook)
	}
	if hasZipEntry(t, output, "XL/CalcChain.XML") {
		t.Fatal("case-variant calc-chain part survived")
	}
	if rels := readEntry(t, output, "Custom/Office/_rels/Book.XML.rels"); strings.Contains(rels, "rIdCalc") {
		t.Fatalf("case-variant calc-chain relationship survived: %s", rels)
	}
	if contentTypes := readEntry(t, output, "[CONTENT_TYPES].XML"); strings.Contains(strings.ToLower(contentTypes), "calcchain") {
		t.Fatalf("case-variant calc-chain override survived: %s", contentTypes)
	}
}

func TestApplyCellMutations_RemovesCalcChainAndForcesFullRecalculation(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/workbook.xml"] = strings.Replace(
		entries["xl/workbook.xml"],
		`</workbook>`,
		`<calcPr calcId="191029" calcMode="manual" custom="literal fullCalcOnLoad='keep'"/><extLst><ext uri="keep"/></extLst></workbook>`,
		1,
	)
	entries["xl/_rels/workbook.xml.rels"] = strings.Replace(
		entries["xl/_rels/workbook.xml.rels"],
		`</Relationships>`,
		`<Relationship Id="rIdCalc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml" TargetMode="Internal"/></Relationships>`,
		1,
	)
	entries["[Content_Types].xml"] = strings.Replace(
		entries["[Content_Types].xml"],
		`</Types>`,
		`<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>`,
		1,
	)
	entries["xl/calcChain.xml"] = `<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="D1" i="1"/></calcChain>`
	op := mutation("formula-membership", "7", CellSetValue, 0, 3)
	op.Value = 7

	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if hasZipEntry(t, output, "xl/calcChain.xml") {
		t.Fatal("stale calc-chain part survived a cell mutation")
	}
	if rels := readEntry(t, output, "xl/_rels/workbook.xml.rels"); strings.Contains(rels, "calcChain") || strings.Contains(rels, "rIdCalc") {
		t.Errorf("calc-chain relationship survived: %s", rels)
	}
	if contentTypes := readEntry(t, output, "[Content_Types].xml"); strings.Contains(contentTypes, "calcChain") {
		t.Errorf("calc-chain content type survived: %s", contentTypes)
	}
	workbook := readEntry(t, output, "xl/workbook.xml")
	for _, want := range []string{
		`calcId="0"`, `fullCalcOnLoad="1"`, `forceFullCalc="1"`, `calcCompleted="0"`,
		`calcMode="manual"`, `custom="literal fullCalcOnLoad='keep'"`, `<extLst><ext uri="keep"/></extLst>`,
	} {
		if !strings.Contains(workbook, want) {
			t.Errorf("recalculation workbook is missing %q: %s", want, workbook)
		}
	}
}

func TestApplyCellMutations_RefusesUnsafeCalcChainTargetModes(t *testing.T) {
	for _, test := range []struct {
		mode, wantError string
	}{
		{mode: "External", wantError: "calc-chain relationship is external"},
		{mode: "Sideways", wantError: `unsupported target mode "Sideways"`},
	} {
		t.Run(test.mode, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["xl/_rels/workbook.xml.rels"] = strings.Replace(
				entries["xl/_rels/workbook.xml.rels"],
				`</Relationships>`,
				`<Relationship Id="rIdCalc" Type="`+relTypeCalcChainTransitional+`" Target="calcChain.xml" TargetMode="`+test.mode+`"/></Relationships>`,
				1,
			)
			entries["xl/calcChain.xml"] = `<calcChain xmlns="` + spreadsheetMLTransitional + `"><c r="A1"/></calcChain>`
			op := mutation("calc-mode", "7", CellSetValue, 0, 0)
			op.Value = "changed"

			output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
			if err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("expected error containing %q, output=%d err=%v", test.wantError, len(output), err)
			}
			if output != nil {
				t.Fatal("refused calc-chain target mode returned partial output")
			}
		})
	}
}

func TestApplyCellMutations_RefusesFormulaGroupInteriors(t *testing.T) {
	for _, formulaType := range []string{"array", "dataTable", "shared"} {
		t.Run(formulaType, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["xl/worksheets/sheet1.xml"] = `<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
				`<row r="1"><c r="A1"><f t="` + formulaType + `" ref="A1:C2">SUM(A3:C3)</f><v>1</v></c></row>` +
				`<row r="2"><c r="C2"><v>2</v></c></row></sheetData></worksheet>`
			op := mutation("group-interior", "7", CellSetValue, 1, 2)
			op.Value = "would-corrupt-group"
			output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
			if err == nil || !strings.Contains(err.Error(), formulaType+" formula range") {
				t.Fatalf("expected %s group refusal, output=%d err=%v", formulaType, len(output), err)
			}
			if output != nil {
				t.Fatal("refused formula-group edit returned partial output")
			}
		})
	}
}

func TestApplyCellMutations_ProtectsStrictMergedCells(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = strings.ReplaceAll(entries["xl/worksheets/sheet1.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`</x:sheetData>`,
		`</x:sheetData><x:mergeCells count="1"><x:mergeCell ref="B2:D4"/></x:mergeCells>`,
		1,
	)
	op := mutation("strict-merged-interior", "7", CellSetValue, 1, 2)
	op.Value = "invisible"
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err == nil || !strings.Contains(err.Error(), "inside merged range") {
		t.Fatalf("expected Strict merged-cell refusal, output=%d err=%v", len(output), err)
	}
}

func TestApplyCellMutations_RefusesModernCellMetadata(t *testing.T) {
	for _, metadataAttribute := range []string{"cm", "vm"} {
		t.Run(metadataAttribute, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["xl/worksheets/sheet1.xml"] = strings.Replace(
				entries["xl/worksheets/sheet1.xml"],
				`<x:c r="A1"`,
				`<x:c r="A1" `+metadataAttribute+`="2"`,
				1,
			)
			op := mutation("metadata-aware", "7", CellSetValue, 0, 0)
			op.Value = "plain-value"
			output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
			if err == nil || !strings.Contains(err.Error(), "metadata-aware edit") {
				t.Fatalf("expected %s metadata refusal, output=%d err=%v", metadataAttribute, len(output), err)
			}
		})
	}
}

func TestApplyCellMutations_RewritesOnlyParsedAttributes(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`<x:dimension ref="A1:D3"/>`,
		`<x:dimension custom="literal ref='Z99'" ref="A1:D3"/>`,
		1,
	)
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`<x:c r="A1" s="7" t="s">`,
		`<x:c r="A1" custom="literal t='keep'" s="7" t="s">`,
		1,
	)
	first := mutation("parsed-cell-type", "7", CellSetValue, 0, 0)
	first.Value = "new"
	second := mutation("parsed-dimension-ref", "7", CellSetValue, 0, 4)
	second.Value = 5
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{first, second})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{`custom="literal t='keep'"`, `custom="literal ref='Z99'"`, `ref="A1:E3"`} {
		if !strings.Contains(sheet, want) {
			t.Errorf("parsed attribute rewrite lost %q: %s", want, sheet)
		}
	}
}

func TestApplyCellMutations_EncodesSpreadsheetStrings(t *testing.T) {
	entries := cellMutationFixture()
	op := mutation("st-xstring", "7", CellSetValue, 1, 0)
	op.Value = "literal _x000A_ and _X0041_ control:\x01 & < >"
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	want := `literal _x005F_x000A_ and _x005F_X0041_ control:_x0001_ &amp; &lt; &gt;`
	if !strings.Contains(sheet, want) {
		t.Errorf("SpreadsheetML string was not encoded safely; want %q in %s", want, sheet)
	}
}

func TestApplyCellMutations_LastOperationWinsAndMissingClearIsContentNoop(t *testing.T) {
	entries := cellMutationFixture()
	original := buildZip(t, entries)
	first := mutation("first", "7", CellSetValue, 0, 0)
	first.Value = "discarded"
	last := mutation("last", "7", CellSetValue, 0, 0)
	last.Value = true
	clearMissing := mutation("clear-missing", "7", CellClearValue, 8, 8)

	output, err := ApplyCellMutations(original, []CellMutation{first, last, clearMissing})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, `<x:c r="A1" s="7" t="b"><x:v>1</x:v><x:extLst>`) {
		t.Errorf("last operation did not win or extension was lost: %s", sheet)
	}
	if strings.Contains(sheet, "discarded") || strings.Contains(sheet, `r="I9"`) {
		t.Errorf("superseded or missing-clear mutation leaked into sheet: %s", sheet)
	}

	noopOutput, err := ApplyCellMutations(original, []CellMutation{mutation("only-clear", "7", CellClearValue, 8, 8)})
	if err != nil {
		t.Fatal(err)
	}
	if got := readEntry(t, noopOutput, "xl/worksheets/sheet1.xml"); got != entries["xl/worksheets/sheet1.xml"] {
		t.Errorf("clear of a missing cell changed worksheet content:\n%s", got)
	}
	if !bytes.Equal(noopOutput, original) {
		t.Error("clear of a missing cell should preserve the entire XLSX byte-for-byte")
	}
}

func TestApplyCellMutations_RefusesInvalidBatchAtomically(t *testing.T) {
	original := buildZip(t, cellMutationFixture())
	valid := mutation("valid", "7", CellSetValue, 0, 0)
	valid.Value = "safe"
	cases := map[string][]CellMutation{
		"empty batch":         nil,
		"oversized batch":     make([]CellMutation, maxCellMutations+1),
		"duplicate operation": {valid, valid},
		"missing sheet":       {func() CellMutation { op := valid; op.OperationID = "missing"; op.SheetID = "404"; return op }()},
		"bad row":             {func() CellMutation { op := valid; op.OperationID = "row"; op.Cell.Row = excelMaxRows; return op }()},
		"bad column": {func() CellMutation {
			op := valid
			op.OperationID = "column"
			op.Cell.Column = excelMaxColumns
			return op
		}()},
		"nonfinite": {func() CellMutation { op := valid; op.OperationID = "nan"; op.Value = math.NaN(); return op }()},
		"invalid UTF-8 value": {func() CellMutation {
			op := valid
			op.OperationID = "utf8-value"
			op.Value = string([]byte{0xff})
			return op
		}()},
		"unknown kind": {func() CellMutation { op := valid; op.OperationID = "unknown"; op.Kind = "row.insert"; return op }()},
		"invalid formula": {func() CellMutation {
			op := mutation("formula", "7", CellSetFormula, 0, 0)
			op.Formula = "SUM(A1)"
			return op
		}()},
		"invalid XML formula": {func() CellMutation {
			op := mutation("xml-formula", "7", CellSetFormula, 0, 0)
			op.Formula = "=bad\x00formula"
			return op
		}()},
		"invalid identifier":   {func() CellMutation { op := valid; op.OperationID = "has space"; return op }()},
		"literal with formula": {func() CellMutation { op := valid; op.OperationID = "mixed"; op.Formula = "=A1"; return op }()},
	}
	for name, operations := range cases {
		t.Run(name, func(t *testing.T) {
			output, err := ApplyCellMutations(original, operations)
			if err == nil {
				t.Fatalf("expected refusal, produced %d bytes", len(output))
			}
			if output != nil {
				t.Fatalf("refused batch returned partial output")
			}
		})
	}
}

func TestApplyCellMutations_RefusesUnknownExistingCellMarkup(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`<x:v>0</x:v><x:extLst>`,
		`<x:v>0</x:v><x:unsupported/><x:extLst>`,
		1,
	)
	original := buildZip(t, entries)
	op := mutation("refuse", "7", CellSetValue, 0, 0)
	op.Value = "would-drop-unknown-markup"
	output, err := ApplyCellMutations(original, []CellMutation{op})
	if err == nil || !strings.Contains(err.Error(), "unsupported direct child") {
		t.Fatalf("expected unsupported-markup refusal, output=%d err=%v", len(output), err)
	}
	if output != nil {
		t.Fatal("refused edit returned partial output")
	}
}

func TestApplyCellMutations_RefusesGroupAwareFormulaAndMergedInterior(t *testing.T) {
	t.Run("attributed formula", func(t *testing.T) {
		entries := cellMutationFixture()
		entries["xl/worksheets/sheet1.xml"] = strings.Replace(
			entries["xl/worksheets/sheet1.xml"],
			`<x:f>OLD()</x:f>`,
			`<x:f t="shared" si="2" ref="D1:D3">OLD()</x:f>`,
			1,
		)
		op := mutation("shared-formula", "7", CellSetFormula, 0, 3)
		op.Formula = "=A1+1"
		output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
		if err == nil || !strings.Contains(err.Error(), "group-aware edit") {
			t.Fatalf("expected shared-formula refusal, output=%d err=%v", len(output), err)
		}
		if output != nil {
			t.Fatal("refused shared-formula edit returned partial output")
		}
	})

	t.Run("merged interior", func(t *testing.T) {
		entries := cellMutationFixture()
		entries["xl/worksheets/sheet1.xml"] = strings.Replace(
			entries["xl/worksheets/sheet1.xml"],
			`</x:sheetData>`,
			`</x:sheetData><x:mergeCells count="1"><x:mergeCell ref="B2:D4"/></x:mergeCells>`,
			1,
		)
		op := mutation("merged-interior", "7", CellSetValue, 1, 2)
		op.Value = "invisible"
		output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
		if err == nil || !strings.Contains(err.Error(), "inside merged range") {
			t.Fatalf("expected merged-interior refusal, output=%d err=%v", len(output), err)
		}
		if output != nil {
			t.Fatal("refused merged-cell edit returned partial output")
		}
	})
}
