package xlsxpatch

import (
	"bytes"
	"math"
	"strings"
	"testing"
)

func layoutMutationFixture() map[string]string {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = `<?xml version="1.0"?><x:worksheet xmlns:x="` + spreadsheetMLTransitional + `"><x:dimension ref="A1:C3"/><x:sheetViews><x:sheetView workbookViewId="0"/></x:sheetViews><x:sheetFormatPr defaultRowHeight="15"/>` +
		`<x:cols><x:col xmlns:z="urn:opaque" min="1" max="3" width="8.43" customWidth="1" style="4" bestFit="1" z:opaque="literal width='keep'"/><!--between--><x:col min="5" max="6" width="20" customWidth="1" hidden="1" outlineLevel="2"></x:col><x:col min="8" max="8" width="30" customWidth="1"></x:col></x:cols>` +
		`<x:sheetData><x:row r="1" hidden="1" ht="0" customHeight="1" customWidth="keep" spans="1:3" data="literal ht='keep'"><x:c r="A1"><x:v>keep-cell</x:v></x:c><x:extLst><x:ext uri="keep-row"/></x:extLst></x:row><x:row r="3" custom="keep"><x:c r="C3" t="inlineStr"><x:is><x:t>opaque</x:t></x:is></x:c></x:row></x:sheetData>` +
		`<x:mergeCells count="1"><x:mergeCell ref="A1:B1"/></x:mergeCells><x:extLst><x:ext uri="keep-sheet"/></x:extLst></x:worksheet>`
	return entries
}

func layoutMutation(id, sheet string, kind LayoutMutationKind) LayoutMutation {
	return LayoutMutation{OperationID: id, SheetID: sheet, Kind: kind}
}

func TestApplyLayoutMutations_WritesNativeDimensionsAndPreservesOpaqueContent(t *testing.T) {
	entries := layoutMutationFixture()
	original := buildZip(t, entries)
	operations := []LayoutMutation{
		func() LayoutMutation {
			op := layoutMutation("row-hide-first", "7", RowSetHeight)
			op.Row, op.HeightPoints = 0, 0
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("row-last-wins", "7", RowSetHeight)
			op.Row, op.HeightPoints = 0, 22.5
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("row-new-hidden", "7", RowSetHeight)
			op.Row, op.HeightPoints = 1, 0
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("row-existing", "7", RowSetHeight)
			op.Row, op.HeightPoints = 2, 18.25
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("other-sheet", "9", RowSetHeight)
			op.Row, op.HeightPoints = 0, 19
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("col-two-old", "7", ColumnSetWidth)
			op.Column, op.Width = 1, 9
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("col-two-last", "7", ColumnSetWidth)
			op.Column, op.Width = 1, 10.25
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("col-three-hidden", "7", ColumnSetWidth)
			op.Column, op.Width = 2, 0
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("col-gap", "7", ColumnSetWidth)
			op.Column, op.Width = 3, 12.25
			return op
		}(),
		func() LayoutMutation {
			op := layoutMutation("col-five-unhide", "7", ColumnSetWidth)
			op.Column, op.Width = 4, 14
			return op
		}(),
	}

	output, err := ApplyLayoutMutations(original, operations)
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<x:row r="1" ht="22.5" customHeight="1" customWidth="keep" spans="1:3" data="literal ht='keep'">`,
		`<x:c r="A1"><x:v>keep-cell</x:v></x:c><x:extLst><x:ext uri="keep-row"/></x:extLst></x:row>`,
		`<x:row r="2" ht="0" customHeight="1" hidden="1"/>`,
		`<x:row r="3" custom="keep" ht="18.25" customHeight="1"><x:c r="C3" t="inlineStr"><x:is><x:t>opaque</x:t></x:is></x:c></x:row>`,
		`<x:col xmlns:z="urn:opaque" min="1" max="1" width="8.43" customWidth="1" style="4" bestFit="1" z:opaque="literal width='keep'"/>`,
		`<x:col xmlns:z="urn:opaque" min="2" max="2" width="10.25" customWidth="1" style="4" bestFit="1" z:opaque="literal width='keep'"/>`,
		`<x:col xmlns:z="urn:opaque" min="3" max="3" width="0" customWidth="1" style="4" bestFit="1" z:opaque="literal width='keep'" hidden="1"/>`,
		`<x:col min="4" max="4" width="12.25" customWidth="1"/>`,
		`<x:col min="5" max="5" width="14" customWidth="1" outlineLevel="2"/>`,
		`<x:col min="6" max="6" width="20" customWidth="1" hidden="1" outlineLevel="2"/>`,
		`<!--between-->`,
		`<x:col min="8" max="8" width="30" customWidth="1"></x:col>`,
		`<x:mergeCells count="1"><x:mergeCell ref="A1:B1"/></x:mergeCells><x:extLst><x:ext uri="keep-sheet"/></x:extLst>`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet is missing %q\n%s", want, sheet)
		}
	}
	other := readEntry(t, output, "xl/worksheets/sheet2.xml")
	if !strings.Contains(other, `<row r="1" ht="19" customHeight="1"/>`) {
		t.Errorf("stable sheet id did not route to the second sheet: %s", other)
	}
	for _, name := range []string{"_rels/.rels", "xl/workbook.xml", "xl/_rels/workbook.xml.rels", "xl/charts/chart1.xml", "customXml/item1.xml"} {
		if got := readEntry(t, output, name); got != entries[name] {
			t.Errorf("untouched part %q changed\n got: %s\nwant: %s", name, got, entries[name])
		}
	}
}

func TestApplyLayoutMutations_InsertsColsInSchemaOrderAndExpandsSelfClosingSheetData(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = `<x:worksheet xmlns:x="` + spreadsheetMLTransitional + `"><x:sheetViews/><x:sheetFormatPr defaultRowHeight="15"/><x:sheetData/><x:extLst><x:ext uri="keep"/></x:extLst></x:worksheet>`
	row := layoutMutation("new-row", "7", RowSetHeight)
	row.Row, row.HeightPoints = 1, 20
	column := layoutMutation("new-column", "7", ColumnSetWidth)
	column.Column, column.Width = 0, 8.5

	output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{row, column})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	want := `<x:sheetFormatPr defaultRowHeight="15"/><x:cols><x:col min="1" max="1" width="8.5" customWidth="1"/></x:cols><x:sheetData><x:row r="2" ht="20" customHeight="1"/></x:sheetData><x:extLst>`
	if !strings.Contains(sheet, want) {
		t.Fatalf("cols/sheetData insertion is not schema ordered:\n%s", sheet)
	}
}

func TestApplyLayoutMutations_InfersOmittedRowNumbers(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData><row><c r="A1"><v>first</v></c></row><row r="3"><c r="A3"><v>third</v></c></row><row><c r="A4"><v>fourth</v></c></row></sheetData></worksheet>`
	first := layoutMutation("first", "7", RowSetHeight)
	first.Row, first.HeightPoints = 0, 10
	second := layoutMutation("insert-second", "7", RowSetHeight)
	second.Row, second.HeightPoints = 1, 15
	fourth := layoutMutation("fourth", "7", RowSetHeight)
	fourth.Row, fourth.HeightPoints = 3, 20

	output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{first, second, fourth})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<row ht="10" customHeight="1"><c r="A1"><v>first</v></c></row><row r="2" ht="15" customHeight="1"/><row r="3">`,
		`<row ht="20" customHeight="1"><c r="A4"><v>fourth</v></c></row>`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("omitted row inference missing %q: %s", want, sheet)
		}
	}
}

func TestApplyLayoutMutations_AcceptsStrictSpreadsheetML(t *testing.T) {
	entries := layoutMutationFixture()
	entries["_rels/.rels"] = strings.ReplaceAll(entries["_rels/.rels"], relTypeOfficeDocumentTransitional, relTypeOfficeDocumentStrict)
	entries["xl/workbook.xml"] = strings.ReplaceAll(entries["xl/workbook.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	entries["xl/workbook.xml"] = strings.ReplaceAll(entries["xl/workbook.xml"], officeRelNamespaceTransitional, officeRelNamespaceStrict)
	entries["xl/_rels/workbook.xml.rels"] = strings.ReplaceAll(entries["xl/_rels/workbook.xml.rels"], relTypeWorksheetTransitional, relTypeWorksheetStrict)
	entries["xl/worksheets/sheet1.xml"] = strings.ReplaceAll(entries["xl/worksheets/sheet1.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	row := layoutMutation("strict-row", "7", RowSetHeight)
	row.Row, row.HeightPoints = 0, 17
	column := layoutMutation("strict-column", "7", ColumnSetWidth)
	column.Column, column.Width = 0, 11

	output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{row, column})
	if err != nil {
		t.Fatal(err)
	}
	if sheet := readEntry(t, output, "xl/worksheets/sheet1.xml"); !strings.Contains(sheet, `ht="17"`) || !strings.Contains(sheet, `width="11"`) {
		t.Fatalf("Strict worksheet layout was not updated: %s", sheet)
	}
}

func TestApplyLayoutMutations_IsByteNoopWhenDimensionsAlreadyMatch(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><cols><col min="1" max="3" width="8.500" customWidth="true" hidden="false"></col></cols><sheetData><row r="1" ht="20.0" customHeight="true" hidden="0"><c r="A1"><v>keep</v></c></row></sheetData></worksheet>`
	original := buildZip(t, entries)
	row := layoutMutation("same-row", "7", RowSetHeight)
	row.HeightPoints = 20
	column := layoutMutation("same-column", "7", ColumnSetWidth)
	column.Width = 8.5

	output, err := ApplyLayoutMutations(original, []LayoutMutation{row, column})
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(output, original) {
		t.Fatal("already-matching layout was not a byte-level no-op")
	}
}

func TestApplyLayoutMutations_RefusesInvalidBatchAtomically(t *testing.T) {
	original := buildZip(t, layoutMutationFixture())
	validRow := layoutMutation("valid-row", "7", RowSetHeight)
	validRow.HeightPoints = 20
	validColumn := layoutMutation("valid-column", "7", ColumnSetWidth)
	validColumn.Width = 10
	cases := map[string][]LayoutMutation{
		"empty batch":               {},
		"oversized batch":           make([]LayoutMutation, maxCellMutations+1),
		"duplicate operation id":    {validRow, validRow},
		"bad operation id":          {func() LayoutMutation { op := validRow; op.OperationID = "bad id"; return op }()},
		"bad sheet id":              {func() LayoutMutation { op := validRow; op.SheetID = " sheet"; return op }()},
		"missing sheet":             {func() LayoutMutation { op := validRow; op.SheetID = "404"; return op }()},
		"negative row":              {func() LayoutMutation { op := validRow; op.Row = -1; return op }()},
		"row overflow":              {func() LayoutMutation { op := validRow; op.Row = excelMaxRows; return op }()},
		"negative height":           {func() LayoutMutation { op := validRow; op.HeightPoints = -1; return op }()},
		"negative zero height":      {func() LayoutMutation { op := validRow; op.HeightPoints = math.Copysign(0, -1); return op }()},
		"infinite height":           {func() LayoutMutation { op := validRow; op.HeightPoints = math.Inf(1); return op }()},
		"height overflow":           {func() LayoutMutation { op := validRow; op.HeightPoints = maxRowHeightPoints + 0.1; return op }()},
		"negative column":           {func() LayoutMutation { op := validColumn; op.Column = -1; return op }()},
		"column overflow":           {func() LayoutMutation { op := validColumn; op.Column = excelMaxColumns; return op }()},
		"negative width":            {func() LayoutMutation { op := validColumn; op.Width = -1; return op }()},
		"negative zero width":       {func() LayoutMutation { op := validColumn; op.Width = math.Copysign(0, -1); return op }()},
		"nan width":                 {func() LayoutMutation { op := validColumn; op.Width = math.NaN(); return op }()},
		"width overflow":            {func() LayoutMutation { op := validColumn; op.Width = maxColumnWidth + 0.1; return op }()},
		"row carries column fields": {func() LayoutMutation { op := validRow; op.Column, op.Width = 3, 10; return op }()},
		"row carries negative zero width": {func() LayoutMutation {
			op := validRow
			op.Width = math.Copysign(0, -1)
			return op
		}()},
		"column carries row fields": {func() LayoutMutation { op := validColumn; op.Row, op.HeightPoints = 3, 10; return op }()},
		"column carries negative zero height": {func() LayoutMutation {
			op := validColumn
			op.HeightPoints = math.Copysign(0, -1)
			return op
		}()},
		"unknown kind": {func() LayoutMutation { op := validRow; op.Kind = "row.insert"; return op }()},
	}
	for name, operations := range cases {
		t.Run(name, func(t *testing.T) {
			output, err := ApplyLayoutMutations(original, operations)
			if err == nil {
				t.Fatalf("expected refusal, output=%d", len(output))
			}
			if output != nil {
				t.Fatal("invalid batch returned partial output")
			}
		})
	}
}

func TestApplyLayoutMutations_RefusesAmbiguousColumnMarkup(t *testing.T) {
	tests := map[string]string{
		"overlap":          `<cols><col min="1" max="3"/><col min="3" max="4"/></cols>`,
		"out of order":     `<cols><col min="5" max="6"/><col min="1" max="2"/></cols>`,
		"reversed":         `<cols><col min="4" max="2"/></cols>`,
		"missing min":      `<cols><col max="2"/></cols>`,
		"noncanonical min": `<cols><col min="01" max="2"/></cols>`,
		"nested markup":    `<cols><col min="1" max="2"><ext/></col></cols>`,
		"foreign child":    `<cols xmlns:f="urn:foreign"><f:col min="1" max="2"/></cols>`,
		"text content":     `<cols>unsafe<col min="1" max="2"/></cols>`,
	}
	for name, cols := range tests {
		t.Run(name, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["xl/worksheets/sheet1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `">` + cols + `<sheetData/></worksheet>`
			op := layoutMutation("unsafe-cols", "7", ColumnSetWidth)
			op.Width = 12
			output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{op})
			if err == nil {
				t.Fatalf("expected refusal, output=%d", len(output))
			}
			if output != nil {
				t.Fatal("unsafe column markup returned partial output")
			}
		})
	}

	for _, worksheet := range []string{
		`<worksheet xmlns="` + spreadsheetMLTransitional + `"><cols/><cols/><sheetData/></worksheet>`,
		`<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData/><cols/></worksheet>`,
	} {
		entries := cellMutationFixture()
		entries["xl/worksheets/sheet1.xml"] = worksheet
		op := layoutMutation("unsafe-cols-position", "7", ColumnSetWidth)
		op.Width = 12
		if output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{op}); err == nil || output != nil {
			t.Fatalf("expected cols position/count refusal, output=%d err=%v", len(output), err)
		}
	}
}

func TestApplyLayoutMutations_RefusesUnsafeRowsAndIsAtomicAcrossSheets(t *testing.T) {
	for name, sheet := range map[string]string{
		"duplicate":     `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData><row r="1"/><row r="1"/></sheetData></worksheet>`,
		"unordered":     `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData><row r="2"/><row r="1"/></sheetData></worksheet>`,
		"foreign child": `<worksheet xmlns="` + spreadsheetMLTransitional + `" xmlns:f="urn:foreign"><sheetData><f:row/></sheetData></worksheet>`,
	} {
		t.Run(name, func(t *testing.T) {
			entries := cellMutationFixture()
			entries["xl/worksheets/sheet1.xml"] = sheet
			first := layoutMutation("valid-other", "9", RowSetHeight)
			first.HeightPoints = 20
			second := layoutMutation("invalid-data", "7", RowSetHeight)
			second.HeightPoints = 20
			output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{first, second})
			if err == nil || output != nil {
				t.Fatalf("expected atomic refusal, output=%d err=%v", len(output), err)
			}
		})
	}
}

func TestApplyLayoutMutations_RoutesRelocatedCaseVariantWorkbookParts(t *testing.T) {
	entries := cellMutationFixture()
	workbook := entries["xl/workbook.xml"]
	workbookRels := entries["xl/_rels/workbook.xml.rels"]
	sheet := entries["xl/worksheets/sheet1.xml"]
	delete(entries, "xl/workbook.xml")
	delete(entries, "xl/_rels/workbook.xml.rels")
	delete(entries, "xl/worksheets/sheet1.xml")
	entries["Custom/Office/Book.XML"] = workbook
	entries["Custom/Office/_rels/Book.XML.rels"] = strings.Replace(workbookRels, `Target="worksheets/sheet1.xml"`, `Target="../../CUSTOM/SHEETS/DATA.XML"`, 1)
	entries["Custom/Sheets/Data.XML"] = sheet
	entries["_rels/.rels"] = `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `" Target="/custom/office/book.xml"/></Relationships>`
	op := layoutMutation("relocated", "7", RowSetHeight)
	op.HeightPoints = 21

	output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if relocated := readEntry(t, output, "Custom/Sheets/Data.XML"); !strings.Contains(relocated, `ht="21"`) {
		t.Fatalf("relocated worksheet was not updated: %s", relocated)
	}
	if !hasZipEntry(t, output, "Custom/Sheets/Data.XML") || hasZipEntry(t, output, "custom/sheets/data.xml") {
		t.Fatal("original OPC part spelling was not preserved")
	}
}

func TestApplyLayoutMutations_RefusesCaseEquivalentPackageParts(t *testing.T) {
	for name, duplicate := range map[string]string{
		"case":   "XL/WORKBOOK.XML",
		"escape": "xl/%77orkbook.xml",
	} {
		t.Run(name, func(t *testing.T) {
			entries := cellMutationFixture()
			entries[duplicate] = entries["xl/workbook.xml"]
			op := layoutMutation("ambiguous-part", "7", RowSetHeight)
			op.HeightPoints = 20
			output, err := ApplyLayoutMutations(buildZip(t, entries), []LayoutMutation{op})
			if err == nil || !strings.Contains(err.Error(), "case/escape-equivalent OPC parts") {
				t.Fatalf("expected equivalent part refusal, output=%d err=%v", len(output), err)
			}
			if output != nil {
				t.Fatal("ambiguous package returned partial output")
			}
		})
	}
}
