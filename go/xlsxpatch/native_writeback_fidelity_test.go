package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// Native write-back fidelity is the cell-value save product: ApplyCellMutations
// must keep unmodeled neighbors (style index, sibling cells, worksheet markup,
// opaque OPC parts, unrelated shared-string items), emit deterministic bytes,
// and still refuse unknown target-cell markup.

func TestNativeWritebackFidelity_StyleIndexSurvivesSetValueAndSiblingCellsStayByteIdentical(t *testing.T) {
	entries := cellMutationFixture()
	siblingBoolean := `<x:c r="B1" s="9" t="b"><x:v>1</x:v></x:c>`
	siblingFormula := `<x:c r="D1" s="3"><x:f>OLD()</x:f><x:v>99</x:v></x:c>`
	siblingInline := `<x:c r="B3" s="2" t="inlineStr"><x:is><x:t>old</x:t></x:is></x:c>`
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`<x:c r="A1" s="7" t="s"><x:v>0</x:v><x:extLst><x:ext uri="keep"><x:payload/></x:ext></x:extLst></x:c><x:c r="D1" s="3"><x:f>OLD()</x:f><x:v>99</x:v></x:c>`,
		`<x:c r="A1" s="7" t="s"><x:v>0</x:v><x:extLst><x:ext uri="keep"><x:payload/></x:ext></x:extLst></x:c>`+
			siblingBoolean+
			`<x:c r="C1" s="11"><x:v>3.5</x:v></x:c>`+
			siblingFormula,
		1,
	)

	shared := mutation("fidelity-shared-string", "7", CellSetValue, 0, 0)
	shared.Value = "writeback-shared"
	numeric := mutation("fidelity-numeric", "7", CellSetValue, 0, 2)
	numeric.Value = 8

	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{shared, numeric})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<x:c r="A1" s="7" t="inlineStr"><x:is><x:t>writeback-shared</x:t></x:is><x:extLst><x:ext uri="keep"><x:payload/></x:ext></x:extLst></x:c>`,
		`<x:c r="C1" s="11"><x:v>8</x:v></x:c>`,
		siblingBoolean,
		siblingFormula,
		siblingInline,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet is missing %q\n%s", want, sheet)
		}
	}
}

func TestNativeWritebackFidelity_UnrelatedWorksheetMarkupSurvivesInRangeSetValue(t *testing.T) {
	// In-range set_value must not rewrite sheetFormatPr, cols, autoFilter, or
	// unrelated mergeCells. Dimension may expand when a write falls outside
	// the stored ref (locked elsewhere); this case stays inside A1:D3 so the
	// original dimension start tag is preserved. XML comments outside the
	// rebuilt row are copied with the surrounding worksheet bytes.
	entries := cellMutationFixture()
	sheetFormat := `<x:sheetFormatPr defaultRowHeight="14.4" defaultColWidth="10.6640625" customHeight="0"/>`
	cols := `<x:cols><x:col min="1" max="2" width="12.25" customWidth="1"/><x:col min="4" max="4" width="18" style="3"/></x:cols>`
	comment := `<!--writeback-fidelity-sheet-->`
	autoFilter := `<x:autoFilter ref="A1:D3"/>`
	merges := `<x:mergeCells count="1"><x:mergeCell ref="C3:D3"/></x:mergeCells>`
	dimension := `<x:dimension ref="A1:D3"/>`
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		dimension+`<x:sheetData>`,
		dimension+sheetFormat+cols+comment+`<x:sheetData>`,
		1,
	)
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`</x:sheetData><x:extLst>`,
		`</x:sheetData>`+autoFilter+merges+`<x:extLst>`,
		1,
	)

	op := mutation("fidelity-in-range", "7", CellSetValue, 0, 0)
	op.Value = "in-range"
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{dimension, sheetFormat, cols, comment, autoFilter, merges} {
		if !strings.Contains(sheet, want) {
			t.Errorf("sheet is missing %q\n%s", want, sheet)
		}
	}
}

func TestNativeWritebackFidelity_OpaqueExtraPartStaysRawIdenticalAfterCellValueChange(t *testing.T) {
	entries := cellMutationFixture()
	const part = "customXml/item2.xml"
	entries[part] = `<writebackFidelity xmlns="urn:injoffice:writeback"><opaque>neighbor-part</opaque></writebackFidelity>`

	op := mutation("fidelity-opaque", "7", CellSetValue, 0, 0)
	op.Value = "changed-cell"
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if got := readEntry(t, output, part); got != entries[part] {
		t.Errorf("opaque part %q changed\n got: %s\nwant: %s", part, got, entries[part])
	}
}

func TestNativeWritebackFidelity_SetValueDoesNotRewriteUnrelatedSharedStringItems(t *testing.T) {
	entries := cellMutationFixture()
	sst := `<sst xmlns="` + spreadsheetMLTransitional + `" count="4" uniqueCount="3">` +
		`<si><t>keep-unrelated-alpha</t></si>` +
		`<si><t>keep-unrelated-omega</t></si>` +
		`<si><t xml:space="preserve">  keep-spaces  </t></si>` +
		`</sst>`
	entries["xl/sharedStrings.xml"] = sst
	entries["xl/_rels/workbook.xml.rels"] = strings.Replace(
		entries["xl/_rels/workbook.xml.rels"],
		`</Relationships>`,
		`<Relationship Id="rIdShared" Type="`+relTypeSharedStringsTransitional+`" Target="sharedStrings.xml"/></Relationships>`,
		1,
	)
	sharedNeighbor := `<x:c r="C3" s="5" t="s"><x:v>1</x:v></x:c>`
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`<x:c r="B3" s="2" t="inlineStr"><x:is><x:t>old</x:t></x:is></x:c>`,
		`<x:c r="B3" s="2" t="inlineStr"><x:is><x:t>old</x:t></x:is></x:c>`+sharedNeighbor,
		1,
	)

	op := mutation("fidelity-sst", "7", CellSetValue, 0, 0)
	op.Value = "new-shared-owner"
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if got := readEntry(t, output, "xl/sharedStrings.xml"); got != sst {
		t.Errorf("shared strings part was rewritten\n got: %s\nwant: %s", got, sst)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, sharedNeighbor) {
		t.Errorf("unrelated shared-string cell was rewritten: %s", sheet)
	}
	for _, want := range []string{
		`<si><t>keep-unrelated-alpha</t></si>`,
		`<si><t>keep-unrelated-omega</t></si>`,
		`<si><t xml:space="preserve">  keep-spaces  </t></si>`,
	} {
		if !strings.Contains(readEntry(t, output, "xl/sharedStrings.xml"), want) {
			t.Errorf("shared strings lost %q", want)
		}
	}
}

func TestNativeWritebackFidelity_SameMutationsFromSameSourceAreByteIdentical(t *testing.T) {
	original := buildZip(t, cellMutationFixture())
	op := mutation("fidelity-deterministic", "7", CellSetValue, 0, 0)
	op.Value = "same-twice"
	first, err := ApplyCellMutations(original, []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	second, err := ApplyCellMutations(original, []CellMutation{op})
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, second) {
		t.Fatalf("same mutations from the same source produced %d vs %d distinct bytes", len(first), len(second))
	}
}

func TestNativeWritebackFidelity_UnknownTargetCellMarkupStillRefuses(t *testing.T) {
	entries := cellMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(
		entries["xl/worksheets/sheet1.xml"],
		`<x:v>0</x:v><x:extLst>`,
		`<x:v>0</x:v><x:unsupported/><x:extLst>`,
		1,
	)
	op := mutation("fidelity-refuse", "7", CellSetValue, 0, 0)
	op.Value = "would-drop-unknown-markup"
	output, err := ApplyCellMutations(buildZip(t, entries), []CellMutation{op})
	if err == nil || !strings.Contains(err.Error(), "unsupported direct child") {
		t.Fatalf("expected unsupported-markup refusal, output=%d err=%v", len(output), err)
	}
	if output != nil {
		t.Fatal("refused edit returned partial output")
	}
}
