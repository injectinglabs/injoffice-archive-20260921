package xlsxpatch

import (
	"bytes"
	"math"
	"strings"
	"testing"
)

func styleMutationFixture() map[string]string {
	entries := cellMutationFixture()
	entries["[Content_Types].xml"] = `<Types xmlns="` + contentTypesNamespace + `"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/styles.xml" ContentType="` + stylesPartContentType + `"/></Types>`
	entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], `</Relationships>`, `<Relationship Id="rIdStyles" Type="`+relTypeStylesTransitional+`" Target="styles.xml"/></Relationships>`, 1)
	entries["xl/styles.xml"] = `<?xml version="1.0"?><x:styleSheet xmlns:x="` + spreadsheetMLTransitional + `">` +
		`<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="$#,##0.00"/></x:numFmts>` +
		`<x:fonts count="2"><x:font><x:name val="Calibri"/><x:sz val="11"/><x:color rgb="FF000000"/></x:font><x:font><x:name val="Arial"/><x:sz val="12"/><x:b/><x:color rgb="FF112233"/></x:font></x:fonts>` +
		`<x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="gray125"/></x:fill></x:fills>` +
		`<x:borders count="1"><x:border/></x:borders>` +
		`<x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"><x:alignment vertical="bottom"/></x:xf></x:cellStyleXfs>` +
		`<x:cellXfs count="2"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><x:xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" wrapText="0"/></x:xf></x:cellXfs>` +
		`<x:cellStyles count="1"><x:cellStyle name="Normal" xfId="0" builtinId="0"/></x:cellStyles>` +
		`</x:styleSheet>`
	entries["xl/worksheets/sheet1.xml"] = `<?xml version="1.0"?><x:worksheet xmlns:x="` + spreadsheetMLTransitional + `"><x:dimension ref="A1:C1"/><x:sheetData><x:row r="1" spans="1:3"><x:c r="A1" s="1"><x:f>SUM(C1,1)</x:f><x:v>9</x:v><x:extLst><x:ext uri="keep"><x:opaque/></x:ext></x:extLst></x:c><x:c r="C1" s="1" t="inlineStr"><x:is><x:t>shared-style-untouched</x:t></x:is></x:c></x:row></x:sheetData><x:extLst><x:ext uri="sheet-keep"/></x:extLst></x:worksheet>`
	return entries
}

func fullStyleDelta() StyleDelta {
	return StyleDelta{
		NumberFormat: SetStyleProperty("0.00%"),
		FontName:     SetStyleProperty("Aptos"), FontSizePoints: SetStyleProperty(13.0),
		Bold: SetStyleProperty(false), Italic: SetStyleProperty(true), FontColor: SetStyleProperty("#AABBCC"),
		FillColor: SetStyleProperty("#DDEEFF"), HorizontalAlignment: SetStyleProperty(HorizontalCenter),
		VerticalAlignment: SetStyleProperty(VerticalMiddle), WrapText: SetStyleProperty(true),
	}
}

func TestApplyStyleMutations_PatchesBoundedRangeAndPreservesCellBodies(t *testing.T) {
	entries := styleMutationFixture()
	originalStyles := entries["xl/styles.xml"]
	mutation := StylePatchMutation{
		OperationID: "style-range", SheetID: "7", Kind: StylePatch,
		Range: StyleRange{Row: 0, Column: 0, EndRow: 1, EndColumn: 1}, Style: fullStyleDelta(),
	}
	output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	for _, want := range []string{
		`<x:dimension ref="A1:C2"/>`,
		`<x:row r="1">`,
		`<x:c r="A1" s="2"><x:f>SUM(C1,1)</x:f><x:v>9</x:v><x:extLst><x:ext uri="keep"><x:opaque/></x:ext></x:extLst></x:c>`,
		`<x:c r="B1" s="2"/>`,
		`<x:c r="C1" s="1" t="inlineStr"><x:is><x:t>shared-style-untouched</x:t></x:is></x:c>`,
		`<x:row r="2"><x:c r="A2" s="2"/><x:c r="B2" s="2"/></x:row>`,
		`<x:extLst><x:ext uri="sheet-keep"/></x:extLst>`,
	} {
		if !strings.Contains(sheet, want) {
			t.Errorf("worksheet missing %q\n%s", want, sheet)
		}
	}
	styles := readEntry(t, output, "xl/styles.xml")
	for _, want := range []string{
		`<x:numFmts count="1">`,
		`<x:fonts count="3">`, `<x:font><x:name val="Aptos"/><x:i/><x:sz val="13"/><x:color rgb="FFAABBCC"/></x:font>`,
		`<x:fills count="3">`, `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FFDDEEFF"/></x:patternFill></x:fill>`,
		`<x:cellXfs count="3">`, `<x:alignment horizontal="center" vertical="center" wrapText="1"/>`,
	} {
		if !strings.Contains(styles, want) {
			t.Errorf("styles missing %q\n%s", want, styles)
		}
	}
	for _, originalRecord := range []string{
		`<x:font><x:name val="Arial"/><x:sz val="12"/><x:b/><x:color rgb="FF112233"/></x:font>`,
		`<x:xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyAlignment="1"><x:alignment horizontal="left" wrapText="0"/></x:xf>`,
	} {
		if !strings.Contains(originalStyles, originalRecord) || !strings.Contains(styles, originalRecord) {
			t.Errorf("shared source style record changed: %s", originalRecord)
		}
	}
	for _, part := range []string{"_rels/.rels", "xl/workbook.xml", "xl/charts/chart1.xml", "customXml/item1.xml"} {
		if got := readEntry(t, output, part); got != entries[part] {
			t.Errorf("unrelated part %q changed", part)
		}
	}
}

func TestApplyStyleMutations_ClearsDirectPropertiesToInheritedStyle(t *testing.T) {
	entries := styleMutationFixture()
	originalStyles := entries["xl/styles.xml"]
	delta := StyleDelta{
		NumberFormat: ClearStyleProperty[string](), FontName: ClearStyleProperty[string](), FontSizePoints: ClearStyleProperty[float64](),
		Bold: ClearStyleProperty[bool](), Italic: ClearStyleProperty[bool](), FontColor: ClearStyleProperty[string](),
		FillColor: ClearStyleProperty[string](), HorizontalAlignment: ClearStyleProperty[HorizontalAlignment](),
		VerticalAlignment: ClearStyleProperty[VerticalAlignment](), WrapText: ClearStyleProperty[bool](),
	}
	mutation := StylePatchMutation{OperationID: "clear-style", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: delta}
	output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, `<x:c r="A1"><x:f>SUM(C1,1)</x:f><x:v>9</x:v>`) {
		t.Fatalf("clear did not inherit cellStyleXf index 0 while preserving content: %s", sheet)
	}
	if styles := readEntry(t, output, "xl/styles.xml"); styles != originalStyles {
		t.Fatal("clearing to an existing inherited style rewrote the shared style table")
	}
}

func TestApplyStyleMutations_SemanticNoopReturnsOriginalArchiveBytes(t *testing.T) {
	entries := styleMutationFixture()
	original := buildZip(t, entries)
	delta := StyleDelta{
		NumberFormat: SetStyleProperty("$#,##0.00"), FontName: SetStyleProperty("Arial"), FontSizePoints: SetStyleProperty(12.0),
		Bold: SetStyleProperty(true), Italic: SetStyleProperty(false), FontColor: SetStyleProperty("#112233"),
		FillColor: ClearStyleProperty[string](), HorizontalAlignment: SetStyleProperty(HorizontalLeft),
		VerticalAlignment: SetStyleProperty(VerticalBottom), WrapText: SetStyleProperty(false),
	}
	first := StylePatchMutation{OperationID: "same-first", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}
	last := StylePatchMutation{OperationID: "same-last", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: delta}
	output, err := ApplyStyleMutations(original, []StylePatchMutation{first, last})
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(output, original) {
		t.Fatalf("last-write semantic no-op did not return the original XLSX bytes\nstyles: %s\nsheet: %s", readEntry(t, output, "xl/styles.xml"), readEntry(t, output, "xl/worksheets/sheet1.xml"))
	}
}

func TestApplyStyleMutations_RefusesInvalidBatchAndRangeBudget(t *testing.T) {
	valid := StylePatchMutation{OperationID: "valid", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(true)}}
	cases := map[string][]StylePatchMutation{
		"empty":                  {},
		"duplicate operation id": {valid, valid},
		"wrong kind":             {func() StylePatchMutation { op := valid; op.Kind = "cell.set_value"; return op }()},
		"reversed range":         {func() StylePatchMutation { op := valid; op.Range.EndRow = -1; return op }()},
		"range budget":           {func() StylePatchMutation { op := valid; op.Range.EndRow = maxStyleRangeCells; return op }()},
		"empty delta":            {func() StylePatchMutation { op := valid; op.Style = StyleDelta{}; return op }()},
		"lowercase color": {func() StylePatchMutation {
			op := valid
			op.Style = StyleDelta{FontColor: SetStyleProperty("#aabbcc")}
			return op
		}()},
		"infinite font size": {func() StylePatchMutation {
			op := valid
			op.Style = StyleDelta{FontSizePoints: SetStyleProperty(math.Inf(1))}
			return op
		}()},
		"value without presence": {func() StylePatchMutation {
			op := valid
			value := true
			op.Style = StyleDelta{Bold: StyleProperty[bool]{Value: &value}}
			return op
		}()},
		"forbidden XML": {func() StylePatchMutation {
			op := valid
			op.Style = StyleDelta{FontName: SetStyleProperty("bad\x01font")}
			return op
		}()},
	}
	original := buildZip(t, styleMutationFixture())
	for name, operations := range cases {
		t.Run(name, func(t *testing.T) {
			output, err := ApplyStyleMutations(original, operations)
			if err == nil || output != nil {
				t.Fatalf("expected atomic refusal, output=%d err=%v", len(output), err)
			}
		})
	}
}

func TestApplyStyleMutations_RoutesRelocatedStrictPartsWithDefaultContentType(t *testing.T) {
	entries := styleMutationFixture()
	workbook := strings.ReplaceAll(entries["xl/workbook.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	workbook = strings.ReplaceAll(workbook, officeRelNamespaceTransitional, officeRelNamespaceStrict)
	rels := strings.ReplaceAll(entries["xl/_rels/workbook.xml.rels"], relTypeWorksheetTransitional, relTypeWorksheetStrict)
	rels = strings.ReplaceAll(rels, relTypeStylesTransitional, relTypeStylesStrict)
	rels = strings.Replace(rels, `Target="worksheets/sheet1.xml"`, `Target="../../CUSTOM/SHEETS/DATA.XML"`, 1)
	rels = strings.Replace(rels, `Target="styles.xml"`, `Target="../Style/Table.STY"`, 1)
	styles := strings.ReplaceAll(entries["xl/styles.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	sheet := strings.ReplaceAll(entries["xl/worksheets/sheet1.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
	delete(entries, "xl/workbook.xml")
	delete(entries, "xl/_rels/workbook.xml.rels")
	delete(entries, "xl/styles.xml")
	delete(entries, "xl/worksheets/sheet1.xml")
	entries["Custom/Office/Book.XML"] = workbook
	entries["Custom/Office/_rels/Book.XML.rels"] = rels
	entries["Custom/Style/Table.STY"] = styles
	entries["Custom/Sheets/Data.XML"] = sheet
	entries["_rels/.rels"] = `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentStrict + `" Target="/custom/office/book.xml"/></Relationships>`
	entries["[Content_Types].xml"] = `<Types xmlns="` + contentTypesNamespace + `"><Default Extension="xml" ContentType="application/xml"/><Default Extension="sty" ContentType="` + stylesPartContentType + `"/></Types>`
	mutation := StylePatchMutation{OperationID: "strict-relocated", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}

	output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
	if err != nil {
		t.Fatal(err)
	}
	if relocated := readEntry(t, output, "Custom/Sheets/Data.XML"); !strings.Contains(relocated, `r="A1" s="2"`) {
		t.Fatalf("relocated Strict worksheet was not patched: %s", relocated)
	}
	if !hasZipEntry(t, output, "Custom/Style/Table.STY") || hasZipEntry(t, output, "custom/style/table.sty") {
		t.Fatal("original relocated styles-part spelling was not preserved")
	}
}

func TestApplyStyleMutations_RefusesAmbiguousOrWrongStylesContentTypes(t *testing.T) {
	base := styleMutationFixture()
	tests := map[string]string{
		"wrong effective type": `<Types xmlns="` + contentTypesNamespace + `"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/styles.xml" ContentType="application/xml"/></Types>`,
		"duplicate override":   `<Types xmlns="` + contentTypesNamespace + `"><Override PartName="/xl/styles.xml" ContentType="` + stylesPartContentType + `"/><Override PartName="/XL/STYLES.XML" ContentType="` + stylesPartContentType + `"/></Types>`,
		"duplicate default":    `<Types xmlns="` + contentTypesNamespace + `"><Default Extension="xml" ContentType="application/xml"/><Default Extension="XML" ContentType="` + stylesPartContentType + `"/></Types>`,
		"spoofed namespace":    `<Types xmlns="urn:spoof"><Override PartName="/xl/styles.xml" ContentType="` + stylesPartContentType + `"/></Types>`,
		"multiple roots":       `<Types xmlns="` + contentTypesNamespace + `"><Override PartName="/xl/styles.xml" ContentType="` + stylesPartContentType + `"/></Types><Types xmlns="` + contentTypesNamespace + `"></Types>`,
	}
	mutation := StylePatchMutation{OperationID: "bad-content-type", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}
	for name, contentTypes := range tests {
		t.Run(name, func(t *testing.T) {
			entries := make(map[string]string, len(base))
			for key, value := range base {
				entries[key] = value
			}
			entries["[Content_Types].xml"] = contentTypes
			output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
			if err == nil || output != nil {
				t.Fatalf("expected content-type refusal, output=%d err=%v", len(output), err)
			}
		})
	}
}

func TestApplyStyleMutations_PreservesExcelThemeFontChildrenDuringTargetedPatch(t *testing.T) {
	themeFont := `<x:font><x:name val="Aptos"/><x:family val="2"/><x:color theme="1"/><x:sz val="11"/><x:scheme val="minor"/></x:font>`
	entries := styleMutationFixture()
	entries["xl/styles.xml"] = strings.Replace(entries["xl/styles.xml"], `<x:font><x:name val="Arial"/><x:sz val="12"/><x:b/><x:color rgb="FF112233"/></x:font>`, themeFont, 1)
	original := buildZip(t, entries)

	t.Run("semantic noop", func(t *testing.T) {
		mutation := StylePatchMutation{OperationID: "theme-noop", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}
		output, err := ApplyStyleMutations(original, []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(output, original) {
			t.Fatal("bold=false on the theme-colored font was not a byte-preserving semantic no-op")
		}
	})

	t.Run("bold preserves theme color and schema position", func(t *testing.T) {
		mutation := StylePatchMutation{OperationID: "theme-bold", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(true)}}
		output, err := ApplyStyleMutations(original, []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		styles := readEntry(t, output, "xl/styles.xml")
		want := `<x:font><x:name val="Aptos"/><x:family val="2"/><x:b/><x:color theme="1"/><x:sz val="11"/><x:scheme val="minor"/></x:font>`
		if !strings.Contains(styles, want) {
			t.Fatalf("targeted bold patch did not preserve and order typical font children:\n%s", styles)
		}
		if !strings.Contains(styles, themeFont) {
			t.Fatal("shared source theme font was mutated")
		}
	})

	t.Run("explicit color safely replaces theme color", func(t *testing.T) {
		mutation := StylePatchMutation{OperationID: "theme-color", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{FontColor: SetStyleProperty("#ABCDEF")}}
		output, err := ApplyStyleMutations(original, []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		styles := readEntry(t, output, "xl/styles.xml")
		want := `<x:font><x:name val="Aptos"/><x:family val="2"/><x:color rgb="FFABCDEF"/><x:sz val="11"/><x:scheme val="minor"/></x:font>`
		if !strings.Contains(styles, want) || !strings.Contains(styles, themeFont) {
			t.Fatalf("explicit font color replacement did not preserve other/shared font children:\n%s", styles)
		}
	})
}

func TestApplyStyleMutations_AcceptsCanonicalSolidFillBackgroundFallback(t *testing.T) {
	entries := styleMutationFixture()
	styles := entries["xl/styles.xml"]
	styles = strings.Replace(styles, `<x:fills count="2">`, `<x:fills count="3">`, 1)
	styles = strings.Replace(styles, `</x:fills>`, `<x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FF112233"/><x:bgColor indexed="64"/></x:patternFill></x:fill></x:fills>`, 1)
	styles = strings.Replace(styles, `fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"`, `fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"`, 1)
	entries["xl/styles.xml"] = styles
	original := buildZip(t, entries)

	noop := StylePatchMutation{OperationID: "fill-noop", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{FillColor: SetStyleProperty("#112233")}}
	output, err := ApplyStyleMutations(original, []StylePatchMutation{noop})
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(output, original) {
		t.Fatal("matching canonical solid fill was not a byte-preserving no-op")
	}

	patch := StylePatchMutation{OperationID: "fill-replace", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{FillColor: SetStyleProperty("#AABBCC")}}
	output, err = ApplyStyleMutations(original, []StylePatchMutation{patch})
	if err != nil {
		t.Fatal(err)
	}
	updated := readEntry(t, output, "xl/styles.xml")
	if !strings.Contains(updated, `<x:patternFill patternType="solid"><x:fgColor rgb="FFAABBCC"/></x:patternFill>`) ||
		!strings.Contains(updated, `<x:fgColor rgb="FF112233"/><x:bgColor indexed="64"/>`) {
		t.Fatalf("solid fill patch did not preserve the shared canonical indexed-64 background fallback:\n%s", updated)
	}
}

func TestApplyStyleMutations_ExplicitSolidFillReplacesThemeFill(t *testing.T) {
	entries := styleMutationFixture()
	styles := entries["xl/styles.xml"]
	styles = strings.Replace(styles, `<x:fills count="2">`, `<x:fills count="3">`, 1)
	styles = strings.Replace(styles, `</x:fills>`, `<x:fill><x:patternFill patternType="solid"><x:fgColor theme="4" tint="0.4"/><x:bgColor indexed="64"/></x:patternFill></x:fill></x:fills>`, 1)
	styles = strings.Replace(styles, `fontId="1" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"`, `fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFill="1"`, 1)
	entries["xl/styles.xml"] = styles
	original := buildZip(t, entries)

	set := StylePatchMutation{OperationID: "theme-fill-replace", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{FillColor: SetStyleProperty("#AABBCC")}}
	output, err := ApplyStyleMutations(original, []StylePatchMutation{set})
	if err != nil {
		t.Fatal(err)
	}
	updated := readEntry(t, output, "xl/styles.xml")
	if !strings.Contains(updated, `<x:patternFill patternType="solid"><x:fgColor rgb="FFAABBCC"/></x:patternFill>`) ||
		!strings.Contains(updated, `<x:fgColor theme="4" tint="0.4"/><x:bgColor indexed="64"/>`) {
		t.Fatalf("explicit RGB fill did not safely replace the target while preserving the shared theme fill:\n%s", updated)
	}

	clear := StylePatchMutation{OperationID: "theme-fill-clear", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{FillColor: ClearStyleProperty[string]()}}
	output, err = ApplyStyleMutations(original, []StylePatchMutation{clear})
	if err == nil || output != nil {
		t.Fatalf("ambiguous theme-fill clear should fail atomically, output=%d err=%v", len(output), err)
	}
}

func TestApplyStyleMutations_AppendsAndDeduplicatesCustomNumberFormat(t *testing.T) {
	entries := styleMutationFixture()
	operations := []StylePatchMutation{
		{OperationID: "format-a1", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{NumberFormat: SetStyleProperty("yyyy-mm-dd hh:mm")}},
		{OperationID: "format-c1", SheetID: "7", Kind: StylePatch, Range: StyleRange{Column: 2, EndColumn: 2}, Style: StyleDelta{NumberFormat: SetStyleProperty("yyyy-mm-dd hh:mm")}},
	}
	output, err := ApplyStyleMutations(buildZip(t, entries), operations)
	if err != nil {
		t.Fatal(err)
	}
	styles := readEntry(t, output, "xl/styles.xml")
	if strings.Count(styles, `formatCode="yyyy-mm-dd hh:mm"`) != 1 || !strings.Contains(styles, `<x:numFmts count="2">`) {
		t.Fatalf("custom number format was not appended exactly once:\n%s", styles)
	}
	if !strings.Contains(styles, `<x:numFmt numFmtId="165" formatCode="yyyy-mm-dd hh:mm"/>`) {
		t.Fatalf("custom number format did not receive the deterministic lowest free id:\n%s", styles)
	}
	sheet := readEntry(t, output, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, `<x:c r="A1" s="2">`) || !strings.Contains(sheet, `<x:c r="C1" s="2"`) {
		t.Fatalf("cells sharing the same derived format did not reuse one cellXf:\n%s", sheet)
	}
}

func TestApplyStyleMutations_RefusesTargetedUnsafeCellAtomically(t *testing.T) {
	entries := styleMutationFixture()
	entries["xl/worksheets/sheet1.xml"] = strings.Replace(entries["xl/worksheets/sheet1.xml"], `<x:c r="C1"`, `<x:c r="B1" s="999"/><x:c r="C1"`, 1)
	original := buildZip(t, entries)
	operations := []StylePatchMutation{
		{OperationID: "valid-first", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}},
		{OperationID: "unsafe-second", SheetID: "7", Kind: StylePatch, Range: StyleRange{Column: 1, EndColumn: 1}, Style: StyleDelta{Bold: SetStyleProperty(true)}},
	}
	output, err := ApplyStyleMutations(original, operations)
	if err == nil || output != nil {
		t.Fatalf("targeted out-of-range style should refuse the entire ordered batch, output=%d err=%v", len(output), err)
	}
}

func TestApplyStyleMutations_RefusesHostileStylesAndWorksheetXML(t *testing.T) {
	base := styleMutationFixture()
	mutation := StylePatchMutation{OperationID: "hostile-xml", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}
	tests := map[string]func(map[string]string){
		"multiple styles roots": func(entries map[string]string) {
			entries["xl/styles.xml"] += `<x:styleSheet xmlns:x="` + spreadsheetMLTransitional + `"></x:styleSheet>`
		},
		"styles processing instruction": func(entries map[string]string) {
			entries["xl/styles.xml"] = strings.Replace(entries["xl/styles.xml"], `<x:fonts`, `<?unsafe value?><x:fonts`, 1)
		},
		"styles directive": func(entries map[string]string) {
			entries["xl/styles.xml"] = strings.Replace(entries["xl/styles.xml"], `<x:fonts`, `<!ENTITY unsafe "value"><x:fonts`, 1)
		},
		"styles record text": func(entries map[string]string) {
			entries["xl/styles.xml"] = strings.Replace(entries["xl/styles.xml"], `<x:patternFill patternType="none"/>`, `<x:patternFill patternType="none">unsafe</x:patternFill>`, 1)
		},
		"cellXf direct text": func(entries map[string]string) {
			entries["xl/styles.xml"] = strings.Replace(entries["xl/styles.xml"], `<x:alignment horizontal="left" wrapText="0"/>`, `<x:alignment horizontal="left" wrapText="0"/>unsafe`, 1)
		},
		"worksheet processing instruction": func(entries map[string]string) {
			entries["xl/worksheets/sheet1.xml"] = strings.Replace(entries["xl/worksheets/sheet1.xml"], `<x:sheetData>`, `<x:sheetData><?unsafe value?>`, 1)
		},
		"worksheet directive": func(entries map[string]string) {
			entries["xl/worksheets/sheet1.xml"] = strings.Replace(entries["xl/worksheets/sheet1.xml"], `<x:sheetData>`, `<x:sheetData><!ENTITY unsafe "value">`, 1)
		},
		"multiple worksheet roots": func(entries map[string]string) {
			entries["xl/worksheets/sheet1.xml"] += `<x:worksheet xmlns:x="` + spreadsheetMLTransitional + `"><x:sheetData/></x:worksheet>`
		},
	}
	for name, mutate := range tests {
		t.Run(name, func(t *testing.T) {
			entries := make(map[string]string, len(base))
			for key, value := range base {
				entries[key] = value
			}
			mutate(entries)
			output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
			if err == nil || output != nil {
				t.Fatalf("hostile XML should fail atomically, output=%d err=%v", len(output), err)
			}
		})
	}
}

func TestApplyStyleMutations_AlignmentComponentDoesNotPartiallyInherit(t *testing.T) {
	t.Run("direct missing vertical uses default not parent", func(t *testing.T) {
		entries := styleMutationFixture()
		styles := entries["xl/styles.xml"]
		styles = strings.Replace(styles, `<x:alignment vertical="bottom"/>`, `<x:alignment vertical="top"/>`, 1)
		styles = strings.Replace(styles, `<x:alignment horizontal="left" wrapText="0"/>`, `<x:alignment horizontal="left" wrapText="0" textRotation="15"/>`, 1)
		entries["xl/styles.xml"] = styles
		mutation := StylePatchMutation{OperationID: "set-parent-value", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{VerticalAlignment: SetStyleProperty(VerticalTop)}}

		output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		updatedStyles := readEntry(t, output, "xl/styles.xml")
		if !strings.Contains(updatedStyles, `<x:alignment textRotation="15" horizontal="left" vertical="top" wrapText="0"/>`) {
			t.Fatalf("vertical=top was incorrectly treated as inherited/no-op or lost opaque alignment attributes:\n%s", updatedStyles)
		}
		if !strings.Contains(readEntry(t, output, "xl/worksheets/sheet1.xml"), `<x:c r="A1" s="2">`) {
			t.Fatal("direct vertical default was not changed to top")
		}
	})

	t.Run("clear materializes parent property in retained direct component", func(t *testing.T) {
		entries := styleMutationFixture()
		styles := entries["xl/styles.xml"]
		styles = strings.Replace(styles, `<x:alignment vertical="bottom"/>`, `<x:alignment horizontal="left"/>`, 1)
		styles = strings.Replace(styles, `<x:alignment horizontal="left" wrapText="0"/>`, `<x:alignment horizontal="center" wrapText="0" textRotation="15"/>`, 1)
		entries["xl/styles.xml"] = styles
		mutation := StylePatchMutation{OperationID: "clear-horizontal", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{HorizontalAlignment: ClearStyleProperty[HorizontalAlignment]()}}

		output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		updatedStyles := readEntry(t, output, "xl/styles.xml")
		if !strings.Contains(updatedStyles, `<x:alignment textRotation="15" horizontal="left" wrapText="0"/>`) {
			t.Fatalf("clear removed the direct attribute instead of materializing parent left, or lost opaque attributes:\n%s", updatedStyles)
		}
	})
}

func TestApplyStyleMutations_PreservesXFBodyOrderAndOpaqueChildren(t *testing.T) {
	entries := styleMutationFixture()
	styles := strings.Replace(entries["xl/styles.xml"], `<x:styleSheet xmlns:x="`+spreadsheetMLTransitional+`">`, `<x:styleSheet xmlns:x="`+spreadsheetMLTransitional+`" xmlns:o="urn:opaque">`, 1)
	originalBody := `<x:protection locked="0"/><x:alignment horizontal="left" wrapText="0"/><o:opaque keep="yes"><o:nested>raw</o:nested></o:opaque><x:extLst><x:ext uri="xf-keep"/></x:extLst>`
	styles = strings.Replace(styles, `<x:alignment horizontal="left" wrapText="0"/>`, originalBody, 1)
	entries["xl/styles.xml"] = styles

	t.Run("non-alignment patch keeps body exact", func(t *testing.T) {
		mutation := StylePatchMutation{OperationID: "xf-font-only", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}
		output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		updated := readEntry(t, output, "xl/styles.xml")
		if strings.Count(updated, originalBody) != 2 {
			t.Fatalf("font-only patch did not preserve both source and cloned XF bodies byte-for-byte:\n%s", updated)
		}
	})

	t.Run("alignment patch replaces in place", func(t *testing.T) {
		mutation := StylePatchMutation{OperationID: "xf-alignment", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{VerticalAlignment: SetStyleProperty(VerticalTop)}}
		output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
		if err != nil {
			t.Fatal(err)
		}
		updated := readEntry(t, output, "xl/styles.xml")
		want := `<x:protection locked="0"/><x:alignment horizontal="left" vertical="top" wrapText="0"/><o:opaque keep="yes"><o:nested>raw</o:nested></o:opaque><x:extLst><x:ext uri="xf-keep"/></x:extLst>`
		if !strings.Contains(updated, want) {
			t.Fatalf("alignment was not replaced at its original position with opaque child order/raw bytes intact:\n%s", updated)
		}
	})
}

func TestApplyStyleMutations_EnforcesExactRoutingDialectAndMIME(t *testing.T) {
	mutation := StylePatchMutation{OperationID: "routing-exact", SheetID: "7", Kind: StylePatch, Range: StyleRange{}, Style: StyleDelta{Bold: SetStyleProperty(false)}}
	for _, relationship := range []string{"root", "styles"} {
		for _, mode := range []string{"internal", "INTERNAL", " Internal", "Internal ", "external", "External "} {
			t.Run(relationship+" target mode "+mode, func(t *testing.T) {
				entries := styleMutationFixture()
				if relationship == "root" {
					entries["_rels/.rels"] = strings.Replace(entries["_rels/.rels"], ` Target="xl/workbook.xml"`, ` Target="xl/workbook.xml" TargetMode="`+mode+`"`, 1)
				} else {
					entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], ` Target="styles.xml"`, ` Target="styles.xml" TargetMode="`+mode+`"`, 1)
				}
				output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
				if err == nil || output != nil {
					t.Fatalf("non-exact TargetMode should fail atomically, output=%d err=%v", len(output), err)
				}
			})
		}
	}

	for name, mutate := range map[string]func(map[string]string){
		"strict root with transitional workbook": func(entries map[string]string) {
			entries["_rels/.rels"] = strings.ReplaceAll(entries["_rels/.rels"], relTypeOfficeDocumentTransitional, relTypeOfficeDocumentStrict)
		},
		"transitional root with strict workbook": func(entries map[string]string) {
			entries["xl/workbook.xml"] = strings.ReplaceAll(entries["xl/workbook.xml"], spreadsheetMLTransitional, spreadsheetMLStrict)
		},
		"opposing styles relationship": func(entries map[string]string) {
			entries["xl/_rels/workbook.xml.rels"] = strings.ReplaceAll(entries["xl/_rels/workbook.xml.rels"], relTypeStylesTransitional, relTypeStylesStrict)
		},
	} {
		t.Run(name, func(t *testing.T) {
			entries := styleMutationFixture()
			mutate(entries)
			output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation})
			if err == nil || output != nil {
				t.Fatalf("opposing routing dialect should fail atomically, output=%d err=%v", len(output), err)
			}
		})
	}

	t.Run("case-insensitive MIME and exact Internal", func(t *testing.T) {
		entries := styleMutationFixture()
		entries["[Content_Types].xml"] = strings.ReplaceAll(entries["[Content_Types].xml"], stylesPartContentType, strings.ToUpper(stylesPartContentType))
		entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], ` Target="styles.xml"`, ` Target="styles.xml" TargetMode="Internal"`, 1)
		if output, err := ApplyStyleMutations(buildZip(t, entries), []StylePatchMutation{mutation}); err != nil || output == nil {
			t.Fatalf("case-insensitive styles MIME with exact Internal should succeed, output=%d err=%v", len(output), err)
		}
	})
}
