package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"encoding/json"
	"slices"
	"strings"
	"testing"
)

func nativeWorkbookFixture(strict bool) map[string]string {
	ssNS, relNS := spreadsheetMLTransitional, officeRelNamespaceTransitional
	officeType, sheetType := relTypeOfficeDocumentTransitional, relTypeWorksheetTransitional
	styleType, sharedType := relTypeStylesTransitional, relTypeSharedStringsTransitional
	hyperlinkType := "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
	if strict {
		ssNS, relNS = spreadsheetMLStrict, officeRelNamespaceStrict
		officeType, sheetType = relTypeOfficeDocumentStrict, relTypeWorksheetStrict
		styleType, sharedType = relTypeStylesStrict, relTypeSharedStringsStrict
		hyperlinkType = "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink"
	}
	return map[string]string{
		"[Content_Types].xml": `<Types xmlns="` + nativeContentTypesNamespace + `">` +
			`<Default Extension="rels" ContentType="` + nativeRelationshipsType + `"/>` +
			`<Default Extension="style" ContentType="` + stylesPartContentType + `"/>` +
			`<Default Extension="bin" ContentType="application/octet-stream"/>` +
			`<Override PartName="/book/%57ORKBOOK.XML" ContentType="` + strings.ToUpper(nativeWorkbookContentType) + `"/>` +
			`<Override PartName="/SHEETS/S1.XML" ContentType="` + nativeWorksheetContentType + `"/>` +
			`<Override PartName="/Sheets/s2.xml" ContentType="` + nativeWorksheetContentType + `"/>` +
			`<Override PartName="/Meta/Strings.XML" ContentType="` + nativeSharedStringsType + `"/>` +
			`<Override PartName="/Charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>` +
			`</Types>`,
		"_rels/.rels": `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rOffice" Type="` + officeType + `" Target="/BOOK/%57orkbook.xml#package"/></Relationships>`,
		"Book/Workbook.xml": `<?xml version="1.0"?><workbook xmlns="` + ssNS + `" xmlns:r="` + relNS + `"><bookViews><workbookView/></bookViews><sheets>` +
			`<sheet name="Data_x0020_Set" sheetId="7" r:id="rSheet1"/><sheet name="Hidden" sheetId="9" state="hidden" r:id="rSheet2"/>` +
			`</sheets><calcPr calcId="191029"/></workbook>`,
		"Book/_rels/Workbook.xml.rels": `<Relationships xmlns="` + packageRelationshipsNamespace + `">` +
			`<Relationship Id="rSheet1" Type="` + sheetType + `" Target="../SHEETS/S1.XML"/>` +
			`<Relationship Id="rSheet2" Type="` + sheetType + `" Target="../Sheets/s2.xml" TargetMode="Internal"/>` +
			`<Relationship Id="rStyles" Type="` + styleType + `" Target="../META/Styles.STYLE"/>` +
			`<Relationship Id="rStrings" Type="` + sharedType + `" Target="../Meta/Strings.XML"/>` +
			`</Relationships>`,
		"Meta/Strings.xml": `<sst xmlns="` + ssNS + `" count="2" uniqueCount="2"><si><r><rPr><b/></rPr><t>Rich </t></r><r><t>Text</t></r></si><si><t>Plain_x0020_Text</t></si></sst>`,
		"Meta/Styles.style": `<styleSheet xmlns="` + ssNS + `"><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>` +
			`<fonts count="2"><font><name val="Calibri"/><family val="2"/><color theme="1"/><sz val="11"/><scheme val="minor"/></font><font><name val="Aptos"/><b/><i/><color rgb="FF112233"/><sz val="12.5"/></font></fonts>` +
			`<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFAABBCC"/><bgColor indexed="64"/></patternFill></fill></fills>` +
			`<borders count="2"><border/><border><left style="thin"><color rgb="FF102030"/></left><right style="double"><color rgb="FF405060"/></right><top style="dashed"><color rgb="FF708090"/></top><bottom style="medium"><color rgb="FFA0B0C0"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
			`<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs>` +
			`<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
		"Sheets/s1.xml": `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:K2"/><sheetFormatPr baseColWidth="8" defaultColWidth="10.6640625" defaultRowHeight="15" customHeight="0" zeroHeight="0"/><cols><col min="1" max="2" width="12.25" customWidth="1"/><col min="4" max="4" hidden="1" style="1"/></cols><sheetData>` +
			`<row r="1" ht="20" customHeight="1"><c r="A1" s="1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>Inline_x0020_Value</t></is></c><c r="C1"><v>001.2300</v></c><c r="D1" t="b"><v>1</v></c><c r="E1" t="e"><v>#DIV/0!</v></c><c r="F1" t="d"><v>2026-08-27T12:00:00Z</v></c><c r="G1" t="str"><f>CONCAT(&quot;a&quot;,&quot;b&quot;)</f><v>cached_x0020_text</v></c><c r="H1"><f t="shared" si="0" ref="H1:H2">SUM(C1)</f><v>2</v></c><c r="J1"><f t="array" ref="J1:J2">ROW()</f><v>1</v></c><c r="K1"><f t="dataTable" ref="K1:K2"/><v>2</v></c></row>` +
			`<row><c r="H2"><f t="shared" si="0"></f><v>3</v></c><c r="I2" s="1"/><c r="J2"><v>4</v></c><c r="K2"><v>5</v></c></row></sheetData><mergeCells count="1"><mergeCell ref="A1:A2"/></mergeCells><conditionalFormatting sqref="C1"><cfRule type="expression" priority="1"><formula>C1&gt;0</formula></cfRule></conditionalFormatting></worksheet>`,
		"Sheets/s2.xml":            `<worksheet xmlns="` + ssNS + `"><sheetData><row r="1"><c r="A1" t="s"><v>1</v></c></row></sheetData></worksheet>`,
		"Sheets/_rels/s1.xml.rels": `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rLink" Type="` + hyperlinkType + `" Target="mailto:test@example.com" TargetMode="External"/></Relationships>`,
		"Charts/chart1.xml":        `<chartSpace><chart/></chartSpace>`,
		"Custom/data.bin":          "opaque-native-bytes",
	}
}

func findNativeCell(t *testing.T, workbook *NativeWorkbookV1, sheet, ref string) NativeWorkbookCellV1 {
	t.Helper()
	for _, candidate := range workbook.Sheets {
		if candidate.ID != sheet {
			continue
		}
		for _, cell := range candidate.Cells {
			if cell.Ref == ref {
				return cell
			}
		}
	}
	t.Fatalf("cell %s!%s not found", sheet, ref)
	return NativeWorkbookCellV1{}
}

func hasNativeWorkbookUnsupported(workbook *NativeWorkbookV1, code string) bool {
	for _, item := range workbook.Unsupported {
		if item.Code == code {
			return true
		}
	}
	return false
}

func TestExtractNativeWorkbookV1EndToEndLexicalFidelity(t *testing.T) {
	data := buildZip(t, nativeWorkbookFixture(false))
	workbook, err := ExtractNativeWorkbookV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if workbook.Protocol != NativeXLSXProtocol || workbook.Version != 1 || len(workbook.Revision) != len("rev:")+64 || len(workbook.Source.PackageSHA256) != len("sha256:")+64 {
		t.Fatalf("invalid identity envelope: %#v", workbook)
	}
	if workbook.Source.WorkbookPart != "Book/Workbook.xml" || workbook.Source.Dialect != "transitional" {
		t.Fatalf("actual workbook spelling/dialect was lost: %#v", workbook.Source)
	}
	if workbook.NormalStyle == nil || workbook.NormalStyle.StyleXFID != 0 || workbook.NormalStyle.FontID != 0 || workbook.NormalStyle.FontName != "Calibri" || workbook.NormalStyle.FontSizePoints != 11 || workbook.NormalStyle.FontBold || workbook.NormalStyle.FontItalic || !nativeWorkbookSHA.MatchString(workbook.NormalStyle.FontRecordSHA256) {
		t.Fatalf("Normal-style font authority mismatch: %#v", workbook.NormalStyle)
	}
	if len(workbook.Sheets) != 2 || workbook.Sheets[0].ID != "7" || workbook.Sheets[0].Name != "Data Set" || workbook.Sheets[1].State != "hidden" {
		t.Fatalf("sheet metadata mismatch: %#v", workbook.Sheets)
	}
	format := workbook.Sheets[0].SheetFormat
	if format == nil || format.BaseColumnWidth == nil || *format.BaseColumnWidth != 8 || format.DefaultColumnWidth == nil || *format.DefaultColumnWidth != 10.6640625 || format.DefaultRowHeightPoints != 15 || format.CustomHeight || format.ZeroHeight {
		t.Fatalf("sheet format geometry mismatch: %#v", format)
	}
	if got := workbook.Sheets[0].MergedRanges; len(got) != 1 || got[0] != (NativeWorkbookMergedRangeV1{Ref: "A1:A2", Row: 0, Column: 0, EndRow: 1, EndColumn: 0, Editable: false}) {
		t.Fatalf("merged ranges mismatch: %#v", got)
	}
	if len(workbook.Sheets[0].Rows) != 1 || workbook.Sheets[0].Rows[0].Row != 0 || *workbook.Sheets[0].Rows[0].HeightPoints != 20 {
		t.Fatalf("row dimensions mismatch: %#v", workbook.Sheets[0].Rows)
	}
	if len(workbook.Sheets[0].Columns) != 2 || workbook.Sheets[0].Columns[0].Column != 0 || workbook.Sheets[0].Columns[0].EndColumn != 1 || *workbook.Sheets[0].Columns[0].Width != 12.25 || !workbook.Sheets[0].Columns[1].Hidden || *workbook.Sheets[0].Columns[1].StyleID != 1 {
		t.Fatalf("column dimensions mismatch: %#v", workbook.Sheets[0].Columns)
	}

	shared := findNativeCell(t, workbook, "7", "A1")
	if shared.OOXMLType == nil || *shared.OOXMLType != "s" || shared.Value == nil || shared.Value.Storage != "shared" || shared.Value.Text == nil || *shared.Value.Text != "Rich Text" || !shared.Value.Rich || shared.Editable {
		t.Fatalf("rich shared string was flattened or editable: %#v", shared)
	}
	inline := findNativeCell(t, workbook, "7", "B1")
	if inline.Value == nil || inline.Value.Storage != "inline" || *inline.Value.Text != "Inline Value" || !inline.Editable {
		t.Fatalf("inline string mismatch: %#v", inline)
	}
	number := findNativeCell(t, workbook, "7", "C1")
	if number.Value == nil || number.Value.Lexical == nil || *number.Value.Lexical != "001.2300" || number.Value.Kind != "number" {
		t.Fatalf("numeric lexical value was coerced: %#v", number)
	}
	for ref, kind := range map[string]string{"D1": "boolean", "E1": "error", "F1": "date"} {
		if cell := findNativeCell(t, workbook, "7", ref); cell.Value == nil || cell.Value.Kind != kind {
			t.Errorf("%s kind = %#v, want %s", ref, cell.Value, kind)
		}
	}
	formula := findNativeCell(t, workbook, "7", "G1")
	if formula.Formula == nil || formula.Formula.Text != `CONCAT("a","b")` || formula.Formula.Cached == nil || formula.Formula.Cached.Text == nil || *formula.Formula.Cached.Text != "cached text" || formula.Value != nil || !formula.Editable {
		t.Fatalf("normal formula/cache mismatch: %#v", formula)
	}
	master, follower := findNativeCell(t, workbook, "7", "H1"), findNativeCell(t, workbook, "7", "H2")
	if master.Formula == nil || master.Formula.Type != "shared" || master.Formula.SharedIndex == nil || *master.Formula.SharedIndex != 0 || master.Editable || follower.Editable || follower.Formula.Text != "" {
		t.Fatalf("shared formula group was not surfaced read-only: master=%#v follower=%#v", master, follower)
	}
	arrayMaster, arrayFollower := findNativeCell(t, workbook, "7", "J1"), findNativeCell(t, workbook, "7", "J2")
	dataTableMaster, dataTableFollower := findNativeCell(t, workbook, "7", "K1"), findNativeCell(t, workbook, "7", "K2")
	if arrayMaster.Formula == nil || arrayMaster.Formula.Type != "array" || arrayFollower.Editable || dataTableMaster.Formula == nil || dataTableMaster.Formula.Type != "dataTable" || dataTableFollower.Editable {
		t.Fatalf("array/data-table followers are not mutation-refused: array=%#v/%#v data=%#v/%#v", arrayMaster, arrayFollower, dataTableMaster, dataTableFollower)
	}
	blank := findNativeCell(t, workbook, "7", "I2")
	if blank.StyleID != 1 || blank.Value != nil || blank.Formula != nil {
		t.Fatalf("styled blank cell mismatch: %#v", blank)
	}

	if len(workbook.Styles) != 2 {
		t.Fatalf("style projection count = %d", len(workbook.Styles))
	}
	style := workbook.Styles[1].Effective
	if style.NumberFormat == nil || *style.NumberFormat != "0.000" || style.FontName == nil || *style.FontName != "Aptos" || style.FontColor == nil || *style.FontColor != "#112233" || style.FillColor == nil || *style.FillColor != "#AABBCC" || style.VerticalAlignment == nil || *style.VerticalAlignment != "middle" || style.WrapText == nil || !*style.WrapText || style.Projection != "full" {
		t.Fatalf("effective style projection mismatch: %#v", style)
	}
	border := style.Border
	if border == nil || border.Origin != "styles-record" || border.BorderID == nil || *border.BorderID != 1 || !nativeWorkbookSHA.MatchString(nativeOptionalString(border.RecordSHA256)) || border.Left == nil || border.Left.Style != "thin" || border.Left.Color != "#102030" || border.Right == nil || border.Right.Style != "double" || border.Right.Color != "#405060" || border.Top == nil || border.Top.Style != "dashed" || border.Top.Color != "#708090" || border.Bottom == nil || border.Bottom.Style != "medium" || border.Bottom.Color != "#A0B0C0" {
		t.Fatalf("effective border projection mismatch: %#v", border)
	}
	if workbook.Styles[0].Effective.Projection != "partial" || !hasNativeWorkbookUnsupported(workbook, "STYLE_FONT_COLOR") {
		t.Fatalf("theme color ambiguity was not inventoried: %#v / %#v", workbook.Styles[0], workbook.Unsupported)
	}
	for _, code := range []string{"RICH_SHARED_STRING", "RICH_CELL_STRING", "FORMULA_SHARED", "FORMULA_ARRAY", "FORMULA_DATATABLE", "FORMULA_GROUP_RANGE", "MERGED_CELLS", "CONDITIONAL_FORMATTING", "EXTERNAL_RELATIONSHIP", "CHART_CONTENT", "OPAQUE_PACKAGE_PART"} {
		if !hasNativeWorkbookUnsupported(workbook, code) {
			t.Errorf("missing unsupported capability %s", code)
		}
	}
	foundOpaque := false
	for _, part := range workbook.PassthroughParts {
		if part.PartName == "Custom/data.bin" {
			foundOpaque = true
			if part.ByteLength == nil || *part.ByteLength != int64(len("opaque-native-bytes")) || len(part.SHA256) != len("sha256:")+64 || part.Policy != "preserve-exact" {
				t.Fatalf("bad passthrough fingerprint: %#v", part)
			}
		}
	}
	if !foundOpaque {
		t.Fatal("opaque part is absent from passthrough inventory")
	}
	encoded, err := EncodeNativeWorkbookV1(workbook)
	if err != nil {
		t.Fatal(err)
	}
	decoded, err := DecodeNativeWorkbookV1(encoded)
	if err != nil || decoded.Revision != workbook.Revision {
		t.Fatalf("contract round trip failed: decoded=%#v err=%v", decoded, err)
	}
	again, err := ExtractNativeWorkbookV1(data)
	if err != nil {
		t.Fatal(err)
	}
	encodedAgain, err := EncodeNativeWorkbookV1(again)
	if err != nil || !bytes.Equal(encodedAgain, encoded) {
		t.Fatalf("native extraction is not deterministic: err=%v", err)
	}
}

func TestExtractNativeWorkbookV1BorderInheritanceAndAmbiguity(t *testing.T) {
	t.Run("cellStyleXf inheritance", func(t *testing.T) {
		entries := nativeWorkbookFixture(false)
		entries["Meta/Styles.style"] = strings.Replace(entries["Meta/Styles.style"],
			`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`,
			`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="1"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0"/>`, 1)
		workbook, err := ExtractNativeWorkbookV1(buildZip(t, entries))
		if err != nil {
			t.Fatal(err)
		}
		border := workbook.Styles[0].Effective.Border
		if border == nil || border.BorderID == nil || *border.BorderID != 1 || border.Left == nil || border.Left.Style != "thin" {
			t.Fatalf("inherited border was not projected exactly: %#v", border)
		}
	})

	for _, test := range []struct {
		name, old, replacement string
	}{
		{name: "theme color", old: `<color rgb="FF102030"/>`, replacement: `<color theme="4"/>`},
		{name: "diagonal", old: `<diagonal/>`, replacement: `<diagonal style="thin"><color rgb="FF102030"/></diagonal>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			entries := nativeWorkbookFixture(false)
			entries["Meta/Styles.style"] = strings.Replace(entries["Meta/Styles.style"], test.old, test.replacement, 1)
			workbook, err := ExtractNativeWorkbookV1(buildZip(t, entries))
			if err != nil {
				t.Fatal(err)
			}
			effective := workbook.Styles[1].Effective
			if effective.Projection != "partial" || effective.Border != nil || !slices.Contains(effective.Unsupported, "border") || !hasNativeWorkbookUnsupported(workbook, "STYLE_BORDER") {
				t.Fatalf("ambiguous border was not failed closed: %#v / %#v", effective, workbook.Unsupported)
			}
		})
	}

	t.Run("applyBorder mismatch", func(t *testing.T) {
		entries := nativeWorkbookFixture(false)
		entries["Meta/Styles.style"] = strings.Replace(entries["Meta/Styles.style"], ` applyBorder="1"`, ``, 1)
		if workbook, err := ExtractNativeWorkbookV1(buildZip(t, entries)); err == nil || workbook != nil || !strings.Contains(err.Error(), "border id differs") {
			t.Fatalf("border override without applyBorder was accepted: workbook=%#v err=%v", workbook, err)
		}
	})
}

func TestNativeStyleRawProjectionDigestRejectsValidLookingTamper(t *testing.T) {
	original, err := ExtractNativeWorkbookV1(buildZip(t, nativeWorkbookFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	clone := func() *NativeWorkbookV1 {
		encoded, encodeErr := EncodeNativeWorkbookV1(original)
		if encodeErr != nil {
			t.Fatal(encodeErr)
		}
		copy, decodeErr := DecodeNativeWorkbookV1(encoded)
		if decodeErr != nil {
			t.Fatal(decodeErr)
		}
		return copy
	}
	for _, mutate := range []func(*NativeWorkbookV1){
		func(workbook *NativeWorkbookV1) {
			color := "#ABCDEF"
			workbook.Styles[1].Effective.FillColor = &color
			workbook.Styles[1].Effective.Fill.Color = &color
		},
		func(workbook *NativeWorkbookV1) { workbook.Styles[1].Effective.Border.Left.Style = "dashed" },
		func(workbook *NativeWorkbookV1) { workbook.Styles[1].Effective.Border.Left.Color = "#ABCDEF" },
	} {
		candidate := clone()
		mutate(candidate)
		if issues := ValidateNativeWorkbookV1(candidate); len(issues) == 0 || issues[0].Path != "/styles/1/raw_projection_sha256" {
			t.Fatalf("valid-looking stale style projection tamper was accepted: %+v", issues)
		}
	}
}

func TestNativeStyleRawProjectionDigestMatchesTypeScriptEscaping(t *testing.T) {
	fontName := "A\u2028B\u2029C"
	digest, err := nativeRawStyleProjectionDigest(NativeWorkbookEffectiveStyleV1{
		FontName: &fontName,
		Fill:     &NativeWorkbookFillV1{Origin: "implicit-default"}, Border: &NativeWorkbookBorderV1{Origin: "implicit-default"},
		Projection: "full", Unsupported: []string{},
	})
	if err != nil {
		t.Fatal(err)
	}
	if digest != "sha256:29a0ff04e8fc39c1f1844e6d23e1e6411c500577f7e91dd90c1b497afd276098" {
		t.Fatalf("canonical style projection digest = %s", digest)
	}
}

func TestExtractNativeWorkbookV1StrictGeometryAndNormalStyleAuthority(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts := nativeWorkbookFixture(strict)
		parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"],
			`<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>`,
			`<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0"/></cellStyleXfs>`, 1)
		parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `name="Normal" xfId="0" builtinId="0"`, `name="Normal" xfId="1" builtinId="0"`, 1)
		workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
		if err != nil {
			t.Fatalf("strict=%v: %v", strict, err)
		}
		if workbook.NormalStyle == nil || workbook.NormalStyle.StyleXFID != 1 || workbook.NormalStyle.FontID != 1 || workbook.NormalStyle.FontName != "Aptos" || workbook.NormalStyle.FontSizePoints != 12.5 || !workbook.NormalStyle.FontBold || !workbook.NormalStyle.FontItalic {
			t.Fatalf("strict=%v Normal style followed cellXf0 instead of builtin Normal: %#v", strict, workbook.NormalStyle)
		}
		if workbook.Sheets[0].SheetFormat == nil || workbook.Sheets[0].SheetFormat.DefaultRowHeightPoints != 15 {
			t.Fatalf("strict=%v sheetFormatPr geometry missing: %#v", strict, workbook.Sheets[0].SheetFormat)
		}
	}
}

func TestExtractNativeWorkbookV1InventoriesFormulaViewGeometry(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<dimension ref="A1:K2"/>`, `<dimension ref="A1:K2"/><sheetViews><sheetView workbookViewId="0" showFormulas="1"/></sheetViews>`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	if !hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") {
		t.Fatalf("showFormulas view did not remain source-authoritative: %#v", workbook.Unsupported)
	}
}

func TestExtractNativeWorkbookV1NarrowsWorksheetMetadataAuthority(t *testing.T) {
	t.Run("selection-only view is benign UI state", func(t *testing.T) {
		parts := nativeWorkbookFixture(false)
		parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<dimension ref="A1:K2"/>`, `<dimension ref="A1:K2"/><sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>`, 1)
		workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		if hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") {
			t.Fatalf("selection-only UI state became geometry authority: %#v", workbook.Unsupported)
		}
	})

	t.Run("excel-default gridlines-on is benign UI state", func(t *testing.T) {
		parts := nativeWorkbookFixture(false)
		parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<dimension ref="A1:K2"/>`, `<dimension ref="A1:K2"/><sheetViews><sheetView showGridLines="1" workbookViewId="0"><selection activeCell="A1" sqref="A1"/></sheetView></sheetViews>`, 1)
		workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		if hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") {
			t.Fatalf("showGridLines=1 became geometry authority: %#v", workbook.Unsupported)
		}
	})

	t.Run("explicit gridlines-off remains view geometry", func(t *testing.T) {
		parts := nativeWorkbookFixture(false)
		parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<dimension ref="A1:K2"/>`, `<dimension ref="A1:K2"/><sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>`, 1)
		workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		if !hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") {
			t.Fatal("showGridLines=0 must remain source-authoritative")
		}
	})

	t.Run("unknown worksheet root attribute fails closed", func(t *testing.T) {
		parts := nativeWorkbookFixture(false)
		parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<worksheet xmlns="`+spreadsheetMLTransitional+`">`, `<worksheet xmlns="`+spreadsheetMLTransitional+`" strange="opaque">`, 1)
		workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
		if err != nil {
			t.Fatal(err)
		}
		if !hasNativeWorkbookUnsupported(workbook, "WORKSHEET_ATTRIBUTES") || workbook.Sheets[0].Editable || workbook.Sheets[0].RefusalCode == nil {
			t.Fatalf("unknown worksheet root authority did not fail closed: sheet=%#v unsupported=%#v", workbook.Sheets[0], workbook.Unsupported)
		}
	})

	for _, test := range []struct {
		name, old, replacement, code string
	}{
		{name: "sheet default descent", old: `<sheetFormatPr baseColWidth="8"`, replacement: `<sheetFormatPr x14ac:dyDescent="0.25" baseColWidth="8"`, code: "SHEET_FORMAT_EXTRAS"},
		{name: "row descent", old: `<row r="1" ht="20"`, replacement: `<row x14ac:dyDescent="0.25" r="1" ht="20"`, code: "ROW_DIMENSION_EXTRAS"},
	} {
		t.Run(test.name+" fails closed", func(t *testing.T) {
			parts := nativeWorkbookFixture(false)
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<worksheet xmlns="`+spreadsheetMLTransitional+`">`, `<worksheet xmlns="`+spreadsheetMLTransitional+`" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">`, 1)
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], test.old, test.replacement, 1)
			workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
			if err != nil {
				t.Fatal(err)
			}
			if !hasNativeWorkbookUnsupported(workbook, test.code) {
				t.Fatalf("x14ac:dyDescent did not remain source-authoritative as %s: %#v", test.code, workbook.Unsupported)
			}
		})
	}
}

func TestExtractNativeWorkbookV1StrictAndPreviousIdentity(t *testing.T) {
	firstBytes := buildZip(t, nativeWorkbookFixture(true))
	first, err := ExtractNativeWorkbookV1(firstBytes)
	if err != nil {
		t.Fatal(err)
	}
	if first.Source.Dialect != "strict" {
		t.Fatalf("dialect = %q", first.Source.Dialect)
	}
	parts := nativeWorkbookFixture(true)
	parts["Custom/data.bin"] = "changed opaque bytes"
	next, err := ExtractNativeWorkbookV1WithOptions(buildZip(t, parts), NativeWorkbookExtractionOptions{Previous: first})
	if err != nil {
		t.Fatal(err)
	}
	if next.DocumentID != first.DocumentID || next.Revision == first.Revision {
		t.Fatalf("previous identity/revision mismatch: first=%s/%s next=%s/%s", first.DocumentID, first.Revision, next.DocumentID, next.Revision)
	}
	explicit, err := ExtractNativeWorkbookV1WithOptions(firstBytes, NativeWorkbookExtractionOptions{DocumentID: "workbook:application-owned"})
	if err != nil || explicit.DocumentID != "workbook:application-owned" {
		t.Fatalf("explicit document id failed: %#v err=%v", explicit, err)
	}
	if _, err := ExtractNativeWorkbookV1WithOptions(firstBytes, NativeWorkbookExtractionOptions{Previous: first, DocumentID: "workbook:other"}); err == nil || !strings.Contains(err.Error(), "conflicts") {
		t.Fatalf("conflicting identity error = %v", err)
	}
	remappedParts := nativeWorkbookFixture(true)
	remappedParts["Sheets/s3.xml"] = remappedParts["Sheets/s1.xml"]
	remappedParts["[Content_Types].xml"] = strings.Replace(remappedParts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Sheets/s3.xml" ContentType="`+nativeWorksheetContentType+`"/></Types>`, 1)
	remappedParts["Book/_rels/Workbook.xml.rels"] = strings.Replace(remappedParts["Book/_rels/Workbook.xml.rels"], `Target="../SHEETS/S1.XML"`, `Target="../Sheets/s3.xml"`, 1)
	remapped, err := ExtractNativeWorkbookV1WithOptions(buildZip(t, remappedParts), NativeWorkbookExtractionOptions{Previous: first})
	if err == nil || remapped != nil || !strings.Contains(err.Error(), "Previous sheet id") {
		t.Fatalf("Previous sheet identity remap was accepted: %#v err=%v", remapped, err)
	}
}

func TestExtractNativeWorkbookV1SupportsRootLevelWorkbookRelationshipPart(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Workbook.xml"] = parts["Book/Workbook.xml"]
	delete(parts, "Book/Workbook.xml")
	parts["_rels/Workbook.xml.rels"] = strings.ReplaceAll(parts["Book/_rels/Workbook.xml.rels"], "../", "")
	delete(parts, "Book/_rels/Workbook.xml.rels")
	parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `/BOOK/%57orkbook.xml#package`, `/Workbook.xml#package`, 1)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `/book/%57ORKBOOK.XML`, `/Workbook.xml`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	if workbook.Source.WorkbookPart != "Workbook.xml" || workbook.Sheets[0].PartName != "Sheets/s1.xml" {
		t.Fatalf("root-level routing lost actual spelling: source=%#v sheet=%#v", workbook.Source, workbook.Sheets[0])
	}
}

func TestExtractNativeWorkbookV1RejectsAdversarialCore(t *testing.T) {
	tests := []struct {
		name string
		edit func(map[string]string)
		want string
	}{
		{name: "target mode is exact", edit: func(parts map[string]string) {
			parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], `TargetMode="Internal"`, `TargetMode="internal"`, 1)
		}, want: `TargetMode "internal"`},
		{name: "opposing worksheet dialect", edit: func(parts map[string]string) {
			parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], relTypeWorksheetTransitional, relTypeWorksheetStrict, 1)
		}, want: "opposing Strict/Transitional"},
		{name: "spoofed relationship namespace", edit: func(parts map[string]string) {
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], packageRelationshipsNamespace, "urn:spoofed", 1)
		}, want: "root is not package Relationships"},
		{name: "duplicate stable sheet id", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `sheetId="9"`, `sheetId="7"`, 1)
		}, want: "duplicated"},
		{name: "duplicate case-insensitive sheet name", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `name="Hidden"`, `name="data_x0020_set"`, 1)
		}, want: "sheet name"},
		{name: "duplicate Unicode case-insensitive sheet name", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `name="Data_x0020_Set"`, `name="Σ"`, 1)
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `name="Hidden"`, `name="ς"`, 1)
		}, want: "sheet name"},
		{name: "boundary apostrophe sheet name", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `name="Data_x0020_Set"`, `name="'Data_x0020_Set"`, 1)
		}, want: "cannot begin or end with an apostrophe"},
		{name: "duplicate content override", edit: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Book/Workbook.xml" ContentType="`+nativeWorkbookContentType+`"/></Types>`, 1)
		}, want: "duplicate case/escape-equivalent"},
		{name: "orphan content override", edit: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/missing.xml" ContentType="application/xml"/></Types>`, 1)
		}, want: "targets missing part"},
		{name: "wrong worksheet content type", edit: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], nativeWorksheetContentType, "application/xml", 1)
		}, want: "expected"},
		{name: "missing internal relationship target", edit: func(parts map[string]string) {
			delete(parts, "Charts/chart1.xml")
			parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], `</Relationships>`, `<Relationship Id="missing" Type="urn:opaque" Target="../Charts/chart1.xml"/></Relationships>`, 1)
		}, want: "targets missing part"},
		{name: "shared string index", edit: func(parts map[string]string) {
			parts["Sheets/s2.xml"] = strings.Replace(parts["Sheets/s2.xml"], `<v>1</v>`, `<v>2</v>`, 1)
		}, want: "shared-string index"},
		{name: "style id", edit: func(parts map[string]string) {
			parts["Sheets/s2.xml"] = strings.Replace(parts["Sheets/s2.xml"], `<c r="A1"`, `<c r="A1" s="2"`, 1)
		}, want: "outside cellXfs"},
		{name: "overlapping columns", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<col min="4" max="4"`, `<col min="2" max="4"`, 1)
		}, want: "overlap"},
		{name: "mismatched cell row", edit: func(parts map[string]string) {
			parts["Sheets/s2.xml"] = strings.Replace(parts["Sheets/s2.xml"], `r="A1"`, `r="A2"`, 1)
		}, want: "does not belong"},
		{name: "DTD", edit: func(parts map[string]string) {
			parts["Sheets/s2.xml"] = strings.Replace(parts["Sheets/s2.xml"], `<worksheet`, `<!DOCTYPE worksheet [<!ENTITY x "boom">]><worksheet`, 1)
		}, want: "directive"},
		{name: "shared count mismatch", edit: func(parts map[string]string) {
			parts["Meta/Strings.xml"] = strings.Replace(parts["Meta/Strings.xml"], `uniqueCount="2"`, `uniqueCount="1"`, 1)
		}, want: "does not match"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			parts := nativeWorkbookFixture(false)
			test.edit(parts)
			workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
			if err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("workbook=%#v err=%v, want %q", workbook, err, test.want)
			}
		})
	}
}

func TestExtractNativeWorkbookV1RejectsCaseEquivalentZIPPartsAndBomb(t *testing.T) {
	base := nativeWorkbookFixture(false)
	var archive bytes.Buffer
	writer := zip.NewWriter(&archive)
	for name, content := range base {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write([]byte(content)); err != nil {
			t.Fatal(err)
		}
	}
	duplicate, err := writer.Create("book/workbook.xml")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := duplicate.Write([]byte(base["Book/Workbook.xml"])); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := ExtractNativeWorkbookV1(archive.Bytes()); err == nil || !strings.Contains(err.Error(), "case/escape-equivalent") {
		t.Fatalf("case-equivalent duplicate error = %v", err)
	}

	bomb := nativeWorkbookFixture(false)
	bomb["Custom/bomb.bin"] = string(bytes.Repeat([]byte{0}, 2*NativeXLSXCompressionRatioSlack))
	if _, err := ExtractNativeWorkbookV1(buildZip(t, bomb)); err == nil || !strings.Contains(err.Error(), "compression-ratio limit") {
		t.Fatalf("zip bomb error = %v", err)
	}
}

func TestDecodeNativeWorkbookV1RejectsUnknownFieldsAndInvalidPrevious(t *testing.T) {
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, nativeWorkbookFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := EncodeNativeWorkbookV1(workbook)
	if err != nil {
		t.Fatal(err)
	}
	unknown := bytes.Replace(encoded, []byte(`"version":1`), []byte(`"version":1,"luckyexcel":true`), 1)
	if _, err := DecodeNativeWorkbookV1(unknown); err == nil || !strings.Contains(err.Error(), "unknown field") {
		t.Fatalf("unknown field error = %v", err)
	}
	bad := *workbook
	bad.Revision = "rev:short"
	if _, err := ExtractNativeWorkbookV1WithOptions(buildZip(t, nativeWorkbookFixture(false)), NativeWorkbookExtractionOptions{Previous: &bad}); err == nil || !strings.Contains(err.Error(), "previous identity contract is invalid") {
		t.Fatalf("invalid Previous error = %v", err)
	}
}

func cloneNativeWorkbookForTest(t *testing.T, workbook *NativeWorkbookV1) *NativeWorkbookV1 {
	t.Helper()
	data, err := json.Marshal(workbook)
	if err != nil {
		t.Fatal(err)
	}
	var clone NativeWorkbookV1
	if err := json.Unmarshal(data, &clone); err != nil {
		t.Fatal(err)
	}
	return &clone
}

func nativeCellPointerForTest(t *testing.T, workbook *NativeWorkbookV1, sheetID, ref string) *NativeWorkbookCellV1 {
	t.Helper()
	for sheetIndex := range workbook.Sheets {
		if workbook.Sheets[sheetIndex].ID != sheetID {
			continue
		}
		for cellIndex := range workbook.Sheets[sheetIndex].Cells {
			if workbook.Sheets[sheetIndex].Cells[cellIndex].Ref == ref {
				return &workbook.Sheets[sheetIndex].Cells[cellIndex]
			}
		}
	}
	t.Fatalf("cell %s!%s not found", sheetID, ref)
	return nil
}

func nativeUnsupportedPointerForTest(t *testing.T, workbook *NativeWorkbookV1, code string) *NativeWorkbookUnsupportedV1 {
	t.Helper()
	for index := range workbook.Unsupported {
		if workbook.Unsupported[index].Code == code {
			return &workbook.Unsupported[index]
		}
	}
	t.Fatalf("unsupported code %s not found", code)
	return nil
}

func recanonicalizeNativeUnsupportedIDForTest(item *NativeWorkbookUnsupportedV1) {
	key := strings.Join([]string{item.Code, item.Capability, item.ScopeID, nativeOptionalString(item.PartName), nativeOptionalString(item.CellRef), nativeOptionalString(item.RangeRef)}, "\x00")
	digest := sha256.Sum256([]byte(key))
	item.ID = "unsupported:" + hex.EncodeToString(digest[:])
}

func TestDecodeNativeWorkbookV1RejectsEveryUntrustedContractFamily(t *testing.T) {
	valid, err := ExtractNativeWorkbookV1(buildZip(t, nativeWorkbookFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	tooLargePart := int64(NativeXLSXMaxPartBytes + 1)
	tests := []struct {
		name string
		edit func(*NativeWorkbookV1)
	}{
		{name: "nil top-level collection", edit: func(value *NativeWorkbookV1) { value.Unsupported = nil }},
		{name: "empty sheets", edit: func(value *NativeWorkbookV1) { value.Sheets = []NativeWorkbookSheetV1{} }},
		{name: "nil sheet cells", edit: func(value *NativeWorkbookV1) { value.Sheets[0].Cells = nil }},
		{name: "editable refusal conflict", edit: func(value *NativeWorkbookV1) { value.Sheets[1].RefusalCode = nativeWorkbookString("SHEET_PROTECTION") }},
		{name: "noncanonical sheet id", edit: func(value *NativeWorkbookV1) { value.Sheets[0].ID = "07" }},
		{name: "value kind storage mismatch", edit: func(value *NativeWorkbookV1) { nativeCellPointerForTest(t, value, "7", "C1").Value.Kind = "string" }},
		{name: "value lexical missing", edit: func(value *NativeWorkbookV1) { nativeCellPointerForTest(t, value, "7", "C1").Value.Lexical = nil }},
		{name: "invalid numeric lexical", edit: func(value *NativeWorkbookV1) { *nativeCellPointerForTest(t, value, "7", "C1").Value.Lexical = "NaN" }},
		{name: "invalid date lexical", edit: func(value *NativeWorkbookV1) {
			*nativeCellPointerForTest(t, value, "7", "F1").Value.Lexical = "2026-02-30"
		}},
		{name: "ooxml storage mismatch", edit: func(value *NativeWorkbookV1) {
			nativeCellPointerForTest(t, value, "7", "D1").OOXMLType = nativeWorkbookString("e")
		}},
		{name: "shared formula missing index", edit: func(value *NativeWorkbookV1) { nativeCellPointerForTest(t, value, "7", "H1").Formula.SharedIndex = nil }},
		{name: "normal formula range", edit: func(value *NativeWorkbookV1) {
			nativeCellPointerForTest(t, value, "7", "G1").Formula.Ref = nativeWorkbookString("G1:G2")
		}},
		{name: "group formula editable", edit: func(value *NativeWorkbookV1) { nativeCellPointerForTest(t, value, "7", "J1").Editable = true }},
		{name: "style projection enum", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.Projection = "approximate" }},
		{name: "style unsupported enum", edit: func(value *NativeWorkbookV1) { value.Styles[0].Effective.Unsupported[0] = "theme-magic" }},
		{name: "style color", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.FontColor = nativeWorkbookString("#abc") }},
		{name: "style alignment", edit: func(value *NativeWorkbookV1) {
			value.Styles[1].Effective.VerticalAlignment = nativeWorkbookString("center")
		}},
		{name: "supported border missing", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.Border = nil }},
		{name: "border origin", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.Border.Origin = "theme-derived" }},
		{name: "border id", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.Border.BorderID = nil }},
		{name: "border digest", edit: func(value *NativeWorkbookV1) {
			value.Styles[1].Effective.Border.RecordSHA256 = nativeWorkbookString("sha256:bad")
		}},
		{name: "border style", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.Border.Left.Style = "triple" }},
		{name: "border color", edit: func(value *NativeWorkbookV1) { value.Styles[1].Effective.Border.Left.Color = "#abc" }},
		{name: "implicit border provenance", edit: func(value *NativeWorkbookV1) {
			value.Styles[0].Effective.Border.Origin = "implicit-default"
		}},
		{name: "capability duplicate", edit: func(value *NativeWorkbookV1) { value.Capabilities[1].Name = value.Capabilities[0].Name }},
		{name: "capability level", edit: func(value *NativeWorkbookV1) { value.Capabilities[0].Level = "write-all" }},
		{name: "passthrough package bound", edit: func(value *NativeWorkbookV1) { value.PassthroughParts[0].ByteLength = &tooLargePart }},
		{name: "passthrough order", edit: func(value *NativeWorkbookV1) {
			value.PassthroughParts[0], value.PassthroughParts[1] = value.PassthroughParts[1], value.PassthroughParts[0]
		}},
		{name: "unsupported id", edit: func(value *NativeWorkbookV1) { value.Unsupported[0].ID = "unsupported:bad" }},
		{name: "unsupported preservation", edit: func(value *NativeWorkbookV1) { value.Unsupported[0].Preservation = "best-effort" }},
		{name: "unsupported location union", edit: func(value *NativeWorkbookV1) {
			value.Unsupported[0].CellRef, value.Unsupported[0].RangeRef = nativeWorkbookString("A1"), nativeWorkbookString("A1:A2")
		}},
		{name: "unsupported scope", edit: func(value *NativeWorkbookV1) { value.Unsupported[0].ScopeID = "sheet:missing" }},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			candidate := cloneNativeWorkbookForTest(t, valid)
			test.edit(candidate)
			data, err := json.Marshal(candidate)
			if err != nil {
				t.Fatal(err)
			}
			if decoded, err := DecodeNativeWorkbookV1(data); err == nil || decoded != nil {
				t.Fatalf("invalid untrusted contract was accepted: %#v", decoded)
			}
			if _, err := ExtractNativeWorkbookV1WithOptions(buildZip(t, nativeWorkbookFixture(false)), NativeWorkbookExtractionOptions{Previous: candidate}); err == nil || !strings.Contains(err.Error(), "previous identity contract is invalid") {
				t.Fatalf("invalid Previous was accepted: %v", err)
			}
		})
	}
}

func TestDecodeNativeWorkbookV1RejectsDuplicateKeysAtEveryDepth(t *testing.T) {
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, nativeWorkbookFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := EncodeNativeWorkbookV1(workbook)
	if err != nil {
		t.Fatal(err)
	}
	for name, hostile := range map[string][]byte{
		"nested source": bytes.Replace(encoded, []byte(`"source":{`), []byte(`"source":{"dialect":"strict",`), 1),
		"nested cell":   bytes.Replace(encoded, []byte(`"row":0,"column":`), []byte(`"row":0,"row":0,"column":`), 1),
	} {
		t.Run(name, func(t *testing.T) {
			if decoded, err := DecodeNativeWorkbookV1(hostile); err == nil || decoded != nil || !strings.Contains(err.Error(), "duplicate object key") {
				t.Fatalf("decoded=%#v err=%v", decoded, err)
			}
		})
	}
}

func TestDecodeAndPreviousCrossBindUnsupportedAuthorityAndEditability(t *testing.T) {
	valid, err := ExtractNativeWorkbookV1(buildZip(t, nativeWorkbookFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		name string
		edit func(*NativeWorkbookV1)
	}{
		{name: "flip affected cell editable", edit: func(value *NativeWorkbookV1) { nativeCellPointerForTest(t, value, "7", "A1").Editable = true }},
		{name: "change code capability pair", edit: func(value *NativeWorkbookV1) {
			item := nativeUnsupportedPointerForTest(t, value, "RICH_CELL_STRING")
			item.Capability = "cell-markup"
			recanonicalizeNativeUnsupportedIDForTest(item)
		}},
		{name: "change canonical code", edit: func(value *NativeWorkbookV1) {
			item := nativeUnsupportedPointerForTest(t, value, "RICH_CELL_STRING")
			item.Code, item.Capability = "CELL_ATTRIBUTES", "cell-markup"
			recanonicalizeNativeUnsupportedIDForTest(item)
		}},
		{name: "remove cell source location", edit: func(value *NativeWorkbookV1) {
			item := nativeUnsupportedPointerForTest(t, value, "RICH_CELL_STRING")
			item.CellRef = nil
			recanonicalizeNativeUnsupportedIDForTest(item)
		}},
		{name: "move cell source location", edit: func(value *NativeWorkbookV1) {
			item := nativeUnsupportedPointerForTest(t, value, "RICH_CELL_STRING")
			item.CellRef = nativeWorkbookString("B1")
			recanonicalizeNativeUnsupportedIDForTest(item)
		}},
		{name: "change exact source part spelling", edit: func(value *NativeWorkbookV1) {
			item := nativeUnsupportedPointerForTest(t, value, "RICH_CELL_STRING")
			item.PartName = nativeWorkbookString("sheets/S1.xml")
			recanonicalizeNativeUnsupportedIDForTest(item)
		}},
		{name: "move formula range", edit: func(value *NativeWorkbookV1) {
			item := nativeUnsupportedPointerForTest(t, value, "FORMULA_GROUP_RANGE")
			item.RangeRef = nativeWorkbookString("A1:A2")
			recanonicalizeNativeUnsupportedIDForTest(item)
		}},
		{name: "remove authoritative reason", edit: func(value *NativeWorkbookV1) {
			filtered := make([]NativeWorkbookUnsupportedV1, 0, len(value.Unsupported)-1)
			for _, item := range value.Unsupported {
				if item.Code != "RICH_CELL_STRING" {
					filtered = append(filtered, item)
				}
			}
			value.Unsupported = filtered
		}},
		{name: "mismatch sheet refusal", edit: func(value *NativeWorkbookV1) { value.Sheets[0].RefusalCode = nativeWorkbookString("SHEET_PROTECTION") }},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			candidate := cloneNativeWorkbookForTest(t, valid)
			test.edit(candidate)
			data, err := json.Marshal(candidate)
			if err != nil {
				t.Fatal(err)
			}
			if decoded, err := DecodeNativeWorkbookV1(data); err == nil || decoded != nil {
				t.Fatalf("authority/editability forgery was decoded: %#v", decoded)
			}
			if previous, err := ExtractNativeWorkbookV1WithOptions(buildZip(t, nativeWorkbookFixture(false)), NativeWorkbookExtractionOptions{Previous: candidate}); err == nil || previous != nil || !strings.Contains(err.Error(), "previous identity contract is invalid") {
				t.Fatalf("authority/editability forgery was accepted as Previous: %#v err=%v", previous, err)
			}
		})
	}
	metadataOnly := cloneNativeWorkbookForTest(t, valid)
	nativeUnsupportedPointerForTest(t, metadataOnly, "RICH_CELL_STRING").Message = "caller-local explanatory metadata may change without changing source authority"
	data, err := json.Marshal(metadataOnly)
	if err != nil {
		t.Fatal(err)
	}
	if decoded, err := DecodeNativeWorkbookV1(data); err != nil || decoded == nil {
		t.Fatalf("non-authoritative message metadata unexpectedly changed validity: %#v err=%v", decoded, err)
	}
}

func TestExtractNativeWorkbookV1CanonicalizesSheetIDsAndRequiresBidirectionalPreviousIdentity(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `sheetId="7"`, `sheetId="007"`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
	if err != nil || workbook.Sheets[0].ID != "7" {
		t.Fatalf("workbook=%#v err=%v", workbook, err)
	}
	duplicate := nativeWorkbookFixture(false)
	duplicate["Book/Workbook.xml"] = strings.Replace(duplicate["Book/Workbook.xml"], `sheetId="9"`, `sheetId="07"`, 1)
	if _, err := ExtractNativeWorkbookV1(buildZip(t, duplicate)); err == nil || !strings.Contains(err.Error(), "duplicated") {
		t.Fatalf("numeric-equivalent sheet ids were accepted: %v", err)
	}
	previousParts := nativeWorkbookFixture(false)
	previousParts["Book/Workbook.xml"] = strings.Replace(previousParts["Book/Workbook.xml"], `sheetId="7"`, `sheetId="8"`, 1)
	previous, err := ExtractNativeWorkbookV1(buildZip(t, previousParts))
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ExtractNativeWorkbookV1WithOptions(buildZip(t, parts), NativeWorkbookExtractionOptions{Previous: previous}); err == nil || !strings.Contains(err.Error(), "has sheet id") {
		t.Fatalf("same part under a new id was accepted: %v", err)
	}
}

func TestExtractNativeWorkbookV1InventoriesSemanticAttributesAndRefusesUnsafeScopes(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `sheetId="9"`, `sheetId="9" vendor="opaque"`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<col min="1"`, `<col outlineLevel="2" collapsed="1" min="1"`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<row r="1"`, `<row r="1" spans="1:11" outlineLevel="1" customFormat="1"`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<c r="B1"`, `<c r="B1" vendor="opaque"`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<c r="G1" t="str"><f>`, `<c r="G1" t="str"><f aca="1">`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<c r="I2" s="1"/>`, `<c r="I2" s="1"><extLst><ext uri="opaque"/></extLst></c>`, 1)
	parts["Meta/Strings.xml"] = strings.Replace(parts["Meta/Strings.xml"], `<sst `, `<sst vendor="opaque" `, 1)
	parts["Meta/Strings.xml"] = strings.Replace(parts["Meta/Strings.xml"], `</sst>`, `<extLst><ext uri="opaque"/></extLst></sst>`, 1)
	parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `</styleSheet>`, `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	for _, code := range []string{"SHEET_DECLARATION_ATTRIBUTES", "COLUMN_DIMENSION_EXTRAS", "ROW_DIMENSION_EXTRAS", "CELL_ATTRIBUTES", "FORMULA_ATTRIBUTES", "CELL_EXTENSIONS", "SHARED_STRING_TABLE_ATTRIBUTES", "SHARED_STRING_TABLE_OPAQUE_CONTENT", "STYLE_TABLE_OPAQUE_CONTENT"} {
		if !hasNativeWorkbookUnsupported(workbook, code) {
			t.Errorf("missing unsupported inventory %s: %#v", code, workbook.Unsupported)
		}
	}
	if workbook.Sheets[1].Editable || workbook.Sheets[1].RefusalCode == nil || findNativeCell(t, workbook, "7", "B1").Editable || findNativeCell(t, workbook, "7", "G1").Editable || findNativeCell(t, workbook, "7", "I2").Editable {
		t.Fatalf("unsafe sheet/cells remained editable: sheets=%#v B1=%#v G1=%#v I2=%#v", workbook.Sheets, findNativeCell(t, workbook, "7", "B1"), findNativeCell(t, workbook, "7", "G1"), findNativeCell(t, workbook, "7", "I2"))
	}
	if workbook.Source.Authority != "exact-package-bytes" {
		t.Fatalf("core unsupported authority is not explicit: %#v", workbook.Source)
	}
	for _, item := range workbook.Unsupported {
		if item.Code == "FORMULA_ATTRIBUTES" && (item.PartName == nil || *item.PartName != "Sheets/s1.xml" || item.Preservation != "preserve-exact") {
			t.Fatalf("modeled core unsupported record lost package authority: %#v", item)
		}
	}
}

func TestExtractNativeWorkbookV1SheetFormatGeometryFailsClosed(t *testing.T) {
	base := nativeWorkbookFixture(false)
	tests := []struct {
		name string
		edit func(map[string]string)
		want string
	}{
		{name: "missing height", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], ` defaultRowHeight="15"`, ``, 1)
		}, want: "requires defaultRowHeight"},
		{name: "negative height", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `defaultRowHeight="15"`, `defaultRowHeight="-1"`, 1)
		}, want: "defaultRowHeight"},
		{name: "bad boolean", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `zeroHeight="0"`, `zeroHeight="sometimes"`, 1)
		}, want: "zeroHeight"},
		{name: "duplicate", edit: func(parts map[string]string) {
			marker := `<sheetFormatPr baseColWidth="8" defaultColWidth="10.6640625" defaultRowHeight="15" customHeight="0" zeroHeight="0"/>`
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], marker, marker+marker, 1)
		}, want: "duplicated"},
		{name: "after columns", edit: func(parts map[string]string) {
			marker := `<sheetFormatPr baseColWidth="8" defaultColWidth="10.6640625" defaultRowHeight="15" customHeight="0" zeroHeight="0"/>`
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], marker, ``, 1)
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</cols>`, `</cols>`+marker, 1)
		}, want: "canonical schema order"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			parts := make(map[string]string, len(base))
			for name, value := range base {
				parts[name] = value
			}
			test.edit(parts)
			if _, err := ExtractNativeWorkbookV1(buildZip(t, parts)); err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("error = %v, want %q", err, test.want)
			}
		})
	}

	extra := nativeWorkbookFixture(false)
	extra["Sheets/s1.xml"] = strings.Replace(extra["Sheets/s1.xml"], `zeroHeight="0"`, `zeroHeight="0" outlineLevelRow="2"`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, extra))
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, item := range workbook.Unsupported {
		if item.Code == "SHEET_FORMAT_EXTRAS" && item.Capability == "dimensions" {
			found = true
		}
	}
	if !found {
		t.Fatal("unmodeled sheetFormatPr attributes were not inventoried")
	}
}

func TestExtractNativeWorkbookV1FormulaGroupIntegrityAndAbsentFollowers(t *testing.T) {
	tests := []struct {
		name string
		edit func(map[string]string)
		want string
	}{
		{name: "shared master missing ref", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], ` si="0" ref="H1:H2"`, ` si="0"`, 1)
		}, want: "master requires ref"},
		{name: "shared follower has ref", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<f t="shared" si="0"></f>`, `<f t="shared" si="0" ref="H1:H2"></f>`, 1)
		}, want: "follower cannot carry ref"},
		{name: "shared follower outside range", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `ref="H1:H2"`, `ref="H1:H1"`, 1)
		}, want: "outside master"},
		{name: "shared follower missing master", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `si="0"></f>`, `si="99"></f>`, 1)
		}, want: "missing master"},
		{name: "overlapping group ranges", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `ref="J1:J2"`, `ref="H1:J2"`, 1)
		}, want: "overlaps"},
		{name: "normal si", edit: func(parts map[string]string) {
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<f>CONCAT`, `<f si="3">CONCAT`, 1)
		}, want: "normal formula cannot carry si"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			parts := nativeWorkbookFixture(false)
			test.edit(parts)
			if workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts)); err == nil || workbook != nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("workbook=%#v err=%v, want %q", workbook, err, test.want)
			}
		})
	}
	absent := nativeWorkbookFixture(false)
	absent["Sheets/s1.xml"] = strings.Replace(absent["Sheets/s1.xml"], `<c r="H2"><f t="shared" si="0"></f><v>3</v></c>`, ``, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, absent))
	if err != nil {
		t.Fatal(err)
	}
	if workbook.Sheets[0].Editable || workbook.Sheets[0].RefusalCode == nil || *workbook.Sheets[0].RefusalCode != "FORMULA_GROUPS" {
		t.Fatalf("absent group follower could bypass sheet refusal: %#v", workbook.Sheets[0])
	}
}

func TestExtractNativeWorkbookV1BlankCachedLessFormulaAndDateLexicals(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Sheets/s2.xml"] = strings.Replace(parts["Sheets/s2.xml"], `</row>`, `<c r="B1"><f>1+1</f></c><c r="C1"/><c r="D1" t="n"/></row>`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	formula := findNativeCell(t, workbook, "9", "B1")
	blank, explicitNumber := findNativeCell(t, workbook, "9", "C1"), findNativeCell(t, workbook, "9", "D1")
	if formula.Formula == nil || formula.Formula.Cached != nil || blank.Value != nil || blank.OOXMLType != nil || explicitNumber.OOXMLType == nil || *explicitNumber.OOXMLType != "n" {
		t.Fatalf("blank/cached-less/absent-vs-explicit type fidelity failed: formula=%#v blank=%#v explicit=%#v", formula, blank, explicitNumber)
	}
	bad := nativeWorkbookFixture(false)
	bad["Sheets/s1.xml"] = strings.Replace(bad["Sheets/s1.xml"], `2026-08-27T12:00:00Z`, `2026-02-30`, 1)
	if _, err := ExtractNativeWorkbookV1(buildZip(t, bad)); err == nil || !strings.Contains(err.Error(), "ISO-8601") {
		t.Fatalf("invalid date lexical was accepted: %v", err)
	}
	hexadecimal := nativeWorkbookFixture(false)
	hexadecimal["Sheets/s1.xml"] = strings.Replace(hexadecimal["Sheets/s1.xml"], `001.2300`, `0x1p2`, 1)
	if _, err := ExtractNativeWorkbookV1(buildZip(t, hexadecimal)); err == nil || !strings.Contains(err.Error(), "numeric lexical") {
		t.Fatalf("non-XML hexadecimal numeric lexical was accepted: %v", err)
	}
}

func TestMarkNativeFormulaGroupsIsBoundedAndSubquadratic(t *testing.T) {
	const groups = 10_000
	sheet := NativeWorkbookSheetV1{ID: "1", PartName: "sheet.xml", Cells: make([]NativeWorkbookCellV1, 0, groups), Editable: true}
	for row := 0; row < groups; row++ {
		ref := cellReference(row, 0)
		index := uint32(row)
		sheet.Cells = append(sheet.Cells, NativeWorkbookCellV1{Row: row, Column: 0, Ref: ref, Formula: &NativeWorkbookFormulaV1{Text: "1", Type: "shared", Ref: &ref, SharedIndex: &index}, Editable: true})
	}
	extractor := nativeWorkbookExtractor{unsupported: []NativeWorkbookUnsupportedV1{}, unsupportedKeys: map[string]bool{}}
	if err := extractor.markNativeFormulaGroupFollowers(&sheet); err != nil {
		t.Fatal(err)
	}
	if len(extractor.unsupported) != groups || sheet.Editable {
		t.Fatalf("bounded formula inventory mismatch: unsupported=%d editable=%v", len(extractor.unsupported), sheet.Editable)
	}
	over := sheet
	extraRef, extraIndex := cellReference(groups, 0), uint32(groups)
	over.Cells = append(over.Cells, NativeWorkbookCellV1{Row: groups, Column: 0, Ref: extraRef, Formula: &NativeWorkbookFormulaV1{Text: "1", Type: "shared", Ref: &extraRef, SharedIndex: &extraIndex}, Editable: true})
	extractor = nativeWorkbookExtractor{unsupported: []NativeWorkbookUnsupportedV1{}, unsupportedKeys: map[string]bool{}}
	if err := extractor.markNativeFormulaGroupFollowers(&over); err == nil || !strings.Contains(err.Error(), "inventory exceeds") {
		t.Fatalf("formula inventory bound error = %v", err)
	}
}

func TestExtractNativeWorkbookV1XMLDeclarationsRootsAndFixedAttributes(t *testing.T) {
	valid := nativeWorkbookFixture(false)
	valid["Book/Workbook.xml"] = strings.Replace(valid["Book/Workbook.xml"], `<?xml version="1.0"?>`, `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>`, 1)
	if _, err := ExtractNativeWorkbookV1(buildZip(t, valid)); err != nil {
		t.Fatalf("standards-valid declaration was rejected: %v", err)
	}
	tests := []struct {
		name string
		edit func(map[string]string)
		want string
	}{
		{name: "declaration order", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `<?xml version="1.0"?>`, `<?xml encoding="UTF-8" version="1.0"?>`, 1)
		}, want: "invalid"},
		{name: "late declaration", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `<workbook`, `<workbook`, 1) + `<?xml version="1.0"?>`
		}, want: "processing instruction"},
		{name: "multiple roots", edit: func(parts map[string]string) {
			parts["Sheets/s2.xml"] += `<worksheet xmlns="` + spreadsheetMLTransitional + `"/>`
		}, want: "multiple root"},
		{name: "workbook direct text", edit: func(parts map[string]string) {
			parts["Book/Workbook.xml"] = strings.Replace(parts["Book/Workbook.xml"], `<sheets>`, `unsafe<sheets>`, 1)
		}, want: "unsupported direct text"},
		{name: "invalid xml space", edit: func(parts map[string]string) {
			parts["Meta/Strings.xml"] = strings.Replace(parts["Meta/Strings.xml"], `<t>Plain`, `<t xml:space="sometimes">Plain`, 1)
		}, want: "invalid xml:space"},
		{name: "content types root attribute", edit: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `<Types xmlns=`, `<Types vendor="x" xmlns=`, 1)
		}, want: "unexpected semantic attribute"},
		{name: "relationship child attribute", edit: func(parts map[string]string) {
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `<Relationship Id=`, `<Relationship vendor="x" Id=`, 1)
		}, want: "unexpected semantic attribute"},
		{name: "duplicate relationship attribute", edit: func(parts map[string]string) {
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `Id="rOffice"`, `Id="rOffice" Id="again"`, 1)
		}, want: "duplicate"},
		{name: "unicode MIME fold alias", edit: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], strings.ToUpper(nativeWorkbookContentType), "APPLICATİON/VND.OPENXMLFORMATS-OFFICEDOCUMENT.SPREADSHEETML.SHEET.MAIN+XML", 1)
		}, want: "expected"},
		{name: "oversized relationship field", edit: func(parts map[string]string) {
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `Id="rOffice"`, `Id="`+strings.Repeat("x", maxRoutingRelationshipIDLength+1)+`"`, 1)
		}, want: "resource bound"},
		{name: "opposing relationship outside workbook rels", edit: func(parts map[string]string) {
			parts["Sheets/_rels/s1.xml.rels"] = strings.Replace(parts["Sheets/_rels/s1.xml.rels"], "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink", 1)
		}, want: "opposing Strict/Transitional"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			parts := nativeWorkbookFixture(false)
			test.edit(parts)
			if workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts)); err == nil || workbook != nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("workbook=%#v err=%v, want %q", workbook, err, test.want)
			}
		})
	}
}

func nativeExtractorForResourceTest(t *testing.T) *nativeWorkbookExtractor {
	t.Helper()
	pkg, err := openNativeWorkbookPackage(buildZip(t, nativeWorkbookFixture(false)))
	if err != nil {
		t.Fatal(err)
	}
	readBytes := func(name string) ([]byte, bool) {
		data, found := pkg.files[name]
		return data, found
	}
	workbook, err := locateWorkbookPartBytes(pkg.index, readBytes)
	if err != nil {
		t.Fatal(err)
	}
	return &nativeWorkbookExtractor{
		pkg: pkg, workbook: workbook, namespace: spreadsheetMLTransitional, relNamespace: officeRelNamespaceTransitional,
		modeled: map[string]bool{}, unsupported: []NativeWorkbookUnsupportedV1{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{},
	}
}

func TestNativeExtractionCumulativeRelationshipAndXMLBudgets(t *testing.T) {
	extractor := nativeExtractorForResourceTest(t)
	extractor.relationshipCount = NativeXLSXMaxRelationships
	if err := extractor.validateAllRelationships(); err == nil || !strings.Contains(err.Error(), "cumulative limit") {
		t.Fatalf("cumulative relationship limit error = %v", err)
	}
	extractor = nativeExtractorForResourceTest(t)
	extractor.xmlTokens = NativeXLSXMaxXMLTokens - 1
	if err := extractor.claimCoreXML(extractor.workbook.part); err == nil || !strings.Contains(err.Error(), "cumulative") {
		t.Fatalf("cumulative XML token limit error = %v", err)
	}
	extractor = nativeExtractorForResourceTest(t)
	extractor.xmlElements = NativeXLSXMaxXMLElements
	if err := extractor.claimCoreXML(extractor.workbook.part); err == nil || !strings.Contains(err.Error(), "cumulative") {
		t.Fatalf("cumulative XML element limit error = %v", err)
	}
}

func TestExtractNativeWorkbookV1RejectsZIPSizeHeaderMismatch(t *testing.T) {
	archive := buildZip(t, nativeWorkbookFixture(false))
	hostile := bytes.Clone(archive)
	signature := []byte{'P', 'K', 1, 2}
	offset := bytes.Index(hostile, signature)
	if offset < 0 || offset+28 > len(hostile) {
		t.Fatal("central-directory entry not found")
	}
	declared := binary.LittleEndian.Uint32(hostile[offset+24 : offset+28])
	binary.LittleEndian.PutUint32(hostile[offset+24:offset+28], declared+1)
	if workbook, err := ExtractNativeWorkbookV1(hostile); err == nil || workbook != nil {
		t.Fatalf("ZIP size-header mismatch was accepted: %#v", workbook)
	}
}

func TestUnsupportedPartCapabilityUsesASCIIOnlyClassification(t *testing.T) {
	tests := []struct {
		name, contentType, partName, code, capability string
	}{
		{name: "ASCII case folds", contentType: "APPLICATION/VND.OPENXMLFORMATS-OFFICEDOCUMENT.DRAWINGML.CHART+XML", partName: "Charts/C1.XML", code: "CHART_CONTENT", capability: "charts"},
		{name: "Kelvin sign is not ASCII k", contentType: "application/xml", partName: "Links/externalLinK.xml", code: "OPAQUE_PACKAGE_PART", capability: "opaque-parts"},
		{name: "dotted I is not ASCII i", contentType: "application/xml", partName: "Media/İMAGE.bin", code: "OPAQUE_PACKAGE_PART", capability: "opaque-parts"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			code, capability := unsupportedPartCapability(test.contentType, test.partName)
			if code != test.code || capability != test.capability {
				t.Fatalf("classification = %s/%s, want %s/%s", code, capability, test.code, test.capability)
			}
		})
	}
}

func TestDecodeSpreadsheetStringPreservesEscapedEscapeSignatures(t *testing.T) {
	for input, want := range map[string]string{
		`Plain_x0020_Text`:    "Plain Text",
		`_x005F_x0041_`:       `_x0041_`,
		`Smile_xD83D__xDE00_`: "Smile😀",
	} {
		got, err := decodeSpreadsheetString(input)
		if err != nil || got != want {
			t.Errorf("decodeSpreadsheetString(%q) = %q, %v; want %q", input, got, err, want)
		}
	}
	if _, err := decodeSpreadsheetString(`bad_xD83D_`); err == nil || !strings.Contains(err.Error(), "unpaired") {
		t.Fatalf("unpaired surrogate error = %v", err)
	}
}

func TestExtractEmptyInlineStrCellIsBlankString(t *testing.T) {
	for _, cellXML := range []string{
		`<c r="B19" t="inlineStr"/>`,
		`<c r="B19" t="inlineStr"></c>`,
		`<c r="B19" t="inlineStr"><is/></c>`,
		`<c r="B19" t="inlineStr"><is><t/></is></c>`,
		`<c r="B19" t="inlineStr"><is><t></t></is></c>`,
	} {
		t.Run(cellXML, func(t *testing.T) {
			parts := nativeWorkbookFixture(false)
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</sheetData>`, `<row r="19">`+cellXML+`</row></sheetData>`, 1)
			workbook, err := ExtractNativeWorkbookV2(buildZip(t, parts))
			if err != nil {
				t.Fatal(err)
			}
			var found *NativeWorkbookCellV2
			for _, sheet := range workbook.Sheets {
				for i := range sheet.Cells {
					if sheet.Cells[i].Ref == "B19" {
						found = &sheet.Cells[i]
					}
				}
			}
			if found == nil || found.Value == nil || found.Value.Kind != "string" || found.Value.Storage != "inline" || found.Value.Text == nil || *found.Value.Text != "" {
				t.Fatalf("empty inlineStr was not a blank string: %#v", found)
			}
		})
	}
}

func TestExtractInlineStrWithValueChildStillFailsClosed(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</sheetData>`, `<row r="19"><c r="B19" t="inlineStr"><v>x</v></c></row></sheetData>`, 1)
	if _, err := ExtractNativeWorkbookV2(buildZip(t, parts)); err == nil || !strings.Contains(err.Error(), "inlineStr") {
		t.Fatalf("inlineStr + <v> must remain fail-closed, err=%v", err)
	}
}
