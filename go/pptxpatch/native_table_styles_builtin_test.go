package pptxpatch

import (
	"strings"
	"testing"
)

const nativeTestMediumStyle2 = "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}"
const nativeTestMediumStyle2Accent1 = "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"

// Text-free cell as PowerPoint writes it for a freshly inserted styled table.
func nativeStyledTableCellXML() string {
	return `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="tr-TR"/></a:p></a:txBody><a:tcPr/></a:tc>`
}

func nativeStyledTableXML(id int, styleID, flags string, rows, columns int) string {
	cells := make([][]string, 0, rows)
	widths := make([]int64, 0, columns)
	heights := make([]int64, 0, rows)
	for column := 0; column < columns; column++ {
		widths = append(widths, 1016000)
	}
	for row := 0; row < rows; row++ {
		heights = append(heights, 370840)
		line := make([]string, 0, columns)
		for column := 0; column < columns; column++ {
			line = append(line, nativeStyledTableCellXML())
		}
		cells = append(cells, line)
	}
	table := nativeExactTableGraphicFrameXML(id, "Styled table", widths, heights, cells, "")
	return strings.Replace(table, `<a:tblPr/>`, `<a:tblPr `+flags+`><a:tableStyleId>`+styleID+`</a:tableStyleId></a:tblPr>`, 1)
}

// The benchmark corpus stores the no-group lock and the PowerPoint modId
// extension on styled tables; the exact projection refuses them.
func nativeStyledTableWithLocksXML(table string) string {
	return strings.Replace(table, `<p:cNvGraphicFramePr/><p:nvPr/>`, `<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1084916369"/></p:ext></p:extLst></p:nvPr>`, 1)
}

// nativeStyledTableFixture pairs the slide children with the exact fixture
// theme (dk1=000000, lt1=FFFFFF, accent1=2F6FED) that catalog colors resolve
// against; the default fixture theme has no color scheme.
func nativeStyledTableFixture(t *testing.T, strict bool, children string, change func(map[string]string)) []byte {
	t.Helper()
	drawing := nsDrawingTransitional
	if strict {
		drawing = nsDrawingStrict
	}
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/themes/theme.xml"] = nativeExactThemeXML(drawing)
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, children+`</p:spTree>`, 1)
		if change != nil {
			change(parts)
		}
	}})
}

func nativeStyledTableElement(t *testing.T, deck NativePPTXDeck) NativeElement {
	t.Helper()
	if len(deck.Slides) != 1 {
		t.Fatalf("expected one slide: %#v", deck.Slides)
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindTable {
			return element
		}
	}
	t.Fatalf("styled table was not projected: %#v", deck.Slides[0].Compatibility)
	return NativeElement{}
}

func nativeTableFills(element NativeElement) [][]string {
	fills := make([][]string, 0, len(element.Table.Rows))
	for _, row := range element.Table.Rows {
		line := make([]string, 0, len(row))
		for _, cell := range row {
			if cell.Fill == nil {
				line = append(line, "")
			} else {
				line = append(line, *cell.Fill)
			}
		}
		fills = append(fills, line)
	}
	return fills
}

func TestExtractNativePPTXBuiltinTableStylePreviewBandsAndHeader(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			table := nativeStyledTableWithLocksXML(nativeStyledTableXML(3, nativeTestMediumStyle2, `firstRow="1" bandRow="1"`, 4, 2))
			deck, err := ExtractNativePPTX(nativeStyledTableFixture(t, strict, table, nil), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract styled table: %v", err)
			}
			element := nativeStyledTableElement(t, deck)
			if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(element.Passthrough) != 0 || element.Source == nil || element.Source.ObjectID != "cNvPr-3" || element.GraphicFrameLayout != nil {
				t.Fatalf("styled table must be a parsed preserve-only projection: %#v", element)
			}
			if !nativeDiagnosticsContain(element.Compatibility.Diagnostics, nativeBuiltinTableStylePreviewCode) || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, nativeBuiltinTableStylePreviewCode) {
				t.Fatalf("preview diagnostic missing: %#v", element.Compatibility)
			}
			for _, diagnostic := range element.Compatibility.Diagnostics {
				if diagnostic.Code == nativeBuiltinTableStylePreviewCode && (diagnostic.Severity != NativeDiagnosticSeverityWarning || !strings.Contains(diagnostic.Message, nativeBuiltinTableStylePolicy) || !strings.Contains(diagnostic.Message, nativeTestMediumStyle2)) {
					t.Fatalf("preview diagnostic must declare the policy and style: %#v", diagnostic)
				}
			}
			// dk1=000000 header, dk1 tint 40% / 20% bands in linear sRGB.
			want := [][]string{{"000000", "000000"}, {"CBCBCB", "CBCBCB"}, {"E7E7E7", "E7E7E7"}, {"CBCBCB", "CBCBCB"}}
			if got := nativeTableFills(element); strings.Join(flattenNativeTestFills(got), ",") != strings.Join(flattenNativeTestFills(want), ",") {
				t.Fatalf("fills = %v, want %v", got, want)
			}
			for _, row := range element.Table.Rows {
				for _, cell := range row {
					if cell.Text == nil || *cell.Text != "" || cell.Paragraphs != nil || cell.TextBody != nil || cell.Align != nil {
						t.Fatalf("styled cells must be text-free legacy cells: %#v", cell)
					}
					if cell.Border == nil || cell.Border.Color != "FFFFFF" || cell.Border.WidthEMU == nil || *cell.Border.WidthEMU != 12700 {
						t.Fatalf("styled cells must carry the lt1 1pt border: %#v", cell.Border)
					}
				}
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("styled deck is invalid: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("marshal styled deck: %v", err)
			}
			decoded, err := DecodeNativePPTXJSON(encoded)
			if err != nil || len(ValidateNativePPTX(decoded)) != 0 {
				t.Fatalf("styled deck JSON round trip: err=%v issues=%#v", err, ValidateNativePPTX(decoded))
			}
		})
	}
}

func flattenNativeTestFills(fills [][]string) []string {
	result := []string{}
	for _, row := range fills {
		result = append(result, strings.Join(row, "|"))
	}
	return result
}

func TestExtractNativePPTXBuiltinTableStyleAccentColumnsAndLastRow(t *testing.T) {
	t.Parallel()
	table := nativeStyledTableXML(3, nativeTestMediumStyle2Accent1, `firstCol="1" bandCol="1" lastRow="1"`, 2, 3)
	deck, err := ExtractNativePPTX(nativeStyledTableFixture(t, false, table, nil), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract accent styled table: %v", err)
	}
	element := nativeStyledTableElement(t, deck)
	// accent1=2F6FED: first column and last row solid, vertical bands at 40%/20%.
	want := [][]string{{"2F6FED", "CDD5F8", "E8EBFC"}, {"2F6FED", "2F6FED", "2F6FED"}}
	if got := nativeTableFills(element); strings.Join(flattenNativeTestFills(got), ",") != strings.Join(flattenNativeTestFills(want), ",") {
		t.Fatalf("fills = %v, want %v", got, want)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("accent styled deck is invalid: %#v", issues)
	}
}

func TestExtractNativePPTXBuiltinTableStyleWithoutSwitchesUsesWholeTable(t *testing.T) {
	t.Parallel()
	table := nativeStyledTableXML(3, nativeTestMediumStyle2Accent1, `firstRow="0" bandRow="0"`, 2, 2)
	deck, err := ExtractNativePPTX(nativeStyledTableFixture(t, false, table, nil), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract whole-table styled table: %v", err)
	}
	element := nativeStyledTableElement(t, deck)
	for _, row := range nativeTableFills(element) {
		for _, fill := range row {
			if fill != "E8EBFC" {
				t.Fatalf("whole-table fill = %q, want accent1 tint 20%%", fill)
			}
		}
	}
}

func TestExtractNativePPTXBuiltinTableStyleRefusals(t *testing.T) {
	t.Parallel()
	base := nativeStyledTableXML(3, nativeTestMediumStyle2, `firstRow="1" bandRow="1"`, 2, 2)
	textCell := `<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Inherited</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>`
	cases := []struct {
		name     string
		mutate   func(string) string
		wantCode string
	}{
		{name: "unknown-guid", mutate: func(value string) string {
			return strings.Replace(value, nativeTestMediumStyle2, "{5940675A-B579-460E-94D1-54222C63F5DA}", 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "malformed-guid", mutate: func(value string) string {
			return strings.Replace(value, nativeTestMediumStyle2, "073A0DAA-6AF3-43AB-8588-CEC1D06C72B9", 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "inline-table-fill", mutate: func(value string) string {
			return strings.Replace(value, `<a:tableStyleId>`, `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:tableStyleId>`, 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "rtl", mutate: func(value string) string {
			return strings.Replace(value, `firstRow="1"`, `rtl="1" firstRow="1"`, 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "non-boolean-switch", mutate: func(value string) string {
			return strings.Replace(value, `bandRow="1"`, `bandRow="yes"`, 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "cell-text", mutate: func(value string) string {
			return strings.Replace(value, nativeStyledTableCellXML(), textCell, 1)
		}, wantCode: nativeBuiltinTableStyleTextCode},
		{name: "explicit-cell-fill", mutate: func(value string) string {
			return strings.Replace(value, `<a:tcPr/>`, `<a:tcPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:tcPr>`, 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "merged-cell", mutate: func(value string) string {
			return strings.Replace(value, `<a:tc>`, `<a:tc gridSpan="2">`, 1)
		}, wantCode: "pptx.table-merge-unavailable"},
	}
	for _, testCase := range cases {
		testCase := testCase
		t.Run(testCase.name, func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeStyledTableFixture(t, false, testCase.mutate(base), nil), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("refused styled table aborted extract: %v", err)
			}
			if len(deck.Slides) != 1 {
				t.Fatalf("slide was not emitted: %#v", deck.Slides)
			}
			for _, element := range deck.Slides[0].Elements {
				if element.Kind == NativeElementKindTable {
					t.Fatalf("refused styled table leaked a projection: %#v", element)
				}
			}
			if !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, testCase.wantCode) {
				t.Fatalf("expected %s refusal: %#v", testCase.wantCode, deck.Slides[0].Compatibility)
			}
			opaque := false
			for _, passthrough := range deck.Slides[0].Passthrough {
				if passthrough.OwnerPart == "relocated/slides/slide-a.xml" {
					opaque = true
				}
			}
			if !opaque {
				t.Fatalf("refused styled table must stay an opaque passthrough: %#v", deck.Slides[0].Passthrough)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("refused deck is invalid: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXBuiltinTableStyleRefusesExplicitPackageDefinition(t *testing.T) {
	t.Parallel()
	table := nativeStyledTableXML(3, nativeTestMediumStyle2, `firstRow="1" bandRow="1"`, 2, 2)
	for _, defined := range []bool{false, true} {
		defined := defined
		t.Run(map[bool]string{false: "empty-style-list", true: "explicit-definition"}[defined], func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{extraParts: []nativeExtractZipPart{{name: "relocated/styles/table.xml", data: "placeholder"}}, mutate: func(parts map[string]string) {
				parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, table+`</p:spTree>`, 1)
				list := `<a:tblStyleLst xmlns:a="` + nsDrawingTransitional + `" def="` + nativeTestMediumStyle2Accent1 + `"></a:tblStyleLst>`
				if defined {
					list = `<a:tblStyleLst xmlns:a="` + nsDrawingTransitional + `" def="` + nativeTestMediumStyle2 + `">` + strings.Replace(paintTestStyle(), "{01234567-89AB-CDEF-0123-456789ABCDEF}", nativeTestMediumStyle2, 1) + `</a:tblStyleLst>`
				}
				parts["relocated/styles/table.xml"] = list
				parts["relocated/_rels/deck.xml.rels"] = strings.Replace(parts["relocated/_rels/deck.xml.rels"], `</Relationships>`, `<Relationship Id="rIdStyles" Type="`+nsOfficeRelsTransitional+`/tableStyles" Target="styles/table.xml"/></Relationships>`, 1)
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/relocated/styles/table.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/></Types>`, 1)
			}})
			deck, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract with table styles part: %v", err)
			}
			projected := false
			for _, element := range deck.Slides[0].Elements {
				if element.Kind == NativeElementKindTable {
					projected = true
				}
			}
			if defined {
				if projected || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.table-style-unavailable") {
					t.Fatalf("explicit package definition must refuse the catalog preview: %#v", deck.Slides[0])
				}
			} else if !projected {
				t.Fatalf("empty style list must not block the catalog preview: %#v", deck.Slides[0].Compatibility)
			}
		})
	}
}

// The exact projection used to lose a locked table entirely. It now keeps it
// and marks it read-only instead: neither a:graphicFrameLocks nor the p14:modId
// extension travels on the wire, so the projection may never be editable, but
// neither can move a pixel, so neither may cost the table.
func TestExtractNativePPTXExactTableKeepsLocksAndModificationIdentifiersReadOnly(t *testing.T) {
	t.Parallel()
	table := nativeStyledTableWithLocksXML(nativeExactTableGraphicFrameXML(3, "Locked exact table", []int64{500000, 500000}, []int64{500000}, [][]string{{
		nativeExactTableCellXML("One", "l", "FFFFFF"), nativeExactTableCellXML("Two", "r", "EEEEEE"),
	}}, ""))
	deck, err := ExtractNativePPTX(nativeStyledTableFixture(t, false, table, nil), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract locked exact table: %v", err)
	}
	var projected *NativeElement
	for index := range deck.Slides[0].Elements {
		if deck.Slides[0].Elements[index].Kind == NativeElementKindTable {
			projected = &deck.Slides[0].Elements[index]
		}
	}
	if projected == nil {
		t.Fatalf("locked exact table was refused: %#v", deck.Slides[0].Elements)
	}
	if projected.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("locked exact table stayed editable: %s", projected.Compatibility.Status)
	}
	if !nativeDiagnosticsContain(projected.Compatibility.Diagnostics, nativeTableNonVisualPreservedCode) {
		t.Fatalf("preserved nonvisual metadata was not disclosed: %#v", projected.Compatibility)
	}
	// The read-only status is enforced, not merely conventional.
	editable := deck
	editable.Slides[0].Elements[len(deck.Slides[0].Elements)-1].Compatibility.Status = NativeCompatibilityStatusEditable
	found := false
	for _, issue := range ValidateNativePPTX(editable) {
		if issue.Code == "native.tableNonVisualPreserved" {
			found = true
		}
	}
	if !found {
		t.Fatalf("an editable locked table validated: %#v", ValidateNativePPTX(editable))
	}
}

func TestValidateNativePPTXBuiltinTableStylePreviewMustStayReadOnly(t *testing.T) {
	t.Parallel()
	table := nativeStyledTableXML(3, nativeTestMediumStyle2, `firstRow="1" bandRow="1"`, 2, 2)
	deck, err := ExtractNativePPTX(nativeStyledTableFixture(t, false, table, nil), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract styled table: %v", err)
	}
	for index := range deck.Slides[0].Elements {
		if deck.Slides[0].Elements[index].Kind == NativeElementKindTable {
			deck.Slides[0].Elements[index].Compatibility.Status = NativeCompatibilityStatusEditable
		}
	}
	deck.Slides[0].Compatibility.Status = NativeCompatibilityStatusEditable
	found := false
	for _, issue := range ValidateNativePPTX(deck) {
		if issue.Code == "native.tableStylePreview" {
			found = true
		}
	}
	if !found {
		t.Fatalf("editable styled table must be rejected: %#v", ValidateNativePPTX(deck))
	}
}

func TestExtractNativePPTXBuiltinTableStyleRefusesThemeWithoutColorScheme(t *testing.T) {
	t.Parallel()
	// The default fixture theme carries no clrScheme, so catalog colors cannot
	// resolve and the styled table must stay opaque instead of guessing.
	table := nativeStyledTableXML(3, nativeTestMediumStyle2, `firstRow="1" bandRow="1"`, 2, 2)
	deck, err := ExtractNativePPTX(nativeTableFixture(t, false, table), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract without theme colors: %v", err)
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindTable {
			t.Fatalf("styled table without theme colors leaked a projection: %#v", element)
		}
	}
	if !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.table-style-unavailable") {
		t.Fatalf("missing theme snapshot must refuse: %#v", deck.Slides[0].Compatibility)
	}
}

func TestApplyNativeLinearSRGBTint(t *testing.T) {
	t.Parallel()
	cases := []struct {
		color string
		tint  int64
		want  string
	}{
		{"000000", 40000, "CBCBCB"},
		{"000000", 20000, "E7E7E7"},
		{"2F6FED", 40000, "CDD5F8"},
		{"2F6FED", 20000, "E8EBFC"},
		{"FFFFFF", 40000, "FFFFFF"},
		{"4F81BD", 100000, "4F81BD"},
		{"4F81BD", 0, "FFFFFF"},
	}
	for _, testCase := range cases {
		color, err := parseNativeSRGBHex(testCase.color)
		if err != nil {
			t.Fatal(err)
		}
		if got := applyNativeLinearSRGBTint(color, testCase.tint).hex(); got != testCase.want {
			t.Fatalf("tint(%s, %d) = %s, want %s", testCase.color, testCase.tint, got, testCase.want)
		}
	}
}

func TestNativeBuiltinTableStyleRegionPrecedence(t *testing.T) {
	t.Parallel()
	style := nativeBuiltinTableStyles[nativeTestMediumStyle2Accent1]
	resolved := &nativeResolvedBuiltinTableStyle{style: style, flags: nativeTableStyleFlags{firstRow: true, lastRow: true, firstCol: true, lastCol: true, bandRow: true, bandCol: true}}
	rows, columns := 4, 4
	// Rows beat columns beat bands beat the whole table; bands start after the header.
	if resolved.regionFill(0, 0, rows, columns) != style.firstRow || resolved.regionFill(0, 3, rows, columns) != style.firstRow {
		t.Fatal("header row must win over column parts")
	}
	if resolved.regionFill(3, 1, rows, columns) != style.lastRow {
		t.Fatal("last row must win over bands")
	}
	if resolved.regionFill(1, 0, rows, columns) != style.firstCol || resolved.regionFill(1, 3, rows, columns) != style.lastCol {
		t.Fatal("first/last column must win over bands")
	}
	if resolved.regionFill(1, 1, rows, columns) != style.band1V || resolved.regionFill(1, 2, rows, columns) != style.band2V {
		t.Fatal("vertical bands must start after the first column")
	}
	resolved.flags = nativeTableStyleFlags{firstRow: true, bandRow: true}
	if resolved.regionFill(1, 0, rows, columns) != style.band1H || resolved.regionFill(2, 0, rows, columns) != style.band2H || resolved.regionFill(3, 0, rows, columns) != style.band1H {
		t.Fatal("horizontal bands must alternate from the first body row")
	}
	resolved.flags = nativeTableStyleFlags{}
	if resolved.regionFill(0, 0, rows, columns) != style.wholeTbl || resolved.regionFill(3, 3, rows, columns) != style.wholeTbl {
		t.Fatal("no switches must fall back to the whole table part")
	}
}
