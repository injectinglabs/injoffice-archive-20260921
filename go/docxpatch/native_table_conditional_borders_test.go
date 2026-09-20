package docxpatch

import (
	"strconv"
	"strings"
	"testing"
)

func conditionalCellBorders(resolved *NativeResolvedLayoutInputV1, doc *NativeDocumentV1) map[[2]int]NativeTableBordersV1 {
	byID := map[string]NativeTableBordersV1{}
	for _, entry := range resolved.Tables[0].ConditionalCellBorders {
		byID[entry.CellID] = entry.Borders
	}
	out := map[[2]int]NativeTableBordersV1{}
	for r, row := range doc.Body.Blocks[0].Table.Rows {
		for c, cell := range row.Cells {
			out[[2]int{r, c}] = byID[cell.ID]
		}
	}
	return out
}

func describeBorder(border *NativeTableBorderV1) string {
	if border == nil {
		return "-"
	}
	if border.Style != "single" {
		return "none"
	}
	color := ""
	if border.ColorRGB != nil {
		color = *border.ColorRGB
	}
	return strings.TrimSpace(color + " " + strconv.FormatInt(border.SizeEighthPoints, 10))
}

// The fixture's chain states auto-coloured 2.25 pt top and bottom rules and a
// first-column region whose top edge is nil and bottom edge 2.25 pt auto --
// the "Medium Shading 2" shape whose Word raster shows a black top rule over
// every column but the first, and a black bottom rule under all of them.
func TestNativeConditionalTableStyleResolvesRegionBordersAsWordPaints(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `">` + strings.Replace(conditionalTableStyleFixture,
		`<w:tblBorders><w:top w:val="single" w:sz="18" w:space="0" w:color="auto"/></w:tblBorders>`,
		`<w:tblBorders><w:top w:val="single" w:sz="18" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="18" w:space="0" w:color="auto"/></w:tblBorders>`, 1) + `</w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = conditionalTableDocument(`<w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>`, 3, 3)
	data := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	borders := conditionalCellBorders(resolved, doc)
	// The first-column region states only its top (nil) and bottom (2.25 pt
	// auto) edges; every other edge falls back to the table, which states no
	// side or inside rules.
	want := map[[2]int][4]string{
		// top, right, bottom, left
		{0, 0}: {"none", "-", "-", "-"},
		{0, 1}: {"000000 18", "-", "-", "-"},
		{0, 2}: {"000000 18", "-", "-", "-"},
		{1, 0}: {"-", "-", "-", "-"},
		{1, 1}: {"-", "-", "-", "-"},
		{2, 0}: {"-", "-", "000000 18", "-"},
		{2, 1}: {"-", "-", "000000 18", "-"},
		{2, 2}: {"-", "-", "000000 18", "-"},
	}
	for key, edges := range want {
		got := borders[key]
		actual := [4]string{describeBorder(got.Top), describeBorder(got.Right), describeBorder(got.Bottom), describeBorder(got.Left)}
		if actual != edges {
			t.Fatalf("cell %v edges %v want %v", key, actual, edges)
		}
	}
	// The exact tier still refuses: the automatic colour and the region
	// borders both leave their diagnostics.
	if !hasResolutionDiagnostic(resolved, "TABLE_STYLE_EFFECTS_PRESERVED") || !hasResolutionDiagnostic(resolved, "CONDITIONAL_TABLE_STYLE_PRESERVED") {
		t.Fatalf("approximate border resolution silenced the exact tier's refusal: %#v", resolved.Diagnostics)
	}
}

// Grid Table 2's shape: a first-row region whose bottom edge is heavier than
// the table's inside rule and whose vertical inside edges are nil, over a
// table that states thin theme-tinted rules everywhere, plus an unselected
// last-row region with a double border this tier cannot read.
func TestNativeConditionalTableStyleResolvesSharedEdgesByWeight(t *testing.T) {
	style := `<w:style w:type="table" w:styleId="Shaded"><w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/>` +
		`<w:tblBorders><w:top w:val="single" w:sz="2" w:space="0" w:color="F4B083" w:themeColor="accent2" w:themeTint="99"/><w:bottom w:val="single" w:sz="2" w:space="0" w:color="F4B083" w:themeColor="accent2" w:themeTint="99"/><w:insideH w:val="single" w:sz="2" w:space="0" w:color="F4B083" w:themeColor="accent2" w:themeTint="99"/><w:insideV w:val="single" w:sz="2" w:space="0" w:color="F4B083" w:themeColor="accent2" w:themeTint="99"/></w:tblBorders></w:tblPr>` +
		`<w:tblStylePr w:type="firstRow"><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="F4B083" w:themeColor="accent2" w:themeTint="99"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/></w:tcPr></w:tblStylePr>` +
		`<w:tblStylePr w:type="lastRow"><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="2" w:space="0" w:color="F4B083"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style>`
	for _, test := range []struct {
		name string
		look string
		rows bool
	}{
		{"unselected double border does not block", `<w:tblLook w:val="00A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>`, true},
		{"selected double border refuses all region borders", `<w:tblLook w:val="00E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>`, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `">` + style + `</w:styles>`)
			parts["word/document.xml"] = conditionalTableDocument(test.look, 3, 2)
			data := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if !test.rows {
				if len(resolved.Tables[0].ConditionalCellBorders) != 0 {
					t.Fatalf("unreadable selected region borders were guessed: %#v", resolved.Tables[0].ConditionalCellBorders)
				}
				return
			}
			borders := conditionalCellBorders(resolved, doc)
			want := map[[2]int][4]string{
				{0, 0}: {"none", "none", "F4B083 12", "-"},
				{0, 1}: {"none", "-", "F4B083 12", "none"},
				{1, 0}: {"F4B083 12", "F4B083 2", "F4B083 2", "-"},
				{1, 1}: {"F4B083 12", "-", "F4B083 2", "F4B083 2"},
				{2, 1}: {"F4B083 2", "-", "F4B083 2", "F4B083 2"},
			}
			for key, edges := range want {
				got := borders[key]
				actual := [4]string{describeBorder(got.Top), describeBorder(got.Right), describeBorder(got.Bottom), describeBorder(got.Left)}
				if actual != edges {
					t.Fatalf("cell %v edges %v want %v", key, actual, edges)
				}
			}
		})
	}
}

func TestNativeConditionalTableStyleLeavesMergedAndDirectlyBorderedCellsToTheTable(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `">` + conditionalTableStyleFixture + `</w:styles>`
	for name, cell := range map[string]string{
		"vertical merge": `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`,
		"direct borders": `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="FF0000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`,
	} {
		t.Run(name, func(t *testing.T) {
			parts := resolvedStylesTestParts(styles)
			document := conditionalTableDocument(`<w:tblLook w:val="01E0"/>`, 2, 2)
			parts["word/document.xml"] = strings.Replace(document, `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`, cell, 1)
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if len(resolved.Tables[0].ConditionalCellBorders) != 0 {
				t.Fatalf("cell borders were resolved over %s: %#v", name, resolved.Tables[0].ConditionalCellBorders)
			}
			if len(resolved.Tables[0].ConditionalCellShading) == 0 {
				t.Fatalf("fills were dropped with the borders")
			}
		})
	}
}
