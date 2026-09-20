package docxpatch

import (
	"strings"
	"testing"
)

// The style below is the "Medium Shading 2" chain of conditionalstyles-tbllook.docx
// with every region carrying a distinct fill, which is what makes Word's own
// raster of that file the oracle for the precedence order implemented here.
const conditionalTableStyleFixture = `<w:style w:type="table" w:default="1" w:styleId="NormalTable"><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
	`<w:style w:type="table" w:styleId="Shaded"><w:basedOn w:val="NormalTable"/><w:pPr><w:spacing w:after="0"/></w:pPr>` +
	`<w:tblPr><w:tblStyleRowBandSize w:val="1"/><w:tblStyleColBandSize w:val="1"/><w:tblBorders><w:top w:val="single" w:sz="18" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>` +
	`<w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/><w:sz w:val="20"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="833C0B"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="lastRow"><w:rPr><w:sz w:val="48"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="B4C6E7"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="firstCol"><w:rPr><w:sz w:val="72"/></w:rPr><w:tblPr/><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:bottom w:val="single" w:sz="18" w:space="0" w:color="auto"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="ED7D31"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="lastCol"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="595959"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="band1Vert"><w:rPr><w:sz w:val="52"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="A6A6A6" w:themeFill="background1" w:themeFillShade="A6"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="band2Vert"><w:rPr><w:sz w:val="52"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="1F4E79"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="band1Horz"><w:rPr><w:sz w:val="20"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="BDD6EE"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="band2Horz"><w:rPr><w:sz w:val="20"/></w:rPr><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="C5E0B3"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="neCell"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FFC000"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="nwCell"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="7F7F7F"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="seCell"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="385623"/></w:tcPr></w:tblStylePr>` +
	`<w:tblStylePr w:type="swCell"><w:tblPr/><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="FF0000"/></w:tcPr></w:tblStylePr></w:style>`

func conditionalTableDocument(look string, rows, columns int) string {
	var b strings.Builder
	b.WriteString(`<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="Shaded"/><w:tblW w:w="0" w:type="auto"/>` + look + `</w:tblPr><w:tblGrid>`)
	for c := 0; c < columns; c++ {
		b.WriteString(`<w:gridCol w:w="1000"/>`)
	}
	b.WriteString(`</w:tblGrid>`)
	for r := 0; r < rows; r++ {
		b.WriteString(`<w:tr>`)
		for c := 0; c < columns; c++ {
			b.WriteString(`<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc>`)
		}
		b.WriteString(`</w:tr>`)
	}
	b.WriteString(`</w:tbl><w:sectPr/></w:body></w:document>`)
	return b.String()
}

func conditionalFills(resolved *NativeResolvedLayoutInputV1, doc *NativeDocumentV1) map[[2]int]string {
	byID := map[string]string{}
	for _, entry := range resolved.Tables[0].ConditionalCellShading {
		byID[entry.CellID] = entry.ShadingRGB
	}
	out := map[[2]int]string{}
	table := doc.Body.Blocks[0].Table
	for r, row := range table.Rows {
		for c, cell := range row.Cells {
			out[[2]int{r, c}] = byID[cell.ID]
		}
	}
	return out
}

func TestNativeConditionalTableStyleResolvesRegionsInWordPrecedence(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` + conditionalTableStyleFixture + `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"/></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = conditionalTableDocument(`<w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/>`, 4, 4)
	data := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	fills := conditionalFills(resolved, doc)
	// Read off Word's raster: corners over rows/columns, first/last row over
	// first/last column, columns over bands, vertical bands over horizontal
	// bands, and banding starting after the first row and column.
	want := map[[2]int]string{
		{0, 0}: "7F7F7F", {0, 1}: "833C0B", {0, 2}: "833C0B", {0, 3}: "FFC000",
		{1, 0}: "ED7D31", {1, 1}: "A6A6A6", {1, 2}: "1F4E79", {1, 3}: "595959",
		{2, 0}: "ED7D31", {2, 1}: "A6A6A6", {2, 2}: "1F4E79", {2, 3}: "595959",
		{3, 0}: "FF0000", {3, 1}: "B4C6E7", {3, 2}: "B4C6E7", {3, 3}: "385623",
	}
	for key, fill := range want {
		if fills[key] != fill {
			t.Fatalf("cell %v fill %q want %q (all %v)", key, fills[key], fill, fills)
		}
	}
	// The run cascade follows the same order: the first column's 36 pt covers
	// the vertical bands' 26 pt in the interior rows, the first row's 10 pt
	// covers both, the last row's 24 pt covers the bands, and the last column
	// (which states no size) still sits in a horizontally banded row at 10 pt.
	sizes := map[string]int{}
	for _, run := range resolved.Runs {
		if run.Properties.FontSizeHalfPoint != nil {
			sizes[run.ParagraphID] = *run.Properties.FontSizeHalfPoint
		}
	}
	table := doc.Body.Blocks[0].Table
	size := func(r, c int) int { return sizes[table.Rows[r].Cells[c].Paragraphs[0].ID] }
	for _, check := range []struct{ r, c, want int }{{0, 1, 20}, {0, 0, 20}, {1, 0, 72}, {1, 1, 52}, {1, 2, 52}, {1, 3, 20}, {3, 1, 48}, {3, 0, 48}} {
		if got := size(check.r, check.c); got != check.want {
			t.Fatalf("cell %d,%d font size %d want %d", check.r, check.c, got, check.want)
		}
	}
	color := ""
	for _, run := range resolved.Runs {
		if run.ParagraphID == table.Rows[0].Cells[1].Paragraphs[0].ID && run.Properties.Color != nil {
			color = *run.Properties.Color
		}
	}
	if color != "FFFFFF" {
		t.Fatalf("first-row run colour %q want FFFFFF", color)
	}
	// The region cascade is resolved, so the style node no longer carries the
	// blanket refusal. The region properties this tier does not model are
	// disclosed at their own elements instead: the first column's cell borders,
	// and the vertical band's theme shade, whose authored w:fill is applied
	// (Word writes the shaded colour into w:fill, as its built-in styles show).
	disclosed := map[string]bool{}
	for _, diagnostic := range resolved.Diagnostics {
		if diagnostic.Code != "CONDITIONAL_TABLE_STYLE_PRESERVED" {
			continue
		}
		if diagnostic.Path == nil {
			t.Fatalf("conditional regions still refused at %#v", diagnostic)
		}
		switch {
		case strings.HasSuffix(*diagnostic.Path, "/w:tblStylePr[5]/w:tcPr[1]/w:shd[1]"):
			disclosed["shade"] = true
		case strings.HasSuffix(*diagnostic.Path, "/w:tblStylePr[3]/w:tcPr[1]/w:tcBorders[1]"):
			disclosed["borders"] = true
		default:
			t.Fatalf("conditional regions still refused at %#v", diagnostic)
		}
	}
	if !disclosed["shade"] || !disclosed["borders"] {
		t.Fatalf("unmodeled region properties were not disclosed: %#v", resolved.Diagnostics)
	}
	// The auto-coloured table border is still outside the exact subset.
	if !hasResolutionDiagnostic(resolved, "TABLE_STYLE_EFFECTS_PRESERVED") || resolved.Tables[0].Borders != nil {
		t.Fatalf("auto border handling changed: %#v", resolved.Tables[0])
	}
}

func TestNativeConditionalTableStyleHonoursLookSwitches(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `">` + conditionalTableStyleFixture + `</w:styles>`
	for _, test := range []struct {
		name string
		look string
		want map[[2]int]string
	}{
		{
			// First row and first column only (Word's 04A0 minus noVBand): the
			// top-left cell is still a corner, row banding starts at row 1,
			// column banding at column 1, and the last row and column are
			// ordinary banded cells.
			name: "first row and column with both bands",
			look: `<w:tblLook w:val="00A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="0"/>`,
			want: map[[2]int]string{{0, 0}: "7F7F7F", {0, 1}: "833C0B", {0, 3}: "833C0B", {1, 0}: "ED7D31", {1, 1}: "A6A6A6", {1, 2}: "1F4E79", {1, 3}: "A6A6A6", {3, 0}: "ED7D31", {3, 1}: "A6A6A6", {3, 3}: "A6A6A6"},
		},
		{
			// Word's default 04A0: no vertical bands, so the horizontal bands show.
			name: "no vertical bands",
			look: `<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>`,
			want: map[[2]int]string{{0, 1}: "833C0B", {1, 0}: "ED7D31", {1, 1}: "BDD6EE", {2, 1}: "C5E0B3", {3, 1}: "BDD6EE", {3, 3}: "BDD6EE"},
		},
		{
			// Every switch off: only the bands apply, from row 0 and column 0.
			name: "bands only",
			look: `<w:tblLook w:val="0000"/>`,
			want: map[[2]int]string{{0, 0}: "A6A6A6", {0, 1}: "1F4E79", {1, 0}: "A6A6A6", {3, 3}: "1F4E79"},
		},
		{
			// Named switches alone select the same regions as the packed value.
			name: "named switches only",
			look: `<w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="0" w:lastColumn="0" w:noHBand="1" w:noVBand="1"/>`,
			want: map[[2]int]string{{0, 0}: "833C0B", {0, 3}: "833C0B", {1, 1}: "", {2, 3}: "", {3, 0}: "B4C6E7", {3, 3}: "B4C6E7"},
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = conditionalTableDocument(test.look, 4, 4)
			data := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			fills := conditionalFills(resolved, doc)
			for key, fill := range test.want {
				if fills[key] != fill {
					t.Fatalf("cell %v fill %q want %q (all %v)", key, fills[key], fill, fills)
				}
			}
		})
	}
}

func TestNativeConditionalTableStyleAppliesBandSizesAndSpans(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `">` + strings.Replace(strings.Replace(conditionalTableStyleFixture, `<w:tblStyleRowBandSize w:val="1"/>`, `<w:tblStyleRowBandSize w:val="2"/>`, 1), `<w:tblStyleColBandSize w:val="1"/>`, `<w:tblStyleColBandSize w:val="2"/>`, 1) + `</w:styles>`
	parts := resolvedStylesTestParts(styles)
	// Six rows, five grid columns; the last row's second cell spans the three
	// interior columns and the last column.
	document := conditionalTableDocument(`<w:tblLook w:val="01E0"/>`, 5, 5)
	document = strings.Replace(document, `</w:tr></w:tbl>`, `</w:tr><w:tr><w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/><w:gridSpan w:val="4"/></w:tcPr><w:p><w:r><w:t>x</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`, 1)
	parts["word/document.xml"] = document
	data := buildNativeDOCX(t, nativeEntries(parts))
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	fills := conditionalFills(resolved, doc)
	// Rows 1-2 are band 1, rows 3-4 band 2; columns 1-2 band 1, column 3 band 2
	// (column 4 is the last column). The spanning cell reaches the last grid
	// column, so it is the bottom-right corner.
	want := map[[2]int]string{{1, 1}: "A6A6A6", {1, 2}: "A6A6A6", {1, 3}: "1F4E79", {3, 1}: "A6A6A6", {4, 3}: "1F4E79", {1, 4}: "595959", {5, 0}: "FF0000", {5, 1}: "385623"}
	for key, fill := range want {
		if fills[key] != fill {
			t.Fatalf("cell %v fill %q want %q (all %v)", key, fills[key], fill, fills)
		}
	}
}

func TestNativeConditionalTableStyleKeepsRefusingUnreadableSelections(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `">` + conditionalTableStyleFixture + `</w:styles>`
	for name, look := range map[string]string{
		"absent":             ``,
		"empty":              `<w:tblLook/>`,
		"disagreeing forms":  `<w:tblLook w:val="0020" w:firstRow="0"/>`,
		"out-of-range bits":  `<w:tblLook w:val="F000"/>`,
		"duplicate elements": `<w:tblLook w:val="0020"/><w:tblLook w:val="0020"/>`,
	} {
		t.Run(name, func(t *testing.T) {
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = conditionalTableDocument(look, 2, 2)
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if len(resolved.Tables[0].ConditionalCellShading) != 0 || !hasResolutionDiagnostic(resolved, "CONDITIONAL_TABLE_STYLE_PRESERVED") {
				t.Fatalf("unreadable look was guessed: %#v %#v", resolved.Tables[0], resolved.Diagnostics)
			}
			for _, run := range resolved.Runs {
				if run.Properties.FontSizeHalfPoint != nil {
					t.Fatalf("region cascade applied without a readable look: %#v", run.Properties)
				}
			}
		})
	}
}
