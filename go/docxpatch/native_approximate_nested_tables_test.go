package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

// Mirrors the benchmark markup: an auto-width outer table whose single cell
// holds paragraphs around a two-column inner table styled with tblBorders and a
// conditional firstRow layer.
const testNestedTableStyles = `<w:styles xmlns:w="` + testW + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults>` +
	`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:before="60" w:after="60"/></w:pPr></w:style>` +
	`<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
	`<w:style w:type="table" w:customStyle="1" w:styleId="Outer"><w:name w:val="Outer"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblInd w:w="360" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>` +
	`<w:style w:type="table" w:customStyle="1" w:styleId="Inner"><w:name w:val="Inner"/><w:basedOn w:val="TableNormal"/><w:tblPr><w:tblInd w:w="360" w:type="dxa"/><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="808080"/><w:left w:val="single" w:sz="12" w:space="0" w:color="808080"/><w:bottom w:val="single" w:sz="12" w:space="0" w:color="808080"/><w:right w:val="single" w:sz="12" w:space="0" w:color="808080"/><w:insideH w:val="single" w:sz="6" w:space="0" w:color="808080"/><w:insideV w:val="single" w:sz="6" w:space="0" w:color="808080"/></w:tblBorders><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="86" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="86" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="D9D9D9"/></w:tcPr></w:tblStylePr></w:style>` +
	`</w:styles>`

func nestedTableCell(width string, body string) string {
	return `<w:tc><w:tcPr><w:tcW w:w="` + width + `" w:type="dxa"/></w:tcPr>` + body + `</w:tc>`
}

func nestedInnerTable(rows string) string {
	return `<w:tbl><w:tblPr><w:tblStyle w:val="Inner"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="01E0" w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/></w:tblPr><w:tblGrid><w:gridCol w:w="3841"/><w:gridCol w:w="4049"/></w:tblGrid>` + rows + `</w:tbl>`
}

func nestedInnerRows() string {
	return `<w:tr><w:trPr><w:cnfStyle w:val="100000000000"/></w:trPr>` + nestedTableCell("3841", `<w:p><w:r><w:t>Setting</w:t></w:r></w:p>`) + nestedTableCell("4049", `<w:p><w:r><w:t>Instructions</w:t></w:r></w:p>`) + `</w:tr>` +
		`<w:tr>` + nestedTableCell("3841", `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Memory</w:t></w:r></w:p>`) + nestedTableCell("4049", `<w:p><w:r><w:t xml:space="preserve">We recommend that you use </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>Dynamic</w:t></w:r><w:r><w:t xml:space="preserve"> memory.</w:t></w:r></w:p>`) + `</w:tr>`
}

func nestedOuterTable(cellBody string) string {
	return `<w:tbl><w:tblPr><w:tblStyle w:val="Outer"/><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="8280"/></w:tblGrid><w:tr>` + nestedTableCell("8280", cellBody) + `</w:tr></w:tbl>`
}

func nativeApproximateNestedSource(t *testing.T, body string) []byte {
	entries := resolvedStylesTestParts(testNestedTableStyles)
	entries["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` + body + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
	return buildNativeDOCX(t, nativeEntries(entries))
}

func inspectNestedTables(t *testing.T, source []byte) (*NativeDocumentV1, *NativeApproximateNestedTablesV1) {
	t.Helper()
	before := append([]byte(nil), source...)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	strict, _ := json.Marshal(doc)
	out, err := InspectNativeApproximateNestedTablesV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if string(before) != string(source) {
		t.Fatal("inspection changed caller bytes")
	}
	again, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if after, _ := json.Marshal(again); string(after) != string(strict) {
		t.Fatal("inspection changed strict extraction")
	}
	return doc, out
}

func TestApproximateNestedTablesBenchmarkShape(t *testing.T) {
	source := nativeApproximateNestedSource(t, nestedOuterTable(`<w:p><w:r><w:t>1.</w:t></w:r></w:p><w:p/>`+nestedInnerTable(nestedInnerRows())+`<w:p/><w:p><w:r><w:t>2.</w:t></w:r></w:p>`))
	doc, out := inspectNestedTables(t, source)
	if out == nil || out.Protocol != NativeApproximateNestedTablesProtocol || out.Policy != NativeApproximateNestedTablePolicy || out.PackageSHA256 != doc.Source.PackageSHA256 || len(out.Items) != 1 || out.OmittedCount != 0 {
		t.Fatalf("unexpected sidecar: %#v", out)
	}
	item := out.Items[0]
	outer := doc.Body.Blocks[0].Table
	cell := outer.Rows[0].Cells[0]
	if item.Status != "supported" || item.Reason != "" || item.TableID != outer.ID || item.CellID != cell.ID || item.PrecedingParagraphs != 2 || item.ID != "approximate-nested-table:"+strings.TrimPrefix(cell.ID, "cell:")+":1" {
		t.Fatalf("unexpected item: %#v", item)
	}
	if len(cell.Paragraphs) != 4 {
		t.Fatalf("outer cell paragraphs = %d", len(cell.Paragraphs))
	}
	// The strict refusal is unchanged and joined at the identical anchor.
	joined := false
	for _, d := range doc.Unsupported {
		if d.Code == "NESTED_TABLE_OR_CELL_MARKUP" && d.ScopeID == outer.ID && d.Anchor != nil && d.Anchor.Path == item.Anchor.Path {
			joined = len(item.DiagnosticIDs) == 1 && d.ID == item.DiagnosticIDs[0] && d.Anchor.XMLSHA256 == item.Anchor.XMLSHA256 && *d.Anchor.StartByte == *item.Anchor.StartByte && *d.Anchor.EndByte == *item.Anchor.EndByte
		}
	}
	if !joined || item.Anchor.Path != "/w:document[1]/w:body[1]/w:tbl[1]/w:tr[1]/w:tc[1]/w:tbl[1]" {
		t.Fatalf("nested refusal not joined: %#v", item)
	}
	table := item.Table
	if table == nil || table.ID != item.ID || table.EditPolicy.Mode != "read-only" || len(table.GridWidthsTwips) != 2 || table.GridWidthsTwips[0] != 3841 || table.GridWidthsTwips[1] != 4049 || len(table.Rows) != 2 || table.TableStyleID == nil || *table.TableStyleID != "Inner" {
		t.Fatalf("unexpected inner table: %#v", table)
	}
	if item.Geometry == nil || item.Geometry.Layout != "autofit" || item.Geometry.IndentTwips != 360 || item.Geometry.WidthType != "auto" || item.Geometry.CellMargins != (NativeTableCellMarginsV1{LeftTwips: 86, RightTwips: 86}) {
		t.Fatalf("unexpected geometry: %#v", item.Geometry)
	}
	if item.StyleBorders == nil || item.StyleBorders.Top == nil || item.StyleBorders.Top.SizeEighthPoints != 12 || *item.StyleBorders.Top.ColorRGB != "808080" || item.StyleBorders.InsideHorizontal == nil || item.StyleBorders.InsideHorizontal.SizeEighthPoints != 6 || item.StyleCellShadingRGB != nil {
		t.Fatalf("unexpected style borders: %#v", item.StyleBorders)
	}
	ids := map[string]bool{}
	paragraphs, runs := 0, 0
	for r, row := range table.Rows {
		if row.ID != item.ID+":r"+string(rune('0'+r)) || len(row.Cells) != 2 {
			t.Fatalf("unexpected row: %#v", row)
		}
		for _, c := range row.Cells {
			if !strings.HasPrefix(c.ID, row.ID+"c") || c.GridSpan == nil || *c.GridSpan != 1 || c.VerticalMerge != "none" || len(c.Paragraphs) != 1 {
				t.Fatalf("unexpected cell: %#v", c)
			}
			for _, p := range c.Paragraphs {
				if !strings.HasPrefix(p.ID, c.ID+":p") || ids[p.ID] || p.EditPolicy.Mode != "read-only" {
					t.Fatalf("unexpected paragraph id or policy: %#v", p)
				}
				ids[p.ID] = true
				paragraphs++
				for _, run := range p.Runs {
					if !strings.HasPrefix(run.ID, p.ID+":r") || ids[run.ID] || run.Kind != "text" {
						t.Fatalf("unexpected run: %#v", run)
					}
					ids[run.ID] = true
					runs++
				}
			}
		}
	}
	if paragraphs != 4 || runs != 6 || len(item.ResolvedParagraphs) != 4 || len(item.ResolvedRuns) != 6 || item.OmittedRuns != 0 {
		t.Fatalf("resolved content mismatch: paragraphs=%d runs=%d resolved=%d/%d omitted=%d", paragraphs, runs, len(item.ResolvedParagraphs), len(item.ResolvedRuns), item.OmittedRuns)
	}
	for _, resolved := range item.ResolvedRuns {
		if !ids[resolved.RunID] || !ids[resolved.ParagraphID] || resolved.Properties.FontFamily == nil || *resolved.Properties.FontFamily != "Arial" {
			t.Fatalf("resolved run does not join inner ids or fonts: %#v", resolved)
		}
	}
	bold := 0
	for _, resolved := range item.ResolvedRuns {
		if resolved.Properties.Bold != nil && *resolved.Properties.Bold {
			bold++
		}
	}
	// Two direct bold runs plus the two header runs the firstRow region makes bold.
	if bold != 4 {
		t.Fatalf("direct and firstRow run formatting lost: %d bold runs", bold)
	}
	if item.FirstRowCellShadingRGB == nil || *item.FirstRowCellShadingRGB != "D9D9D9" {
		t.Fatalf("firstRow fill not described: %#v", item.FirstRowCellShadingRGB)
	}
	if item.OuterGeometry == nil || item.OuterGeometry.IndentTwips != 360 || item.OuterGeometry.CellMargins != (NativeTableCellMarginsV1{}) || item.OuterGeometry.WidthType != "auto" {
		t.Fatalf("outer geometry not described: %#v", item.OuterGeometry)
	}
	notes := strings.Join(item.Notes, "|")
	if !strings.Contains(notes, "firstRow conditional table-style region applied") || !strings.Contains(notes, "w:cnfStyle") {
		t.Fatalf("approximations not noted: %#v", item.Notes)
	}
	// A look that switches the first row off keeps the region unapplied.
	plainLook := strings.Replace(nestedInnerTable(nestedInnerRows()), `w:firstRow="1"`, `w:firstRow="0"`, 1)
	_, off := inspectNestedTables(t, nativeApproximateNestedSource(t, nestedOuterTable(`<w:p><w:r><w:t>1.</w:t></w:r></w:p>`+plainLook+`<w:p/>`)))
	if off == nil || len(off.Items) != 1 || off.Items[0].FirstRowCellShadingRGB != nil || !strings.Contains(strings.Join(off.Items[0].Notes, "|"), "other than firstRow not applied") {
		t.Fatalf("disabled firstRow look still applied: %#v", off.Items[0].Notes)
	}
	offBold := 0
	for _, resolved := range off.Items[0].ResolvedRuns {
		if resolved.Properties.Bold != nil && *resolved.Properties.Bold {
			offBold++
		}
	}
	if offBold != 2 {
		t.Fatalf("disabled firstRow look changed run formatting: %d bold runs", offBold)
	}
	// Body ids never collide with sidecar ids.
	for _, block := range doc.Body.Blocks {
		if block.Table != nil {
			for _, row := range block.Table.Rows {
				for _, c := range row.Cells {
					for _, p := range c.Paragraphs {
						if ids[p.ID] {
							t.Fatalf("sidecar reused body paragraph id %q", p.ID)
						}
					}
				}
			}
		}
	}
}

func TestApproximateNestedTablesOmitDepthTwoAndMerges(t *testing.T) {
	deeper := nestedInnerTable(`<w:tr>` + nestedTableCell("3841", `<w:p><w:r><w:t>a</w:t></w:r></w:p>`+nestedInnerTable(nestedInnerRows())+`<w:p/>`) + nestedTableCell("4049", `<w:p><w:r><w:t>b</w:t></w:r></w:p>`) + `</w:tr>`)
	source := nativeApproximateNestedSource(t, nestedOuterTable(`<w:p><w:r><w:t>1.</w:t></w:r></w:p>`+deeper+`<w:p/>`))
	doc, out := inspectNestedTables(t, source)
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "nested-depth-limit" || out.Items[0].Table != nil || out.Items[0].PrecedingParagraphs != 1 || len(out.Items[0].DiagnosticIDs) != 1 {
		t.Fatalf("depth two was not omitted: %#v", out)
	}
	if !hasUnsupportedCode(doc, "NESTED_TABLE_OR_CELL_MARKUP") {
		t.Fatal("strict nested refusal disappeared")
	}
	merged := nestedInnerTable(`<w:tr><w:tc><w:tcPr><w:tcW w:w="7890" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>span</w:t></w:r></w:p></w:tc></w:tr>` + `<w:tr>` + nestedTableCell("3841", `<w:p><w:r><w:t>a</w:t></w:r></w:p>`) + nestedTableCell("4049", `<w:p><w:r><w:t>b</w:t></w:r></w:p>`) + `</w:tr>`)
	source = nativeApproximateNestedSource(t, nestedOuterTable(`<w:p><w:r><w:t>1.</w:t></w:r></w:p>`+merged+`<w:p/>`))
	_, out = inspectNestedTables(t, source)
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "grid-mismatch" {
		t.Fatalf("horizontal merge was not omitted: %#v", out)
	}
	vmerge := nestedInnerTable(`<w:tr>` + `<w:tc><w:tcPr><w:tcW w:w="3841" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>` + nestedTableCell("4049", `<w:p><w:r><w:t>b</w:t></w:r></w:p>`) + `</w:tr>` + `<w:tr>` + `<w:tc><w:tcPr><w:tcW w:w="3841" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>` + nestedTableCell("4049", `<w:p><w:r><w:t>c</w:t></w:r></w:p>`) + `</w:tr>`)
	source = nativeApproximateNestedSource(t, nestedOuterTable(`<w:p><w:r><w:t>1.</w:t></w:r></w:p>`+vmerge+`<w:p/>`))
	_, out = inspectNestedTables(t, source)
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "merged-cells" {
		t.Fatalf("vertical merge was not omitted: %#v", out)
	}
}

func TestApproximateNestedTablesCountRefusedRunsAndSkipPlainDocuments(t *testing.T) {
	inner := nestedInnerTable(`<w:tr>` + nestedTableCell("3841", `<w:p><w:r><w:t>a</w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>`) + nestedTableCell("4049", `<w:p><w:r><w:t>b</w:t></w:r></w:p>`) + `</w:tr>`)
	source := nativeApproximateNestedSource(t, nestedOuterTable(inner+`<w:p><w:r><w:t>1.</w:t></w:r></w:p>`))
	_, out := inspectNestedTables(t, source)
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "supported" || out.Items[0].OmittedRuns == 0 || out.Items[0].PrecedingParagraphs != 0 {
		t.Fatalf("refused runs were not counted: %#v", out)
	}
	if runs := out.Items[0].Table.Rows[0].Cells[0].Paragraphs[0].Runs; len(runs) != 1 || runs[0].Kind != "text" {
		t.Fatalf("refused run was merged into the preview: %#v", runs)
	}
	plain := nativeApproximateNestedSource(t, nestedOuterTable(`<w:p><w:r><w:t>only text</w:t></w:r></w:p>`)+`<w:p><w:r><w:t>body</w:t></w:r></w:p>`)
	_, out = inspectNestedTables(t, plain)
	if out != nil {
		t.Fatalf("plain document produced a sidecar: %#v", out)
	}
	if _, err := InspectNativeApproximateNestedTablesV1(nil); err == nil {
		t.Fatal("empty package was accepted")
	}
}
