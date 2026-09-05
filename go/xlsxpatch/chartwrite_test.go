package xlsxpatch

import (
	"strings"
	"testing"
)

func writeSpec(typ string) ChartWriteSpec {
	return ChartWriteSpec{
		SheetName: "Data",
		Type:      typ,
		Title:     "Revenue by Quarter",
		Series: []WriteSeries{
			{Name: "Revenue", NameRef: "Data!$B$1", CategoriesRef: "Data!$A$2:$A$5", ValuesRef: "Data!$B$2:$B$5"},
			{Name: "Costs", NameRef: "Data!$C$1", CategoriesRef: "Data!$A$2:$A$5", ValuesRef: "Data!$C$2:$C$5"},
		},
		Anchor: ChartAnchor{FromCol: 5, FromRow: 2, ToCol: 13, ToRow: 18},
	}
}

func TestResolveRelPathCanonicalizesInternalOPCTargets(t *testing.T) {
	tests := []struct {
		name, base, target, want string
	}{
		{name: "leading dot", base: "xl", target: "./worksheets/sheet1.xml", want: "xl/worksheets/sheet1.xml"},
		{name: "interior parent", base: "xl/worksheets", target: "tables/../sheet1.xml", want: "xl/worksheets/sheet1.xml"},
		{name: "encoded interior parent", base: "xl/worksheets", target: "tables/%2E%2E/sheet1.xml", want: "xl/worksheets/sheet1.xml"},
		{name: "root relative", base: "ignored/base", target: "/xl/worksheets/sheet1.xml", want: "xl/worksheets/sheet1.xml"},
		{name: "part fragment", base: "xl", target: "worksheets/sheet1.xml#anchor", want: "xl/worksheets/sheet1.xml"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			got, err := resolveRelPath(test.base, test.target)
			if err != nil {
				t.Fatal(err)
			}
			if got != test.want {
				t.Fatalf("resolveRelPath(%q, %q) = %q, want %q", test.base, test.target, got, test.want)
			}
		})
	}
}

func TestResolveRelPathRefusesTraversalAbovePackageRoot(t *testing.T) {
	for _, target := range []string{"../../worksheets/sheet1.xml", "../%2e%2e/worksheets/sheet1.xml", "/../xl/worksheets/sheet1.xml"} {
		t.Run(target, func(t *testing.T) {
			if got, err := resolveRelPath("xl", target); err == nil || !strings.Contains(err.Error(), "above the package root") {
				t.Fatalf("resolveRelPath returned %q, err=%v", got, err)
			}
		})
	}
}

func TestResolveRelPathRefusesNonCanonicalPathSegments(t *testing.T) {
	tests := []struct {
		target, wantError string
	}{
		{target: "worksheets//sheet1.xml", wantError: "empty path segment"},
		{target: "worksheets%2Fsheet1.xml", wantError: "encoded path separator"},
		{target: "worksheets%5Csheet1.xml", wantError: "encoded path separator"},
		{target: "worksheets/%00sheet1.xml", wantError: "control character"},
		{target: "worksheets./sheet1.xml", wantError: "ends with a dot"},
		{target: "worksheets%2E/sheet1.xml", wantError: "ends with a dot"},
		{target: "#anchor", wantError: "no resolvable part path"},
	}
	for _, test := range tests {
		t.Run(test.target, func(t *testing.T) {
			if got, err := resolveRelPath("xl", test.target); err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("resolveRelPath returned %q, err=%v; want %q", got, err, test.wantError)
			}
		})
	}
}

func TestAddChart_FreshWorkbook_RoundTripsThroughReader(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := AddChart(orig, writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}

	// Our own reader — the same code the browser bridge uses — must see
	// exactly what we wrote. This is the self-validating loop.
	charts, err := ReadCharts(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(charts))
	}
	c := charts[0]
	if c.Type != "column" || c.Title != "Revenue by Quarter" {
		t.Errorf("round-trip mismatch: type=%q title=%q", c.Type, c.Title)
	}
	if len(c.Series) != 2 || c.Series[0].Name != "Revenue" || c.Series[1].ValuesRef != "Data!$C$2:$C$5" {
		t.Errorf("series round-trip mismatch: %+v", c.Series)
	}
	if c.Series[0].CategoriesRef != "Data!$A$2:$A$5" || c.Series[0].NameRef != "Data!$B$1" {
		t.Errorf("refs mismatch: %+v", c.Series[0])
	}

	// Structural hookup: worksheet gained its drawing element, rels and
	// content types exist, untouched parts survived (Apply verified already,
	// but assert the worksheet content specifically since it was Replaced).
	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, `r:id="rId1"/>`) || !strings.Contains(sheet, "<drawing ") {
		t.Errorf("worksheet missing drawing element: %s", sheet)
	}
	if !strings.Contains(sheet, "Quarter") {
		t.Error("worksheet data lost")
	}
	rels := readEntry(t, out, "xl/worksheets/_rels/sheet1.xml.rels")
	if !strings.Contains(rels, "../drawings/drawing1.xml") {
		t.Errorf("sheet rels missing drawing target: %s", rels)
	}
	drawingRels := readEntry(t, out, "xl/drawings/_rels/drawing1.xml.rels")
	if !strings.Contains(drawingRels, "../charts/chart1.xml") {
		t.Errorf("drawing rels missing chart target: %s", drawingRels)
	}
	ct := readEntry(t, out, "[Content_Types].xml")
	for _, want := range []string{"/xl/charts/chart1.xml", "/xl/drawings/drawing1.xml"} {
		if !strings.Contains(ct, `PartName="`+want+`"`) {
			t.Errorf("content types missing %s: %s", want, ct)
		}
	}
	drawing := readEntry(t, out, "xl/drawings/drawing1.xml")
	for _, want := range []string{"<xdr:col>5</xdr:col>", "<xdr:row>2</xdr:row>", "<xdr:col>13</xdr:col>", "<xdr:row>18</xdr:row>"} {
		if !strings.Contains(drawing, want) {
			t.Errorf("anchor missing %s", want)
		}
	}
}

func TestAddChart_AllWritableTypes(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	wantType := map[string]string{
		"column": "column", "bar": "bar", "line": "line", "area": "area",
		"pie": "pie", "doughnut": "doughnut", "scatter": "scatter",
	}
	for typ, want := range wantType {
		out, err := AddChart(orig, writeSpec(typ))
		if err != nil {
			t.Fatalf("%s: %v", typ, err)
		}
		charts, err := ReadCharts(out)
		if err != nil || len(charts) != 1 {
			t.Fatalf("%s: read failed: %v (%d charts)", typ, err, len(charts))
		}
		if charts[0].Type != want {
			t.Errorf("%s: round-tripped as %q", typ, charts[0].Type)
		}
	}
}

func TestAddChart_SecondChartExtendsExistingDrawing(t *testing.T) {
	// Build a workbook whose sheet already hooks up a drawing (the shape an
	// agent-generated openpyxl chart produces) — the second chart must extend
	// that drawing part, not create a second <drawing> element.
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>`
	entries["xl/worksheets/_rels/sheet1.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="` + relTypeDrawing + `" Target="../drawings/drawing1.xml"/></Relationships>`
	entries["xl/drawings/drawing1.xml"] = drawingXMLDoc(`<xdr:twoCellAnchor><xdr:clientData/></xdr:twoCellAnchor>`)
	entries["xl/drawings/_rels/drawing1.xml.rels"] = relsDoc(relationshipXML("rId1", relTypeChart, "../charts/chart1.xml"))
	entries["xl/charts/chart1.xml"] = barChartXML
	orig := buildZip(t, entries)

	out, err := AddChart(orig, writeSpec("line"))
	if err != nil {
		t.Fatal(err)
	}
	charts, err := ReadCharts(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 2 {
		t.Fatalf("expected 2 charts, got %d", len(charts))
	}
	// New chart got the next free part number.
	if charts[1].Part != "xl/charts/chart2.xml" {
		t.Errorf("new chart part = %q", charts[1].Part)
	}
	// The sheet still has exactly ONE drawing element.
	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	if strings.Count(sheet, "<drawing ") != 1 {
		t.Errorf("worksheet must keep exactly one drawing element: %s", sheet)
	}
	// The existing drawing part gained a second anchor + rel rId2.
	drawing := readEntry(t, out, "xl/drawings/drawing1.xml")
	if strings.Count(drawing, "twoCellAnchor") < 4 { // 2 anchors × open+close
		t.Errorf("drawing not extended: %s", drawing)
	}
	drawingRels := readEntry(t, out, "xl/drawings/_rels/drawing1.xml.rels")
	if !strings.Contains(drawingRels, `Id="rId2"`) || !strings.Contains(drawingRels, "../charts/chart2.xml") {
		t.Errorf("drawing rels not extended: %s", drawingRels)
	}
	// Original chart untouched.
	if got := readEntry(t, out, "xl/charts/chart1.xml"); got != barChartXML {
		t.Error("existing chart part changed")
	}
}

func TestAddChart_DrawingElementPlacedBeforeTableParts(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData/><tableParts count="1"><tablePart r:id="rId9"/></tableParts></worksheet>`
	orig := buildZip(t, entries)
	out, err := AddChart(orig, writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}
	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	drawingIdx := strings.Index(sheet, "<drawing ")
	tableIdx := strings.Index(sheet, "<tableParts")
	if drawingIdx < 0 || tableIdx < 0 || drawingIdx > tableIdx {
		t.Errorf("drawing must precede tableParts: %s", sheet)
	}
}

func TestContentTypesWithSortsOverridesDeterministically(t *testing.T) {
	const contentTypes = `<Types><Default Extension="xml" ContentType="application/xml"/></Types>`
	overrides := map[string]string{
		"/xl/drawings/drawing1.xml": "drawing",
		"/xl/charts/chart1.xml":     "chart",
	}
	const want = `<Types><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="chart"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="drawing"/></Types>`
	for i := 0; i < 100; i++ {
		got, err := contentTypesWith(contentTypes, overrides)
		if err != nil {
			t.Fatal(err)
		}
		if got != want {
			t.Fatalf("content type override order is nondeterministic:\n got: %s\nwant: %s", got, want)
		}
	}
}

func TestAddChart_Rejections(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	if _, err := AddChart(orig, writeSpec("heatmap")); err == nil {
		t.Error("browser-only type must be rejected")
	}
	bad := writeSpec("column")
	bad.Series = nil
	if _, err := AddChart(orig, bad); err == nil {
		t.Error("empty series must be rejected")
	}
	deg := writeSpec("column")
	deg.Anchor = ChartAnchor{FromCol: 3, FromRow: 3, ToCol: 3, ToRow: 3}
	if _, err := AddChart(orig, deg); err == nil {
		t.Error("degenerate anchor must be rejected")
	}
	ghost := writeSpec("column")
	ghost.SheetName = "NoSuchSheet"
	if _, err := AddChart(orig, ghost); err == nil {
		t.Error("unknown sheet must be rejected")
	}
}

func TestAddChart_TitleEscaping(t *testing.T) {
	spec := writeSpec("column")
	spec.Title = `Q1 <Revenue> & "Costs"`
	out, err := AddChart(buildZip(t, fixtureWorkbook(false)), spec)
	if err != nil {
		t.Fatal(err)
	}
	charts, err := ReadCharts(out)
	if err != nil || len(charts) != 1 {
		t.Fatalf("read: %v", err)
	}
	if charts[0].Title != spec.Title {
		t.Errorf("title round-trip: %q", charts[0].Title)
	}
}
