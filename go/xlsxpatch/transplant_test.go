package xlsxpatch

import (
	"strings"
	"testing"
)

// chartedWorkbook builds a workbook whose chart came from our own writer —
// the realistic transplant source (agent-generated files pass through the
// same DrawingML shapes).
func chartedWorkbook(t *testing.T) []byte {
	t.Helper()
	out, err := AddChart(buildZip(t, fixtureWorkbook(false)), ChartWriteSpec{
		SheetName: "Data",
		Type:      "column",
		Title:     "Revenue by Quarter",
		Series: []WriteSeries{
			{Name: "Revenue", NameRef: "Data!$B$1", CategoriesRef: "Data!$A$2:$A$5", ValuesRef: "Data!$B$2:$B$5"},
		},
		Anchor: ChartAnchor{FromCol: 6, FromRow: 3, ToCol: 14, ToRow: 20},
	})
	if err != nil {
		t.Fatal(err)
	}
	return out
}

func TestReadChartAnchors(t *testing.T) {
	anchors, err := ReadChartAnchors(chartedWorkbook(t))
	if err != nil {
		t.Fatal(err)
	}
	a, ok := anchors["xl/charts/chart1.xml"]
	if !ok {
		t.Fatalf("no anchor for chart1: %v", anchors)
	}
	if a != (ChartAnchor{FromCol: 6, FromRow: 3, ToCol: 14, ToRow: 20}) {
		t.Errorf("anchor = %+v", a)
	}
}

func TestTransplantCharts_PreservesChartIntoRebuiltFile(t *testing.T) {
	src := chartedWorkbook(t)
	// The "rebuilt by the editor" file: same sheet, edited data, no charts.
	rebuilt := fixtureWorkbook(false)
	rebuilt["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>EDITED</v></c></row></sheetData></worksheet>`
	dst := buildZip(t, rebuilt)

	merged, err := TransplantCharts(src, dst)
	if err != nil {
		t.Fatal(err)
	}

	// The edit survives...
	if got := readEntry(t, merged, "xl/worksheets/sheet1.xml"); !strings.Contains(got, "EDITED") {
		t.Error("rebuilt sheet content lost")
	}
	// ...and so does the chart, with its identity and position.
	charts, err := ReadCharts(merged)
	if err != nil || len(charts) != 1 {
		t.Fatalf("charts after transplant: %v (%d)", err, len(charts))
	}
	if charts[0].Type != "column" || charts[0].Title != "Revenue by Quarter" {
		t.Errorf("chart identity lost: %+v", charts[0])
	}
	if charts[0].Series[0].ValuesRef != "Data!$B$2:$B$5" {
		t.Errorf("series binding lost: %+v", charts[0].Series)
	}
	anchors, _ := ReadChartAnchors(merged)
	if anchors["xl/charts/chart1.xml"] != (ChartAnchor{FromCol: 6, FromRow: 3, ToCol: 14, ToRow: 20}) {
		t.Errorf("anchor lost: %+v", anchors)
	}
}

func TestTransplantCharts_NoChartsIsIdentity(t *testing.T) {
	dst := buildZip(t, fixtureWorkbook(false))
	merged, err := TransplantCharts(buildZip(t, fixtureWorkbook(false)), dst)
	if err != nil {
		t.Fatal(err)
	}
	if len(merged) != len(dst) {
		t.Error("chartless transplant must return dst unchanged")
	}
}

func TestTransplantCharts_FailsClosed(t *testing.T) {
	dst := buildZip(t, fixtureWorkbook(false))

	// Unwritable chart type in the source (bubbleChart → "unknown").
	entries := fixtureWorkbook(false)
	entries["xl/charts/chart1.xml"] = `<c:chartSpace xmlns:c="x"><c:chart><c:plotArea><c:bubbleChart/></c:plotArea></c:chart></c:chartSpace>`
	if _, err := TransplantCharts(buildZip(t, entries), dst); err == nil {
		t.Error("unwritable chart type must fail the transplant")
	}

	// Series bound to a sheet the rebuilt file doesn't have.
	src := chartedWorkbook(t)
	ghost := fixtureWorkbook(false)
	ghost["xl/workbook.xml"] = `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Other" sheetId="1" r:id="rId1"/></sheets></workbook>`
	if _, err := TransplantCharts(src, buildZip(t, ghost)); err == nil {
		t.Error("missing target sheet must fail the transplant")
	}
}

func TestSheetNameFromRef(t *testing.T) {
	cases := map[string]string{
		"Data!$B$2:$B$5":    "Data",
		"'My Sheet'!$A$1":   "My Sheet",
		"'It''s 2026'!$A$1": "It's 2026",
	}
	for ref, want := range cases {
		got, err := sheetNameFromRef(ref)
		if err != nil || got != want {
			t.Errorf("%q → %q (%v), want %q", ref, got, err, want)
		}
	}
	if _, err := sheetNameFromRef("$A$1"); err == nil {
		t.Error("unqualified ref must error")
	}
}
