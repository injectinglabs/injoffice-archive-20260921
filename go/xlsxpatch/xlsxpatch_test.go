package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
)

// buildZip assembles an in-memory zip from name→content, in map-independent
// deterministic order.
func buildZip(t *testing.T, entries map[string]string) []byte {
	t.Helper()
	var names []string
	for n := range entries {
		names = append(names, n)
	}
	// stable order
	for i := 0; i < len(names); i++ {
		for j := i + 1; j < len(names); j++ {
			if names[j] < names[i] {
				names[i], names[j] = names[j], names[i]
			}
		}
	}
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for _, n := range names {
		w, err := zw.Create(n)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(entries[n])); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func readEntry(t *testing.T, data []byte, name string) string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name == name {
			rc, err := f.Open()
			if err != nil {
				t.Fatal(err)
			}
			defer rc.Close()
			b, err := io.ReadAll(rc)
			if err != nil {
				t.Fatal(err)
			}
			return string(b)
		}
	}
	t.Fatalf("entry %q not found", name)
	return ""
}

const barChartXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title><c:tx><c:rich><a:p><a:r><a:t>Quarterly Revenue</a:t></a:r></a:p></c:rich></c:tx></c:title>
    <c:plotArea>
      <c:layout/>
      <c:barChart>
        <c:barDir val="col"/>
        <c:grouping val="clustered"/>
        <c:ser>
          <c:idx val="0"/>
          <c:tx><c:strRef><c:f>Data!$B$1</c:f><c:strCache><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Data!$B$2:$B$5</c:f></c:numRef></c:val>
        </c:ser>
        <c:ser>
          <c:idx val="1"/>
          <c:tx><c:strRef><c:f>Data!$C$1</c:f><c:strCache><c:pt idx="0"><c:v>Costs</c:v></c:pt></c:strCache></c:strRef></c:tx>
          <c:cat><c:strRef><c:f>Data!$A$2:$A$5</c:f></c:strRef></c:cat>
          <c:val><c:numRef><c:f>Data!$C$2:$C$5</c:f></c:numRef></c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>
</c:chartSpace>`

func fixtureWorkbook(withChart bool) map[string]string {
	m := map[string]string{
		"[Content_Types].xml":        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/></Types>`,
		"_rels/.rels":                `<Relationships/>`,
		"xl/workbook.xml":            `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
		"xl/worksheets/sheet1.xml":   `<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>Quarter</v></c></row></sheetData></worksheet>`,
		"xl/styles.xml":              `<styleSheet/>`,
		"docProps/core.xml":          `<coreProperties/>`,
	}
	if withChart {
		m["xl/charts/chart1.xml"] = barChartXML
		m["xl/drawings/drawing1.xml"] = `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>`
	}
	return m
}

func TestApply_ReplacesOnlyTargetedEntry(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(true))
	newSheet := `<worksheet><sheetData><row r="1"><c r="A1" t="str"><v>EDITED</v></c></row></sheetData></worksheet>`

	out, err := Apply(orig, Patch{Replace: map[string][]byte{"xl/worksheets/sheet1.xml": []byte(newSheet)}})
	if err != nil {
		t.Fatal(err)
	}
	if got := readEntry(t, out, "xl/worksheets/sheet1.xml"); got != newSheet {
		t.Errorf("sheet not replaced: %q", got)
	}
	// The chart — a part the "editor" doesn't model — survives byte-for-byte.
	if got := readEntry(t, out, "xl/charts/chart1.xml"); got != barChartXML {
		t.Error("untouched chart part changed")
	}
}

func TestApply_AddAndDelete(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := Apply(orig, Patch{
		Add:    map[string][]byte{"xl/charts/chart1.xml": []byte(barChartXML)},
		Delete: map[string]bool{"docProps/core.xml": true},
	})
	if err != nil {
		t.Fatal(err)
	}
	if got := readEntry(t, out, "xl/charts/chart1.xml"); got != barChartXML {
		t.Error("added entry wrong")
	}
	zr, _ := zip.NewReader(bytes.NewReader(out), int64(len(out)))
	for _, f := range zr.File {
		if f.Name == "docProps/core.xml" {
			t.Error("deleted entry still present")
		}
	}
}

func TestApply_FailsClosed(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	cases := []Patch{
		{Replace: map[string][]byte{"xl/nope.xml": []byte("x")}}, // replace missing
		{Add: map[string][]byte{"xl/workbook.xml": []byte("x")}}, // add existing
		{Delete: map[string]bool{"xl/nope.xml": true}},           // delete missing
		{Replace: map[string][]byte{"xl/workbook.xml": []byte("x")}, Delete: map[string]bool{"xl/workbook.xml": true}},
	}
	for i, p := range cases {
		if _, err := Apply(orig, p); err == nil {
			t.Errorf("case %d: expected error, got success", i)
		}
	}
	if _, err := Apply([]byte("not a zip"), Patch{}); err == nil {
		t.Error("corrupt input must error")
	}
}

func TestApply_EmptyPatchIsIdentityOnContent(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(true))
	out, err := Apply(orig, Patch{})
	if err != nil {
		t.Fatal(err)
	}
	for name, content := range fixtureWorkbook(true) {
		if got := readEntry(t, out, name); got != content {
			t.Errorf("entry %q changed under empty patch", name)
		}
	}
}

func TestInspect(t *testing.T) {
	withChart, _ := Inspect(buildZip(t, fixtureWorkbook(true)))
	if !withChart.HasCharts() || len(withChart.ChartParts) != 1 {
		t.Errorf("expected 1 chart part, got %v", withChart.ChartParts)
	}
	if len(withChart.DrawingParts) != 1 {
		t.Errorf("expected 1 drawing part, got %v", withChart.DrawingParts)
	}
	plain, _ := Inspect(buildZip(t, fixtureWorkbook(false)))
	if plain.HasCharts() {
		t.Error("chartless workbook reported charts")
	}
	if _, err := Inspect([]byte("junk")); err == nil {
		t.Error("corrupt input must error")
	}
}

func TestReadCharts(t *testing.T) {
	charts, err := ReadCharts(buildZip(t, fixtureWorkbook(true)))
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 1 {
		t.Fatalf("expected 1 chart, got %d", len(charts))
	}
	c := charts[0]
	if c.Type != "column" {
		t.Errorf("type = %q, want column", c.Type)
	}
	if c.Title != "Quarterly Revenue" {
		t.Errorf("title = %q", c.Title)
	}
	if len(c.Series) != 2 {
		t.Fatalf("series = %d, want 2", len(c.Series))
	}
	if c.Series[0].Name != "Revenue" || c.Series[1].Name != "Costs" {
		t.Errorf("series names = %q, %q", c.Series[0].Name, c.Series[1].Name)
	}
	if c.Series[0].ValuesRef != "Data!$B$2:$B$5" {
		t.Errorf("values ref = %q", c.Series[0].ValuesRef)
	}
	if c.Series[0].CategoriesRef != "Data!$A$2:$A$5" {
		t.Errorf("categories ref = %q", c.Series[0].CategoriesRef)
	}
	if c.Series[0].NameRef != "Data!$B$1" {
		t.Errorf("name ref = %q", c.Series[0].NameRef)
	}
}

func TestReadCharts_BarDirAndUnknown(t *testing.T) {
	horizontal := strings.Replace(barChartXML, `<c:barDir val="col"/>`, `<c:barDir val="bar"/>`, 1)
	entries := fixtureWorkbook(false)
	entries["xl/charts/chart1.xml"] = horizontal
	entries["xl/charts/chart2.xml"] = `<c:chartSpace xmlns:c="x"><c:chart><c:plotArea><c:bubbleChart/></c:plotArea></c:chart></c:chartSpace>`
	charts, err := ReadCharts(buildZip(t, entries))
	if err != nil {
		t.Fatal(err)
	}
	if len(charts) != 2 {
		t.Fatalf("expected 2 charts, got %d", len(charts))
	}
	if charts[0].Type != "bar" {
		t.Errorf("chart1 type = %q, want bar", charts[0].Type)
	}
	// bubbleChart is unmapped: reported, not silently dropped.
	if charts[1].Type != "unknown" {
		t.Errorf("chart2 type = %q, want unknown", charts[1].Type)
	}
}
