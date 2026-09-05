package docxpatch

import (
	"archive/zip"
	"bytes"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func buildChartTestDocx(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
			`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
			`</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
			`</Relationships>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
			`<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>` +
			`<w:sectPr/>` +
			`</w:body></w:document>`,
		"word/styles.xml": `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
	}
	for name, content := range parts {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(content)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func sampleChartSpec() ChartWriteSpec {
	return ChartWriteSpec{
		Type:  "column",
		Title: "Quarterly Revenue",
		Series: []ChartSeries{
			{Name: "2026", Categories: []string{"Q1", "Q2", "Q3"}, Values: []float64{10.5, 12, 15.25}},
		},
	}
}

func TestInsertChart_NoExistingRels(t *testing.T) {
	src := buildChartTestDocx(t)
	out, err := InsertChart(src, 0, sampleChartSpec(), PixelsToEMU(500), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}

	chartXML, ok := zipPart(t, out, "word/charts/chart1.xml")
	if !ok {
		t.Fatal("word/charts/chart1.xml missing")
	}
	if !strings.Contains(chartXML, "<c:barChart>") || !strings.Contains(chartXML, `barDir val="col"`) {
		t.Fatalf("expected a column (bar/col) chart: %s", chartXML)
	}
	if !strings.Contains(chartXML, "Quarterly Revenue") {
		t.Fatalf("title missing: %s", chartXML)
	}
	if !strings.Contains(chartXML, "<c:v>10.5</c:v>") || !strings.Contains(chartXML, "<c:v>Q1</c:v>") {
		t.Fatalf("cached values/categories missing: %s", chartXML)
	}

	rels, _ := zipPart(t, out, docRelsPart)
	if !strings.Contains(rels, `Type="`+relTypeChart+`"`) || !strings.Contains(rels, `Target="charts/chart1.xml"`) {
		t.Fatalf("relationship missing/wrong: %s", rels)
	}

	ct, _ := zipPart(t, out, contentTypes)
	if !strings.Contains(ct, `PartName="/word/charts/chart1.xml"`) {
		t.Fatalf("content types missing chart override: %s", ct)
	}

	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, "<c:chart ") {
		t.Fatalf("document.xml missing chart reference: %s", doc)
	}
	firstIdx := strings.Index(doc, "First paragraph.")
	chartIdx := strings.Index(doc, "<w:drawing>")
	secondIdx := strings.Index(doc, "Second paragraph.")
	if !(firstIdx < chartIdx && chartIdx < secondIdx) {
		t.Fatalf("chart paragraph not positioned between paragraph 0 and 1: %s", doc)
	}

	origStyles, _ := zipPart(t, src, "word/styles.xml")
	outStyles, _ := zipPart(t, out, "word/styles.xml")
	if origStyles != outStyles {
		t.Fatal("word/styles.xml must be byte-identical")
	}
}

func TestInsertChart_SecondChartGetsOwnPartAndRelationship(t *testing.T) {
	src := buildChartTestDocx(t)
	mid, err := InsertChart(src, 0, sampleChartSpec(), PixelsToEMU(400), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}
	pieSpec := ChartWriteSpec{Type: "pie", Series: []ChartSeries{{Name: "Share", Categories: []string{"A", "B"}, Values: []float64{60, 40}}}}
	out, err := InsertChart(mid, 1, pieSpec, PixelsToEMU(300), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := zipPart(t, out, "word/charts/chart1.xml"); !ok {
		t.Fatal("chart1.xml missing")
	}
	chart2, ok := zipPart(t, out, "word/charts/chart2.xml")
	if !ok || !strings.Contains(chart2, "<c:pieChart>") {
		t.Fatalf("chart2.xml missing/wrong: ok=%v %s", ok, chart2)
	}
	rels, _ := zipPart(t, out, docRelsPart)
	if strings.Count(rels, "relationships/chart") != 2 {
		t.Fatalf("expected two distinct chart relationships: %s", rels)
	}
	// docPr ids must not collide between the two charts.
	doc, _ := zipPart(t, out, docPart)
	if strings.Count(doc, `wp:docPr id="1"`) != 1 || strings.Count(doc, `wp:docPr id="2"`) != 1 {
		t.Fatalf("docPr ids should be 1 and 2, exactly once each: %s", doc)
	}
}

func TestInsertChart_DocPrIDDoesNotCollideWithAnImage(t *testing.T) {
	src := buildChartTestDocx(t)
	withImage, err := InsertImageAfter(src, 0, onePxPNG, "png", PixelsToEMU(100), PixelsToEMU(100))
	if err != nil {
		t.Fatal(err)
	}
	out, err := InsertChart(withImage, 1, sampleChartSpec(), PixelsToEMU(400), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}
	doc, _ := zipPart(t, out, docPart)
	if strings.Count(doc, `wp:docPr id="1"`) != 1 || strings.Count(doc, `wp:docPr id="2"`) != 1 {
		t.Fatalf("image and chart must not share a docPr id: %s", doc)
	}
}

func TestInsertChart_RejectsBadInput(t *testing.T) {
	src := buildChartTestDocx(t)
	cases := []struct {
		name string
		fn   func() error
	}{
		{"unsupported chart type", func() error {
			_, err := InsertChart(src, 0, ChartWriteSpec{Type: "sankey", Series: []ChartSeries{{Values: []float64{1}}}}, 100, 100)
			return err
		}},
		{"no series", func() error {
			_, err := InsertChart(src, 0, ChartWriteSpec{Type: "column"}, 100, 100)
			return err
		}},
		{"series with no values", func() error {
			_, err := InsertChart(src, 0, ChartWriteSpec{Type: "column", Series: []ChartSeries{{Name: "empty"}}}, 100, 100)
			return err
		}},
		{"zero dimensions", func() error {
			_, err := InsertChart(src, 0, sampleChartSpec(), 0, 100)
			return err
		}},
		{"out-of-range paragraph", func() error {
			_, err := InsertChart(src, 99, sampleChartSpec(), 100, 100)
			return err
		}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			if err := c.fn(); err == nil {
				t.Fatal("expected an error, got nil")
			}
		})
	}
}

func TestInsertChart_ScatterUsesNumericCategoriesAsXValues(t *testing.T) {
	src := buildChartTestDocx(t)
	spec := ChartWriteSpec{
		Type:   "scatter",
		Series: []ChartSeries{{Name: "pts", Categories: []string{"1", "2", "3"}, Values: []float64{4, 5, 6}}},
	}
	out, err := InsertChart(src, 0, spec, PixelsToEMU(300), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}
	chartXML, _ := zipPart(t, out, "word/charts/chart1.xml")
	if !strings.Contains(chartXML, "<c:xVal>") || !strings.Contains(chartXML, "<c:yVal>") {
		t.Fatalf("scatter chart missing xVal/yVal: %s", chartXML)
	}
}

func TestInsertChart_ScatterFallsBackWithoutXValsForNonNumericCategories(t *testing.T) {
	src := buildChartTestDocx(t)
	spec := ChartWriteSpec{
		Type:   "scatter",
		Series: []ChartSeries{{Name: "pts", Categories: []string{"a", "b"}, Values: []float64{4, 5}}},
	}
	out, err := InsertChart(src, 0, spec, PixelsToEMU(300), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}
	chartXML, _ := zipPart(t, out, "word/charts/chart1.xml")
	if strings.Contains(chartXML, "<c:xVal>") {
		t.Fatalf("non-numeric categories must not produce a bogus xVal: %s", chartXML)
	}
	if !strings.Contains(chartXML, "<c:yVal>") {
		t.Fatalf("yVal must still be present: %s", chartXML)
	}
}

// TestInsertChart_OpensWithPythonDocxOxml independently validates the
// produced .docx via python-docx's low-level docx.oxml/lxml access —
// python-docx has NO chart object model at all (not even the partial
// support it has for footnotes), so this follows the same pattern the
// footnotes validation established: python-docx's real OPC package/
// relationship machinery opens the file and resolves the chart
// relationship independently of anything this package wrote, then lxml
// parses the raw chart part and confirms the actual cached values (title,
// categories, and numeric values) are really there.
func TestInsertChart_OpensWithPythonDocxOxml(t *testing.T) {
	py, err := exec.LookPath("python3")
	if err != nil {
		t.Skip("python3 not available")
	}
	if err := exec.Command(py, "-c", "import docx, lxml.etree").Run(); err != nil {
		t.Skip("python-docx / lxml not installed")
	}

	out, err := InsertChart(buildChartTestDocx(t), 0, sampleChartSpec(), PixelsToEMU(500), PixelsToEMU(300))
	if err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	docxPath := filepath.Join(dir, "out.docx")
	if err := os.WriteFile(docxPath, out, 0o644); err != nil {
		t.Fatal(err)
	}

	script := `
import sys
import docx
from lxml import etree

d = docx.Document(sys.argv[1])
paras = [p.text for p in d.paragraphs]
assert paras[0] == "First paragraph.", paras
assert paras[-1] == "Second paragraph.", paras

chart_rel = None
for rel in d.part.rels.values():
    if rel.reltype.endswith("/chart"):
        chart_rel = rel
        break
assert chart_rel is not None, "no chart relationship found by python-docx"
chart_xml = chart_rel.target_part.blob

ns = {"c": "http://schemas.openxmlformats.org/drawingml/2006/chart", "a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
root = etree.fromstring(chart_xml)
bar = root.find(".//c:barChart", ns)
assert bar is not None, etree.tostring(root)

title_text = "".join(t.text or "" for t in root.iter("{%s}t" % ns["a"]))
assert "Quarterly Revenue" in title_text, title_text

cats = [pt.find("c:v", ns).text for pt in root.findall(".//c:cat//c:pt", ns)]
assert cats == ["Q1", "Q2", "Q3"], cats
vals = [float(pt.find("c:v", ns).text) for pt in root.findall(".//c:val//c:pt", ns)]
assert vals == [10.5, 12.0, 15.25], vals

print("OK")
`
	cmd := exec.Command(py, "-c", script, docxPath)
	outBytes, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("python-docx/lxml validation failed: %v\n%s", err, outBytes)
	}
	if !strings.Contains(string(outBytes), "OK") {
		t.Fatalf("unexpected validation output: %s", outBytes)
	}
}
