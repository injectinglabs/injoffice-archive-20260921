package pptxpatch

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
)

func chartZipEntry(t *testing.T, data []byte, name string) []byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name != name {
			continue
		}
		r, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		defer r.Close() //nolint:errcheck
		out, err := io.ReadAll(r)
		if err != nil {
			t.Fatal(err)
		}
		return out
	}
	t.Fatalf("missing zip entry %q", name)
	return nil
}

func replaceChartZipEntry(t *testing.T, data []byte, name string, replacement []byte) []byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	found := false
	for _, f := range zr.File {
		w, err := zw.Create(f.Name)
		if err != nil {
			t.Fatal(err)
		}
		if f.Name == name {
			found = true
			if _, err := w.Write(replacement); err != nil {
				t.Fatal(err)
			}
			continue
		}
		r, err := f.Open()
		if err != nil {
			t.Fatal(err)
		}
		body, err := io.ReadAll(r)
		r.Close() //nolint:errcheck
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write(body); err != nil {
			t.Fatal(err)
		}
	}
	if !found {
		t.Fatalf("cannot replace missing zip entry %q", name)
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func chartFixture(t *testing.T) ([]byte, []byte) {
	t.Helper()
	chartXML := []byte(`<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:externalData r:id="rId1"/><c:chart><c:plotArea/></c:chart></c:chartSpace>`)
	embedded := []byte{0x50, 0x4b, 0x03, 0x04, 0x00, 0xff}
	data, err := BuildPPTX(Deck{Slides: []Slide{{Shapes: []Shape{{
		Kind: KindChart, Name: "Quarterly revenue", X: 100, Y: 200, Cx: 300, Cy: 400,
		Chart: &ChartPart{RelationshipID: "rId9", PartName: "ppt/charts/chart1.xml", XML: chartXML,
			Relationships: []ChartRelationship{{ID: "rId1", Type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package", Target: "../embeddings/book1.xlsx"}},
			EmbeddedParts: []ChartEmbeddedPart{{Name: "ppt/embeddings/book1.xlsx", Data: embedded, ContentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}},
		},
	}}}}})
	if err != nil {
		t.Fatal(err)
	}
	return data, chartXML
}

func TestOpaqueChartGraphicFrameRoundTripsCompleteLocalGraph(t *testing.T) {
	data, chartXML := chartFixture(t)
	if got := chartZipEntry(t, data, "ppt/charts/chart1.xml"); !bytes.Equal(got, chartXML) {
		t.Fatal("chart XML changed")
	}
	if rels := string(chartZipEntry(t, data, "ppt/slides/_rels/slide1.xml.rels")); !strings.Contains(rels, `Id="rId9"`) || !strings.Contains(rels, `Target="../charts/chart1.xml"`) {
		t.Fatalf("chart relationship missing: %s", rels)
	}
	if rels := string(chartZipEntry(t, data, "ppt/charts/_rels/chart1.xml.rels")); !strings.Contains(rels, `Target="../embeddings/book1.xlsx"`) {
		t.Fatalf("child relationship missing: %s", rels)
	}
	got, err := ParsePPTX(data)
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Slides) != 1 || len(got.Slides[0].Shapes) != 1 {
		t.Fatalf("chart frame not parsed: %+v", got)
	}
	shape := got.Slides[0].Shapes[0]
	if shape.Kind != KindChart || shape.Chart == nil || !bytes.Equal(shape.Chart.XML, chartXML) {
		t.Fatalf("chart frame mismatch: %+v", shape)
	}
	again, err := BuildPPTX(got)
	if err != nil {
		t.Fatal(err)
	}
	if rebuilt := chartZipEntry(t, again, "ppt/charts/chart1.xml"); !bytes.Equal(rebuilt, chartXML) {
		t.Fatal("chart XML did not survive parse/build")
	}
}

func TestParsePPTXSkipsChartWithExternalOrIncompleteGraph(t *testing.T) {
	data, _ := chartFixture(t)
	for name, rels := range map[string]string{
		"external": xmlDecl + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="package" Target="https://attacker.invalid/book.xlsx" TargetMode="External"/></Relationships>`,
		"missing":  xmlDecl + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="package" Target="../embeddings/missing.xlsx"/></Relationships>`,
	} {
		t.Run(name, func(t *testing.T) {
			got, err := ParsePPTX(replaceChartZipEntry(t, data, "ppt/charts/_rels/chart1.xml.rels", []byte(rels)))
			if err != nil {
				t.Fatal(err)
			}
			if len(got.Slides) != 1 || len(got.Slides[0].Shapes) != 0 {
				t.Fatalf("unsafe chart was not skipped: %+v", got)
			}
		})
	}
}
