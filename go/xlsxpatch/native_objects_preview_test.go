package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func previewChartFixture() string {
	return `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:ser><c:tx><c:v>Saved series</c:v></c:tx><c:val><c:numRef><c:f>[external.xlsx]Sheet1!A1:A2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>1.25</c:v></c:pt><c:pt idx="1"><c:v>-3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`
}
func TestNativeObjectsPreviewUsesSavedCachesWithoutMutation(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Charts/chart1.xml"] = previewChartFixture()
	source := buildZip(t, parts)
	before := bytes.Clone(source)
	result, err := InspectNativeWorkbookObjectsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before, source) {
		t.Fatal("inspection mutated original")
	}
	if len(result.Charts) != 1 || result.Charts[0].Type != "col" || len(result.Charts[0].Series) != 1 {
		t.Fatalf("unexpected projection: %+v", result)
	}
	s := result.Charts[0].Series[0]
	if s.Name != "Saved series" || *s.Values[0] != 1.25 || *s.Values[1] != -3 {
		t.Fatalf("lost saved cache: %+v", s)
	}
	workbook, err := ExtractNativeWorkbookV2(source)
	if err != nil || result.PackageSHA256 != workbook.Source.PackageSHA256 {
		t.Fatal("unbound object source")
	}
}
func TestNativeObjectCachesRefuseUnsupportedOrAmbiguousShapes(t *testing.T) {
	for _, edit := range [][2]string{{`val="clustered"`, `val="stacked"`}, {`<c:ptCount val="2"/>`, `<c:ptCount val="3"/>`}, {`idx="1"`, `idx="0"`}, {`<c:v>-3</c:v>`, `<c:v>NaN</c:v>`}, {`<c:ptCount val="2"/>`, `<c:ptCount val="20000"/>`}, {`<c:barChart>`, `<c:bar3DChart>`}, {`<c:val>`, `<c:val xmlns:c="urn:spoof">`}} {
		raw := strings.Replace(previewChartFixture(), edit[0], edit[1], 1)
		if edit[1] == `<c:bar3DChart>` {
			raw = strings.Replace(raw, `</c:barChart>`, `</c:bar3DChart>`, 1)
		}
		root, err := parsePreviewXML([]byte(raw))
		if err != nil {
			continue
		}
		got := previewChart(root, "chart.xml")
		if got.Type != "unsupported" {
			t.Fatalf("unsupported cache painted for %v: %+v", edit, got)
		}
	}
	if _, err := parsePreviewXML([]byte(`<!DOCTYPE x [<!ENTITY e SYSTEM "file:///etc/passwd">]><x>&e;</x>`)); err == nil {
		t.Fatal("DTD accepted")
	}
}

func TestNativeTablePreviewRequiresExactActiveRelationship(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, foreign := range []bool{false, true} {
			for _, active := range []bool{false, true} {
				parts := nativeWorkbookFixture(strict)
				ss, rel := spreadsheetMLTransitional, officeRelNamespaceTransitional
				if strict {
					ss, rel = spreadsheetMLStrict, officeRelNamespaceStrict
				}
				parts["Charts/chart1.xml"] = previewChartFixture()
				parts["Tables/table.xml"] = `<table xmlns="` + ss + `" displayName="Original" ref="A1:C2"><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Tables/table.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>`, 1)
				relType := rel + "/table"
				if foreign {
					relType = "urn:foreign/table"
				}
				parts["Sheets/_rels/s1.xml.rels"] = strings.Replace(parts["Sheets/_rels/s1.xml.rels"], `</Relationships>`, `<Relationship Id="table1" Type="`+relType+`" Target="../Tables/table.xml"/></Relationships>`, 1)
				if active {
					parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</worksheet>`, `<tableParts count="1"><tablePart xmlns:r="`+rel+`" r:id="table1"/></tableParts></worksheet>`, 1)
				}
				got, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
				if err != nil {
					t.Fatal(err)
				}
				if len(got.Tables) != 1 {
					t.Fatal("lost table")
				}
				if (got.Tables[0].SheetPart != "") != (active && !foreign) {
					t.Fatalf("false table ownership strict=%v foreign=%v active=%v: %+v", strict, foreign, active, got.Tables[0])
				}
			}
		}
	}
}
