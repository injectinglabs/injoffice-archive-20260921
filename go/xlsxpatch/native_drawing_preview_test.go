package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func drawingPreviewFixture(strict bool) (map[string]string, string) {
	parts := nativeWorkbookFixture(strict)
	r := officeRelNamespaceTransitional
	ss := spreadsheetMLTransitional
	xdr := "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
	a := "http://schemas.openxmlformats.org/drawingml/2006/main"
	c := "http://schemas.openxmlformats.org/drawingml/2006/chart"
	if strict {
		r = officeRelNamespaceStrict
		ss = spreadsheetMLStrict
		xdr = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
		a = "http://purl.oclc.org/ooxml/drawingml/main"
		c = "http://purl.oclc.org/ooxml/drawingml/chart"
	}
	parts["Charts/chart1.xml"] = strings.ReplaceAll(previewChartFixture(), "http://schemas.openxmlformats.org/drawingml/2006/chart", c)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Drawings/drawing.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`, 1)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</worksheet>`, `<drawing xmlns="`+ss+`" xmlns:r="`+r+`" r:id="drawing1"/></worksheet>`, 1)
	parts["Sheets/_rels/s1.xml.rels"] = strings.Replace(parts["Sheets/_rels/s1.xml.rels"], `</Relationships>`, `<Relationship Id="drawing1" Type="`+r+`/drawing" Target="../Drawings/drawing.xml"/></Relationships>`, 1)
	parts["Drawings/_rels/drawing.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="chart1" Type="` + r + `/chart" Target="../Charts/chart1.xml"/></Relationships>`
	drawing := `<xdr:wsDr xmlns:xdr="` + xdr + `" xmlns:a="` + a + `" xmlns:c="` + c + `" xmlns:r="` + r + `"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:colOff>100</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>200</xdr:rowOff></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:colOff>300</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>400</xdr:rowOff></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="` + c + `"><c:chart r:id="chart1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>`
	parts["Drawings/drawing.xml"] = drawing
	return parts, drawing
}
func TestNativeDrawingPreviewSourceOwnershipBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		parts, _ := drawingPreviewFixture(strict)
		source := buildZip(t, parts)
		before := bytes.Clone(source)
		o, e := InspectNativeWorkbookObjectsV1(source)
		if e != nil {
			t.Fatal(e)
		}
		if len(o.DrawingObjects) != 1 {
			t.Fatalf("missing source object %+v", o.DrawingObjects)
		}
		d := o.DrawingObjects[0]
		if d.Kind != "chart" || d.ChartPart != "Charts/chart1.xml" || d.SheetPart != "Sheets/s1.xml" || d.Anchor.From.ColumnOffset != 100 || d.Anchor.To.RowOffset != 400 {
			t.Fatalf("wrong projection %+v", d)
		}
		if !bytes.Equal(source, before) {
			t.Fatal("source modified")
		}
	}
}

func TestDrawingCreationExtensionExact(t *testing.T) {
	a := "http://schemas.openxmlformats.org/drawingml/2006/main"
	source := `<a:extLst xmlns:a="` + a + `"><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><x:creationId xmlns:x="http://schemas.microsoft.com/office/drawing/2014/main" id="{11111111-2222-3333-4444-555555555555}"/></a:ext></a:extLst>`
	n, e := parsePreviewXML([]byte(source))
	if e != nil || !previewDrawingCreationID(n, a) {
		t.Fatal("known identity extension refused")
	}
	for _, bad := range []string{strings.Replace(source, "creationId", "foreign", 1), strings.Replace(source, "FF2B5EF4", "00000000", 1), strings.Replace(source, `/></a:ext>`, `><x:paint/></x:creationId></a:ext>`, 1), strings.Replace(source, "11111111-2222-3333-4444-555555555555", "invalid", 1)} {
		n, e := parsePreviewXML([]byte(bad))
		if e == nil && previewDrawingCreationID(n, a) {
			t.Fatal("unknown extension qualified")
		}
	}
}
func TestNativeDrawingPreviewRetainsUnsupportedObjects(t *testing.T) {
	for _, kind := range []string{"offset", "duplicate", "rotation", "external-chart", "foreign-chart-rel", "foreign-anchor", "unknown-frame", "one-cell"} {
		t.Run(kind, func(t *testing.T) {
			parts, drawing := drawingPreviewFixture(false)
			switch kind {
			case "offset":
				drawing = strings.Replace(drawing, `<xdr:colOff>100</xdr:colOff>`, `<xdr:colOff>0x10</xdr:colOff>`, 1)
			case "duplicate":
				drawing = strings.Replace(drawing, `<xdr:colOff>100</xdr:colOff>`, `<xdr:colOff>100</xdr:colOff><xdr:colOff>100</xdr:colOff>`, 1)
			case "rotation":
				drawing = strings.Replace(drawing, `<xdr:xfrm>`, `<xdr:xfrm rot="60000">`, 1)
			case "external-chart":
				parts["Drawings/_rels/drawing.xml.rels"] = strings.Replace(parts["Drawings/_rels/drawing.xml.rels"], `Target="../Charts/chart1.xml"`, `Target="https://example.invalid/chart.xml" TargetMode="External"`, 1)
			case "foreign-chart-rel":
				parts["Drawings/_rels/drawing.xml.rels"] = strings.Replace(parts["Drawings/_rels/drawing.xml.rels"], officeRelNamespaceTransitional+"/chart", "urn:evil/chart", 1)
			case "foreign-anchor":
				drawing = strings.Replace(drawing, `<xdr:twoCellAnchor>`, `<xdr:twoCellAnchor xmlns:xdr="urn:foreign">`, 1)
			case "unknown-frame":
				drawing = strings.ReplaceAll(drawing, "xdr:graphicFrame", "xdr:sp")
			case "one-cell":
				start := strings.Index(drawing, "<xdr:to>")
				end := strings.Index(drawing, "</xdr:to>") + len("</xdr:to>")
				drawing = drawing[:start] + `<xdr:ext cx="2000000" cy="1000000"/>` + drawing[end:]
				drawing = strings.ReplaceAll(drawing, "twoCellAnchor", "oneCellAnchor")
			}
			parts["Drawings/drawing.xml"] = drawing
			o, e := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
			if e != nil {
				t.Fatal(e)
			}
			if len(o.DrawingObjects) != 1 {
				t.Fatal("lost unsupported source owner")
			}
			d := o.DrawingObjects[0]
			if kind == "one-cell" {
				if d.Kind != "chart" || d.Anchor.Width != 2000000 {
					t.Fatalf("lost extent %+v", d)
				}
			} else if d.Kind != "unsupported" || d.ChartPart != "" {
				t.Fatalf("unsupported rendered %+v", d)
			}
			if d.SheetPart != "Sheets/s1.xml" || len(d.Warnings) == 0 {
				t.Fatal("source ownership/warning lost")
			}
		})
	}
}
