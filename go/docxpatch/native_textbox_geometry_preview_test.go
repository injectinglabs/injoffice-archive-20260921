package docxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func nativeGeometryFixture() string {
	return `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `"><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="2743200" cy="914400"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Rectangle"/><a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm rot="0" flipH="0" flipV="0"><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="204060"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/><w:ind w:left="0" w:right="0" w:firstLine="0"/><w:jc w:val="left"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:b w:val="0"/><w:i w:val="0"/><w:color w:val="102030"/><w:sz w:val="24"/><w:lang w:val="en-US"/></w:rPr><w:t xml:space="preserve">Rectangle source</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr vert="horz" anchor="t" anchorCtr="0" wrap="none" numCol="1" rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" lIns="91440" tIns="91440" rIns="91440" bIns="91440"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
}
func TestNativeTextboxGeometry(t *testing.T) {
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+nativeGeometryFixture()+`</w:r></w:p>`))))
	before := append([]byte(nil), source...)
	encoded, err := InspectNativePartialSourceV1(source)
	if err != nil {
		t.Fatal(err)
	}
	var result struct {
		Geometry *NativeTextboxGeometryEvidenceV1 `json:"textbox_geometry"`
		Document NativeDocumentV1                 `json:"document"`
	}
	if err = json.Unmarshal(encoded, &result); err != nil {
		t.Fatal(err)
	}
	if result.Geometry == nil || len(result.Geometry.Items) != 1 {
		t.Fatalf("missing evidence: %s", encoded)
	}
	item := result.Geometry.Items[0]
	if item.Geometry == nil || item.Owner.Status != "supported" || item.Geometry.WidthEMU != 2743200 || item.Geometry.InsetsEMU[0] != 91440 || item.Geometry.FillRGB != "FFF2CC" || item.Geometry.LineWidthEMU != 12700 || item.Owner.Paragraphs[0] != "Rectangle source" {
		t.Fatalf("bad geometry: %#v", item)
	}
	if !bytes.Equal(source, before) || result.Document.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" || !hasUnsupportedCode(&result.Document, "PICTURE_GRAPHIC_REQUIRED") {
		t.Fatal("source/mutation/diagnostics changed")
	}
}
func TestNativeTextboxGeometryRefusals(t *testing.T) {
	for _, test := range []struct{ name, from, to string }{
		{"autofit", "a:noAutofit", "a:spAutoFit"}, {"rotation", `rot="0"`, `rot="60000"`}, {"theme", `a:srgbClr val="FFF2CC"`, `a:schemeClr val="accent1"`},
		{"hidden", `<w:b w:val="0"/>`, `<w:vanish/><w:b w:val="0"/>`}, {"extent mismatch", `<a:ext cx="2743200"`, `<a:ext cx="2743201"`}, {"nonexact geometry", `cx="2743200"`, `cx="2743201"`},
		{"paragraph style", `<w:pPr>`, `<w:pPr><w:pStyle w:val="Unknown"/>`}, {"default whitespace", ` xml:space="preserve"`, ``}, {"text newline", `Rectangle source`, "Line\nTwo"},
		{"line alignment", `algn="ctr"`, `algn="in"`}, {"line inherited", `<a:prstDash val="solid"/>`, ``}, {"effects", `<wps:spPr>`, `<wps:spPr><a:effectLst/>`}, {"unknown body attr", `vert="horz"`, `unknown="1" vert="horz"`},
		{"inset overflow", `lIns="91440"`, `lIns="2743200"`}, {"unknown source rPr", `<w:rPr>`, `<w:rPr unknown="1">`},
	} {
		t.Run(test.name, func(t *testing.T) {
			drawing := strings.ReplaceAll(nativeGeometryFixture(), test.from, test.to)
			source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+drawing+`</w:r></w:p>`))))
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			out, err := inspectNativeTextboxGeometry(source, doc)
			if err != nil {
				t.Fatal(err)
			}
			if out != nil {
				for _, item := range out.Items {
					if item.Geometry != nil || item.Owner.Status == "supported" || len(item.Owner.Paragraphs) > 0 {
						t.Fatalf("unsafe content admitted: %#v", item)
					}
				}
			}
		})
	}
}
func TestNativeTextboxGeometryNoPaint(t *testing.T) {
	drawing := strings.Replace(nativeGeometryFixture(), `<a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill>`, `<a:noFill/>`, 1)
	start := strings.Index(drawing, `<a:ln `)
	end := strings.Index(drawing, `</a:ln>`) + len(`</a:ln>`)
	drawing = drawing[:start] + `<a:ln><a:noFill/></a:ln>` + drawing[end:]
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+drawing+`</w:r></w:p>`))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	out, err := inspectNativeTextboxGeometry(source, doc)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || out.Items[0].Geometry == nil || out.Items[0].Geometry.FillRGB != "none" || out.Items[0].Geometry.LineWidthEMU != 0 {
		t.Fatal("explicit no-paint geometry was lost")
	}
}
