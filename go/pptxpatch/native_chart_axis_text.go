package pptxpatch

import (
	"encoding/xml"
	"strings"
	"unicode"
	"unicode/utf8"
)

// A deliberately complete source style. All three script face declarations
// agree, so the renderer never treats an unspecified EA/CS font as Latin.
type nativeChartAxisLabelStyle = NativeChartAxisLabelStyle

func extractNativeChartAxisLabelStyle(node *nativeXMLNode, d nativeExtractDialect) (*nativeChartAxisLabelStyle, bool) {
	if node == nil || node.Name != (xml.Name{Space: d.chart, Local: "txPr"}) {
		return nil, false
	}
	text := nativeChartChildren(node, d.drawing)
	body := text.take("bodyPr")
	attrs := []xml.Name{}
	expected := map[string]string{"rot": "0", "vert": "horz", "wrap": "none", "lIns": "0", "rIns": "0", "tIns": "0", "bIns": "0", "anchor": "t", "anchorCtr": "0", "horzOverflow": "overflow", "vertOverflow": "overflow"}
	for key := range expected {
		attrs = append(attrs, xml.Name{Local: key})
	}
	if body == nil || requireOnlyNativeAttrs(body, attrs...) != nil || !onlyNativeXMLSpace(body.Text) {
		return nil, false
	}
	for key, want := range expected {
		if got, ok := exactNativeAttr(body, "", key); !ok || got != want {
			return nil, false
		}
	}
	// The cursor's container attribute check is intentionally separate: bodyPr
	// has the complete style attributes above, unlike ordinary empty containers.
	if len(body.Children) != 1 || body.Children[0].Name != (xml.Name{Space: d.drawing, Local: "noAutofit"}) || requireEmptyNativeElement(body.Children[0]) != nil {
		return nil, false
	}
	if requireEmptyNativeElement(text.take("lstStyle")) != nil {
		return nil, false
	}
	paragraph := nativeChartChildren(text.take("p"), d.drawing)
	properties := paragraph.take("pPr")
	if properties == nil || requireOnlyNativeAttrs(properties, xml.Name{Local: "algn"}) != nil || !onlyNativeXMLSpace(properties.Text) {
		return nil, false
	}
	if align, ok := exactNativeAttr(properties, "", "algn"); !ok || align != "ctr" {
		return nil, false
	}
	if len(properties.Children) != 1 || properties.Children[0].Name != (xml.Name{Space: d.drawing, Local: "defRPr"}) || !paragraph.done() || !text.done() {
		return nil, false
	}
	style := properties.Children[0]
	if requireOnlyNativeAttrs(style, xml.Name{Local: "b"}, xml.Name{Local: "i"}, xml.Name{Local: "sz"}, xml.Name{Local: "lang"}, xml.Name{Local: "kern"}) != nil || !onlyNativeXMLSpace(style.Text) {
		return nil, false
	}
	out := &nativeChartAxisLabelStyle{}
	for _, item := range []struct {
		name   string
		target *bool
	}{{"b", &out.Bold}, {"i", &out.Italic}} {
		raw, ok := exactNativeAttr(style, "", item.name)
		if !ok || (raw != "0" && raw != "1") {
			return nil, false
		}
		*item.target = raw == "1"
	}
	size, ok := exactNativeAttr(style, "", "sz")
	if !ok {
		return nil, false
	}
	var err error
	out.FontSize, err = parseCanonicalNativeInt(size, 1, 400000)
	if err != nil {
		return nil, false
	}
	if kern, ok := exactNativeAttr(style, "", "kern"); !ok || kern != "0" {
		return nil, false
	}
	out.Language, ok = exactNativeAttr(style, "", "lang")
	if !ok || !validNativeLanguage(out.Language) {
		return nil, false
	}
	if len(style.Children) != 4 {
		return nil, false
	}
	fill := style.Children[0]
	if fill.Name != (xml.Name{Space: d.drawing, Local: "solidFill"}) || !nativeLiteralPiePaintSequence(fill, d.drawing, "srgbClr") {
		return nil, false
	}
	color, ok := nativeChartAttribute(fill.Children[0])
	if !ok || !inspectionRGB.MatchString(color) {
		return nil, false
	}
	out.Color = strings.ToUpper(color)
	for i, name := range []string{"latin", "ea", "cs"} {
		face := style.Children[i+1]
		if face.Name != (xml.Name{Space: d.drawing, Local: name}) || requireOnlyNativeAttrs(face, xml.Name{Local: "typeface"}) != nil || requireOnlyNativeChildren(face) != nil {
			return nil, false
		}
		family, ok := exactNativeAttr(face, "", "typeface")
		if !ok || family == "" || len(family) > 128 || !utf8.ValidString(family) || strings.HasPrefix(family, "+") || strings.IndexFunc(family, unicode.IsControl) >= 0 {
			return nil, false
		}
		if i == 0 {
			out.FontFamily = family
		} else if family != out.FontFamily {
			return nil, false
		}
	}
	return out, true
}
