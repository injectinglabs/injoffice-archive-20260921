package docxpatch

import (
	"encoding/xml"
	"strings"
)

// A separate local shape model, never a picture, paragraph layout or edit target.
type NativeTextboxGeometryV1 struct {
	WidthEMU           int64    `json:"width_emu"`
	HeightEMU          int64    `json:"height_emu"`
	InsetsEMU          [4]int64 `json:"insets_emu"` // left, top, right, bottom
	FillRGB            string   `json:"fill_rgb"`   // explicit sRGB or none
	LineRGB            string   `json:"line_rgb"`
	LineWidthEMU       int64    `json:"line_width_emu"`
	FontFamily         string   `json:"font_family"`
	FontSizeHalfPoints int64    `json:"font_size_half_points"`
	TextRGB            string   `json:"text_rgb"`
	TextWrap           string   `json:"text_wrap,omitempty"`
}
type NativeTextboxGeometryItemV1 struct {
	Owner           NativePartialTextboxV1          `json:"owner"`
	Geometry        *NativeTextboxGeometryV1        `json:"geometry"`
	HardBreakLayout *NativeTextboxHardBreakLayoutV1 `json:"hard_break_layout,omitempty"`
	WrapLayout      *NativeTextboxWrapLayoutV1      `json:"wrap_layout,omitempty"`
	PageAnchor      *NativeTextboxPageAnchorV1      `json:"page_anchor,omitempty"`
}
type NativeTextboxGeometryEvidenceV1 struct {
	Items        []NativeTextboxGeometryItemV1 `json:"items"`
	OmittedCount int                           `json:"omitted_count"`
}

func inspectNativeTextboxGeometry(data []byte, doc *NativeDocumentV1) (*NativeTextboxGeometryEvidenceV1, error) {
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	ns, main := resolver.wordNS, resolver.mainPart
	raw := resolver.pkg.files[main]
	out := &NativeTextboxGeometryEvidenceV1{Items: []NativeTextboxGeometryItemV1{}}
	contextOK := nativeExactContainer(resolver.mainRoot) && len(resolver.diagnostics) == 0
	if resolver.parts.StylesPart != nil {
		styles, err := parseNativeXML(*resolver.parts.StylesPart, resolver.pkg.files[*resolver.parts.StylesPart])
		if err != nil {
			return nil, err
		}
		contextOK = contextOK && styles.Name == (xml.Name{Space: ns, Local: "styles"}) && nativeExactLeaf(styles)
	}
	for _, d := range doc.Unsupported {
		if d.Capability != "drawings" && d.Code != "DEFAULT_SECTION_INFERRED" {
			contextOK = false
		}
	}
	nodes := map[string]*nativeXMLNode{}
	for _, n := range resolver.nodeByAnchor {
		nodes[n.Path] = n
	}
	units := 0
	for _, block := range doc.Body.Blocks {
		p := block.Paragraph
		if p == nil {
			continue
		}
		owner := nodes[p.Anchor.Path]
		if owner == nil || owner.parent == nil {
			continue
		}
		for _, d := range doc.Unsupported {
			if d.Anchor == nil || d.Capability != "drawings" || d.ScopeID != p.ID || d.Anchor.PartName != main {
				continue
			}
			n := nodes[d.Anchor.Path]
			if n == nil || nativeSHA(raw[n.Start:n.End]) != d.Anchor.XMLSHA256 {
				continue
			}
			drawing := n
			for depth := 0; drawing != nil && drawing.Name != (xml.Name{Space: ns, Local: "drawing"}) && depth < 4; depth++ {
				drawing = drawing.parent
			}
			if drawing == nil || drawing.Name != (xml.Name{Space: ns, Local: "drawing"}) || drawing.parent == nil || drawing.parent.Name != (xml.Name{Space: ns, Local: "r"}) || drawing.parent.parent != owner || !nativePartialTextboxDiagnosticAnchor(drawing, n, ns) || len(nativeDescendants(drawing, ns, "txbxContent")) != 1 {
				continue
			}
			if len(out.Items) >= 64 {
				out.OmittedCount++
				continue
			}
			item := NativeTextboxGeometryItemV1{Owner: NativePartialTextboxV1{doc.Source.PackageSHA256, nativeSHA(raw), p.ID, d.ID, *d.Anchor, "drawingml", "omitted", []string{}, "unsupported-shape-geometry"}}
			if contextOK && nativeExactContainer(owner.parent) && nativePartialTextboxOwner(owner, drawing.parent, ns) && nativeGeometryOwnerProperties(owner, drawing.parent, ns) {
				geometry, text := nativeParseTextboxGeometry(drawing, ns)
				if geometry != nil && units+len(text) <= 100000 {
					item.Geometry = geometry
					item.HardBreakLayout = nativeTextboxHardBreakEvidence(drawing, ns, main, raw)
					item.WrapLayout = nativeTextboxWrapEvidence(drawing, ns, main, raw)
					item.PageAnchor = nativeTextboxPageAnchorEvidence(drawing, ns, main, raw)
					item.Owner.Status = "supported"
					item.Owner.Reason = ""
					item.Owner.Paragraphs = []string{text}
					units += len(text)
				}
			}
			out.Items = append(out.Items, item)
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}
func nativeGeometryOwnerProperties(p, r *nativeXMLNode, ns string) bool {
	for _, pair := range []struct {
		n     *nativeXMLNode
		local string
	}{{p, "pPr"}, {r, "rPr"}} {
		for _, c := range directNativeChildren(pair.n, ns, pair.local) {
			if !nativeExactLeaf(c) {
				return false
			}
		}
	}
	return true
}
func nativeGeometryChildren(n *nativeXMLNode, ns string, names ...string) bool {
	if len(n.Children) != len(names) {
		return false
	}
	for i, name := range names {
		if n.Children[i].Name != (xml.Name{Space: ns, Local: name}) {
			return false
		}
	}
	return true
}
func nativeGeometryAttrs(n *nativeXMLNode, ns string, values map[string]string) bool {
	attrs := []xml.Name{}
	for key := range values {
		attrs = append(attrs, xml.Name{Space: ns, Local: key})
	}
	if !nativeExactContainer(n, attrs...) {
		return false
	}
	for key, want := range values {
		value, ok := nativeAttr(n, ns, key)
		if !ok || value != want {
			return false
		}
	}
	return true
}
func nativeGeometryPaint(n *nativeXMLNode, a string) (string, bool) {
	if n.Name == (xml.Name{Space: a, Local: "noFill"}) && nativeExactLeaf(n) {
		return "none", true
	}
	if n.Name != (xml.Name{Space: a, Local: "solidFill"}) || !nativeExactContainer(n) || !nativeGeometryChildren(n, a, "srgbClr") {
		return "", false
	}
	c := n.Children[0]
	value, ok := nativeAttr(c, "", "val")
	if !ok || !nativeExactLeaf(c, xml.Name{Local: "val"}) {
		return "", false
	}
	rgb, valid := nativeExactRGB(value)
	return rgb, valid
}
func nativeParseTextboxGeometry(drawing *nativeXMLNode, ns string) (*NativeTextboxGeometryV1, string) {
	wp, a := wordDrawingTransitional, drawingMLTransitional
	if ns == wordMLStrict {
		wp, a = wordDrawingStrict, drawingMLStrict
	}
	if !nativeExactContainer(drawing) || len(drawing.Children) != 1 {
		return nil, ""
	}
	inline := drawing.Children[0]
	var extent, effect, docPr, graphic *nativeXMLNode
	if inline.Name == (xml.Name{Space: wp, Local: "inline"}) {
		if !nativeGeometryAttrs(inline, "", map[string]string{"distT": "0", "distB": "0", "distL": "0", "distR": "0"}) || len(inline.Children) != 4 {
			return nil, ""
		}
		extent, effect, docPr, graphic = inline.Children[0], inline.Children[1], inline.Children[2], inline.Children[3]
	} else {
		if _, _, ok := nativeTextboxPageOffsets(inline, wp, a); !ok {
			return nil, ""
		}
		extent, effect, docPr, graphic = inline.Children[3], inline.Children[4], inline.Children[6], inline.Children[7]
	}
	if extent.Name != (xml.Name{Space: wp, Local: "extent"}) || !nativeExactLeaf(extent, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) || effect.Name != (xml.Name{Space: wp, Local: "effectExtent"}) || !nativeGeometryAttrs(effect, "", map[string]string{"l": "0", "t": "0", "r": "0", "b": "0"}) || len(effect.Children) != 0 || docPr.Name != (xml.Name{Space: wp, Local: "docPr"}) || !nativeExactLeaf(docPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}) {
		return nil, ""
	}
	width, wOK := nativePositiveInt64Attr(extent, "", "cx")
	height, hOK := nativePositiveInt64Attr(extent, "", "cy")
	if !wOK || !hOK || width > 127000000 || height > 127000000 {
		return nil, ""
	}
	if graphic.Name != (xml.Name{Space: a, Local: "graphic"}) || !nativeExactContainer(graphic) || !nativeGeometryChildren(graphic, a, "graphicData") {
		return nil, ""
	}
	gd := graphic.Children[0]
	if !nativeGeometryAttrs(gd, "", map[string]string{"uri": nativeTextboxWPS}) || !nativeGeometryChildren(gd, nativeTextboxWPS, "wsp") {
		return nil, ""
	}
	shape := gd.Children[0]
	if !nativeExactContainer(shape) || !nativeGeometryChildren(shape, nativeTextboxWPS, "cNvSpPr", "spPr", "txbx", "bodyPr") {
		return nil, ""
	}
	if !nativeGeometryAttrs(shape.Children[0], "", map[string]string{"txBox": "1"}) || len(shape.Children[0].Children) != 0 {
		return nil, ""
	}
	sp := shape.Children[1]
	if !nativeExactContainer(sp) || len(sp.Children) != 4 {
		return nil, ""
	}
	xfrm, geom, fill, line := sp.Children[0], sp.Children[1], sp.Children[2], sp.Children[3]
	if xfrm.Name != (xml.Name{Space: a, Local: "xfrm"}) || !nativeGeometryAttrs(xfrm, "", map[string]string{"rot": "0", "flipH": "0", "flipV": "0"}) || !nativeGeometryChildren(xfrm, a, "off", "ext") {
		return nil, ""
	}
	if !nativeGeometryAttrs(xfrm.Children[0], "", map[string]string{"x": "0", "y": "0"}) || len(xfrm.Children[0].Children) != 0 || !nativeExactLeaf(xfrm.Children[1], xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) {
		return nil, ""
	}
	innerW, okW := nativePositiveInt64Attr(xfrm.Children[1], "", "cx")
	innerH, okH := nativePositiveInt64Attr(xfrm.Children[1], "", "cy")
	if !okW || !okH || innerW != width || innerH != height {
		return nil, ""
	}
	if geom.Name != (xml.Name{Space: a, Local: "prstGeom"}) || !nativeGeometryAttrs(geom, "", map[string]string{"prst": "rect"}) || !nativeGeometryChildren(geom, a, "avLst") || !nativeExactLeaf(geom.Children[0]) {
		return nil, ""
	}
	fillRGB, fillOK := nativeGeometryPaint(fill, a)
	if !fillOK || line.Name != (xml.Name{Space: a, Local: "ln"}) {
		return nil, ""
	}
	lineRGB := "none"
	lineWidth := int64(0)
	if nativeExactContainer(line) && nativeGeometryChildren(line, a, "noFill") && nativeExactLeaf(line.Children[0]) {
	} else {
		if !nativeExactContainer(line, xml.Name{Local: "w"}, xml.Name{Local: "cap"}, xml.Name{Local: "cmpd"}, xml.Name{Local: "algn"}) || len(line.Children) != 3 {
			return nil, ""
		}
		for key, want := range map[string]string{"cap": "flat", "cmpd": "sng", "algn": "ctr"} {
			value, ok := nativeAttr(line, "", key)
			if !ok || value != want {
				return nil, ""
			}
		}
		var ok bool
		lineWidth, ok = nativePositiveInt64Attr(line, "", "w")
		if !ok || lineWidth > 127000 || lineWidth >= width || lineWidth >= height {
			return nil, ""
		}
		lineRGB, ok = nativeGeometryPaint(line.Children[0], a)
		if !ok || lineRGB == "none" || line.Children[1].Name != (xml.Name{Space: a, Local: "prstDash"}) || !nativeGeometryAttrs(line.Children[1], "", map[string]string{"val": "solid"}) || len(line.Children[1].Children) != 0 || line.Children[2].Name != (xml.Name{Space: a, Local: "miter"}) || !nativeGeometryAttrs(line.Children[2], "", map[string]string{"lim": "800000"}) || len(line.Children[2].Children) != 0 {
			return nil, ""
		}
	}
	body := shape.Children[3]
	keys := map[string]string{"vert": "horz", "anchor": "t", "anchorCtr": "0", "wrap": "none", "numCol": "1", "rot": "0", "spcFirstLastPara": "0", "vertOverflow": "overflow", "horzOverflow": "overflow"}
	wrap, _ := nativeAttr(body, "", "wrap")
	if wrap == "square" {
		keys["wrap"] = "square"
	}
	attrs := []xml.Name{}
	for key := range keys {
		attrs = append(attrs, xml.Name{Local: key})
	}
	for _, key := range []string{"lIns", "tIns", "rIns", "bIns"} {
		attrs = append(attrs, xml.Name{Local: key})
	}
	if !nativeExactContainer(body, attrs...) || !nativeGeometryChildren(body, a, "noAutofit") || !nativeExactLeaf(body.Children[0]) {
		return nil, ""
	}
	for key, want := range keys {
		value, ok := nativeAttr(body, "", key)
		if !ok || value != want {
			return nil, ""
		}
	}
	insets := [4]int64{}
	for i, key := range []string{"lIns", "tIns", "rIns", "bIns"} {
		value, ok := nativeNonnegativeInt64Attr(body, "", key)
		if !ok || value > 127000000 {
			return nil, ""
		}
		insets[i] = value
	}
	if insets[0]+insets[2] >= width || insets[1]+insets[3] >= height {
		return nil, ""
	}
	box := shape.Children[2]
	if !nativeExactContainer(box) || !nativeGeometryChildren(box, ns, "txbxContent") {
		return nil, ""
	}
	content := box.Children[0]
	if !nativeExactContainer(content) || !nativeGeometryChildren(content, ns, "p") {
		return nil, ""
	}
	p := content.Children[0]
	if !nativeExactContainer(p) || !nativeGeometryChildren(p, ns, "pPr", "r") {
		return nil, ""
	}
	pp := p.Children[0]
	if !nativeExactContainer(pp) || !nativeGeometryChildren(pp, ns, "spacing", "ind", "jc") {
		return nil, ""
	}
	for i, values := range []map[string]string{{"left": "0", "right": "0", "firstLine": "0"}, {"val": "left"}} {
		if !nativeGeometryAttrs(pp.Children[i+1], ns, values) || len(pp.Children[i+1].Children) != 0 {
			return nil, ""
		}
	}
	r := p.Children[1]
	if !nativeExactContainer(r) || len(r.Children) < 2 || r.Children[0].Name != (xml.Name{Space: ns, Local: "rPr"}) {
		return nil, ""
	}
	rp := r.Children[0]
	if !nativeExactContainer(rp) || !nativeGeometryChildren(rp, ns, "rFonts", "b", "i", "color", "sz", "lang") {
		return nil, ""
	}
	rf := rp.Children[0]
	if !nativeExactLeaf(rf, xml.Name{Space: ns, Local: "ascii"}, xml.Name{Space: ns, Local: "hAnsi"}) {
		return nil, ""
	}
	font, fOK := nativeAttr(rf, ns, "ascii")
	other, oOK := nativeAttr(rf, ns, "hAnsi")
	if !fOK || !oOK || font != other || font == "" || len(font) > 128 || strings.TrimSpace(font) != font {
		return nil, ""
	}
	for _, c := range rp.Children[1:3] {
		if !nativeGeometryAttrs(c, ns, map[string]string{"val": "0"}) || len(c.Children) != 0 {
			return nil, ""
		}
	}
	color := rp.Children[3]
	rgb, _ := nativeAttr(color, ns, "val")
	rgb, cOK := nativeExactRGB(rgb)
	if !cOK || !nativeExactLeaf(color, xml.Name{Space: ns, Local: "val"}) {
		return nil, ""
	}
	sizeNode := rp.Children[4]
	size, sOK := nativePositiveInt64Attr(sizeNode, ns, "val")
	if !sOK || size > 400 || !nativeExactLeaf(sizeNode, xml.Name{Space: ns, Local: "val"}) || !nativeGeometryAttrs(rp.Children[5], ns, map[string]string{"val": "en-US"}) || len(rp.Children[5].Children) != 0 {
		return nil, ""
	}
	text, _, ok := nativeTextboxHardBreakText(p, ns)
	if wrap == "square" {
		text, _, ok = nativeTextboxWrapText(p, ns)
	}
	if !ok {
		return nil, ""
	}
	for _, v := range []int64{width, height, insets[0], insets[1], insets[2], insets[3], lineWidth} {
		if v%127 != 0 {
			return nil, ""
		}
	}
	textWrap := ""
	if wrap == "square" {
		textWrap = "square"
	}
	return &NativeTextboxGeometryV1{width, height, insets, fillRGB, lineRGB, lineWidth, font, size, rgb, textWrap}, text
}
