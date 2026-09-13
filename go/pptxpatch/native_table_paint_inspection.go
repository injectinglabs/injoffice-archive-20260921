package pptxpatch

import (
	"encoding/xml"
	"strings"
)

// This evidence qualifies only a no-fill one-cell style cascade and four
// identical source solid borders. Host text/layout remains a separate policy.
type NativePPTXTablePaintSource struct {
	PartName string `json:"part_name"`
	SHA256   string `json:"sha256"`
}
type NativePPTXTablePaintBorder struct {
	Color    string `json:"color"`
	WidthEMU int64  `json:"width_emu"`
	Preset   string `json:"preset,omitempty"`
}
type NativePPTXTablePaint struct {
	Policy  string                       `json:"policy"`
	StyleID string                       `json:"style_id"`
	Sources []NativePPTXTablePaintSource `json:"sources"`
	Fill    string                       `json:"fill"`
	Border  *NativePPTXTablePaintBorder  `json:"border"`
}
type nativeTablePaintContext struct {
	styles    map[string]*nativeXMLNode
	theme     nativeResolvedTheme
	sources   []NativePPTXTablePaintSource
	qualified bool
}

func (extractor *nativeExtractor) tablePaintContext(slidePart string, slideRoot *nativeXMLNode, d nativeExtractDialect) nativeTablePaintContext {
	c := nativeTablePaintContext{}
	graph, err := extractor.resolveSlideDependencyGraph(slidePart, d)
	if err != nil {
		return c
	}
	// No guessed precedence for slide/layout color overrides.
	for _, root := range []*nativeXMLNode{slideRoot, graph.layoutRoot} {
		mapping, e := nativeSingleton(root, d.presentation, "clrMapOvr", false)
		if e != nil {
			return c
		}
		if mapping != nil {
			if !onlyNativeXMLSpace(mapping.Text) || requireOnlyNativeAttrs(mapping) != nil || len(mapping.Children) != 1 || mapping.Children[0].Name != (xml.Name{Space: d.drawing, Local: "masterClrMapping"}) || requireEmptyNativeElement(mapping.Children[0]) != nil {
				return c
			}
		}
	}
	theme, err := resolveNativeTheme(graph, d)
	if err != nil {
		return c
	}
	roots, err := extractor.parseRelationships("")
	if err != nil {
		return c
	}
	office, err := uniqueNativeInternalRelationship(roots, d.rels+"/officeDocument", "presentation")
	if err != nil {
		return c
	}
	rels, err := extractor.parseRelationships(office.Part)
	if err != nil {
		return c
	}
	rel, err := uniqueNativeInternalRelationship(rels, d.rels+"/tableStyles", "table styles")
	if err != nil {
		return c
	}
	if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(rel.Part), "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml") {
		return c
	}
	payload := extractor.pkg.parts[rel.Part]
	if len(payload) > 128*1024 {
		return c
	}
	root, err := parseNativeXML(payload, rel.Part)
	if err != nil {
		return c
	}
	if root.Name != (xml.Name{Space: d.drawing, Local: "tblStyleLst"}) || requireOnlyNativeAttrs(root, xml.Name{Local: "def"}) != nil || requireOnlyNativeChildren(root, xml.Name{Space: d.drawing, Local: "tblStyle"}) != nil || len(root.Children) > 256 {
		return c
	}
	styles := map[string]*nativeXMLNode{}
	for _, style := range root.Children {
		id, ok := exactNativeAttr(style, "", "styleId")
		if !ok || !inspectionUUID.MatchString(id) || styles[strings.ToUpper(id)] != nil {
			return c
		}
		styles[strings.ToUpper(id)] = style
	}
	c.styles, c.theme, c.qualified = styles, theme, true
	for _, part := range []string{rel.Part, graph.themePart, graph.masterPart, graph.layoutPart} {
		if part == "" {
			return nativeTablePaintContext{}
		}
		c.sources = append(c.sources, NativePPTXTablePaintSource{part, nativeSHA256(extractor.pkg.parts[part])})
	}
	return c
}

func nativeTableNoStylePaint(style *nativeXMLNode, d nativeExtractDialect) bool {
	if style == nil || !nativeTablePaintWhitespace(style) || requireOnlyNativeAttrs(style, xml.Name{Local: "styleId"}, xml.Name{Local: "styleName"}) != nil || !nativePaintOnlyChild(style, d, "wholeTbl") {
		return false
	}
	whole := style.Children[0]
	if requireOnlyNativeAttrs(whole) != nil || len(whole.Children) != 2 || whole.Children[0].Name != (xml.Name{Space: d.drawing, Local: "tcTxStyle"}) || whole.Children[1].Name != (xml.Name{Space: d.drawing, Local: "tcStyle"}) {
		return false
	}
	// Text metadata is structurally qualified but deliberately not applied by the
	// host text policy. No effects, conditional style regions, or hidden subtree.
	text := whole.Children[0]
	if requireOnlyNativeAttrs(text) != nil || len(text.Children) != 2 || text.Children[0].Name != (xml.Name{Space: d.drawing, Local: "fontRef"}) || text.Children[1].Name != (xml.Name{Space: d.drawing, Local: "schemeClr"}) {
		return false
	}
	font := text.Children[0]
	idx, ok := exactNativeAttr(font, "", "idx")
	if !ok || idx != "minor" || requireOnlyNativeAttrs(font, xml.Name{Local: "idx"}) != nil || len(font.Children) != 1 {
		return false
	}
	color := font.Children[0]
	if color.Name != (xml.Name{Space: d.drawing, Local: "scrgbClr"}) || requireOnlyNativeAttrs(color, xml.Name{Local: "r"}, xml.Name{Local: "g"}, xml.Name{Local: "b"}) != nil || len(color.Children) != 0 {
		return false
	}
	for _, name := range []string{"r", "g", "b"} {
		v, ok := exactNativeAttr(color, "", name)
		if !ok || v != "0" {
			return false
		}
	}
	scheme := text.Children[1]
	slot, ok := exactNativeAttr(scheme, "", "val")
	if !ok || slot != "tx1" || requireOnlyNativeAttrs(scheme, xml.Name{Local: "val"}) != nil || len(scheme.Children) != 0 {
		return false
	}
	cell := whole.Children[1]
	if requireOnlyNativeAttrs(cell) != nil || len(cell.Children) != 2 || cell.Children[0].Name != (xml.Name{Space: d.drawing, Local: "tcBdr"}) || cell.Children[1].Name != (xml.Name{Space: d.drawing, Local: "fill"}) || requireOnlyNativeAttrs(cell.Children[1]) != nil || !nativePaintOnlyChild(cell.Children[1], d, "noFill") || requireEmptyNativeElement(cell.Children[1].Children[0]) != nil {
		return false
	}
	edges := cell.Children[0]
	if requireOnlyNativeAttrs(edges) != nil || len(edges.Children) != 6 {
		return false
	}
	for i, name := range []string{"left", "right", "top", "bottom", "insideH", "insideV"} {
		edge := edges.Children[i]
		if edge.Name != (xml.Name{Space: d.drawing, Local: name}) || requireOnlyNativeAttrs(edge) != nil || !nativePaintOnlyChild(edge, d, "ln") || requireOnlyNativeAttrs(edge.Children[0]) != nil || !nativePaintOnlyChild(edge.Children[0], d, "noFill") || requireEmptyNativeElement(edge.Children[0].Children[0]) != nil {
			return false
		}
	}
	return true
}
func nativePaintOnlyChild(n *nativeXMLNode, d nativeExtractDialect, name string) bool {
	return n != nil && onlyNativeXMLSpace(n.Text) && len(n.Children) == 1 && n.Children[0].Name == (xml.Name{Space: d.drawing, Local: name})
}

func inspectNativeTablePaint(frame *nativeXMLNode, c nativeTablePaintContext, d nativeExtractDialect) *NativePPTXTablePaint {
	if !c.qualified {
		return nil
	}
	graphic, _ := nativeSingleton(frame, d.drawing, "graphic", true)
	if graphic == nil {
		return nil
	}
	data, _ := nativeSingleton(graphic, d.drawing, "graphicData", true)
	if data == nil {
		return nil
	}
	table, _ := nativeSingleton(data, d.drawing, "tbl", true)
	if table == nil {
		return nil
	}
	props, _ := nativeSingleton(table, d.drawing, "tblPr", true)
	if props == nil {
		return nil
	}
	styleNode, _ := nativeSingleton(props, d.drawing, "tableStyleId", true)
	if styleNode == nil {
		return nil
	}
	id := styleNode.Text
	if !nativeTableNoStylePaint(c.styles[strings.ToUpper(id)], d) {
		return nil
	}
	rows := nativeChildren(table, d.drawing, "tr")
	if len(rows) != 1 {
		return nil
	}
	cells := nativeChildren(rows[0], d.drawing, "tc")
	if len(cells) != 1 {
		return nil
	}
	cell, _ := nativeSingleton(cells[0], d.drawing, "tcPr", true)
	if cell == nil || hasNativeSemanticAttrs(cell) {
		return nil
	}
	result := &NativePPTXTablePaint{Policy: "source-no-style-solid-border-v1", StyleID: id, Sources: c.sources, Fill: "none"}
	seen := map[string]bool{}
	var first *NativePPTXTablePaintBorder
	for _, edge := range cell.Children {
		name := edge.Name.Local
		if seen[name] {
			return nil
		}
		seen[name] = true
		if name == "noFill" {
			if requireEmptyNativeElement(edge) != nil {
				return nil
			}
			continue
		}
		if name == "lnTlToBr" || name == "lnBlToTr" {
			if !nativeTableInvisibleDiagonal(edge, d) {
				return nil
			}
			continue
		}
		if name != "lnL" && name != "lnR" && name != "lnT" && name != "lnB" {
			return nil
		}
		border, ok := nativeTablePresetPaintEdge(edge, c.theme, d)
		if !ok {
			return nil
		}
		if name == "lnL" {
			first = border
		} else if (first == nil) != (border == nil) || first != nil && (first.Color != border.Color || first.WidthEMU != border.WidthEMU || first.Preset != border.Preset) {
			return nil
		}
	}
	for _, name := range []string{"lnL", "lnR", "lnT", "lnB"} {
		if !seen[name] {
			return nil
		}
	}
	result.Border = first
	if first != nil && first.Preset != "" {
		result.Policy = "source-no-style-preset-border-v1"
	}
	return result
}
func nativeTableInvisibleDiagonal(edge *nativeXMLNode, d nativeExtractDialect) bool {
	if requireOnlyNativeAttrs(edge, xml.Name{Local: "w"}, xml.Name{Local: "cmpd"}) != nil || len(edge.Children) != 2 || edge.Children[0].Name != (xml.Name{Space: d.drawing, Local: "noFill"}) || requireEmptyNativeElement(edge.Children[0]) != nil || edge.Children[1].Name != (xml.Name{Space: d.drawing, Local: "prstDash"}) {
		return false
	}
	w, ok := exactNativeAttr(edge, "", "w")
	cmpd, present := exactNativeAttr(edge, "", "cmpd")
	dash := edge.Children[1]
	v, has := exactNativeAttr(dash, "", "val")
	return ok && w == "12700" && present && cmpd == "sng" && has && v == "solid" && requireOnlyNativeAttrs(dash, xml.Name{Local: "val"}) == nil && len(dash.Children) == 0
}
func nativeTablePresetPaintEdge(edge *nativeXMLNode, theme nativeResolvedTheme, d nativeExtractDialect) (*NativePPTXTablePaintBorder, bool) {
	if requireOnlyNativeAttrs(edge) == nil && nativePaintOnlyChild(edge, d, "noFill") && requireEmptyNativeElement(edge.Children[0]) == nil {
		return nil, true
	}
	if requireOnlyNativeAttrs(edge, xml.Name{Local: "w"}, xml.Name{Local: "cap"}, xml.Name{Local: "cmpd"}, xml.Name{Local: "algn"}) != nil {
		return nil, false
	}
	for name, want := range map[string]string{"cap": "flat", "cmpd": "sng", "algn": "ctr"} {
		v, ok := exactNativeAttr(edge, "", name)
		if !ok || v != want {
			return nil, false
		}
	}
	width, e := requiredCanonicalNativeTableInt(edge, "w", 1, 127000)
	if e != nil {
		return nil, false
	}
	if len(edge.Children) != 5 {
		return nil, false
	}
	for i, name := range []string{"solidFill", "prstDash", "round", "headEnd", "tailEnd"} {
		if edge.Children[i].Name != (xml.Name{Space: d.drawing, Local: name}) {
			return nil, false
		}
	}
	color, e := exactNativeSolidColor(edge.Children[0], d, theme)
	if e != nil {
		return nil, false
	}
	dash := edge.Children[1]
	v, ok := exactNativeAttr(dash, "", "val")
	if !ok || !nativeTablePaintPreset(v) || requireOnlyNativeAttrs(dash, xml.Name{Local: "val"}) != nil || len(dash.Children) != 0 || requireEmptyNativeElement(edge.Children[2]) != nil {
		return nil, false
	}
	for _, end := range edge.Children[3:] {
		if requireOnlyNativeAttrs(end, xml.Name{Local: "type"}, xml.Name{Local: "w"}, xml.Name{Local: "len"}) != nil || len(end.Children) != 0 {
			return nil, false
		}
		for name, want := range map[string]string{"type": "none", "w": "med", "len": "med"} {
			v, ok := exactNativeAttr(end, "", name)
			if !ok || v != want {
				return nil, false
			}
		}
	}
	if v == "solid" {
		v = ""
	}
	return &NativePPTXTablePaintBorder{Color: color, WidthEMU: width, Preset: v}, true
}

func nativeTablePaintWhitespace(node *nativeXMLNode) bool {
	if !onlyNativeXMLSpace(node.Text) {
		return false
	}
	for _, child := range node.Children {
		if !nativeTablePaintWhitespace(child) {
			return false
		}
	}
	return true
}

// Preset lengths are defined by ECMA-376 Part 1, 20.1.10.49.
// Phase/corner replay is an explicit preview policy, not Office raster evidence.
func nativeTablePaintPreset(value string) bool {
	switch value {
	case "solid", "dash", "dashDot", "dot", "lgDash", "lgDashDot", "lgDashDotDot", "sysDash", "sysDashDot", "sysDashDotDot", "sysDot":
		return true
	default:
		return false
	}
}
