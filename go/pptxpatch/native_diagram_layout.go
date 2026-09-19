package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strconv"
	"strings"
)

// SmartArt without a populated PowerPoint drawing fallback has nothing
// source-backed to paint verbatim. Under the opt-in approximate tier
// (AllowInheritedTextPreview) the frame is instead laid out from the four
// diagram parts by the bounded ECMA-376 §21.4 subset in
// native_diagram_layout_model.go and native_diagram_layout_hier.go. Every
// element is labeled with nativeDiagramLayoutPreviewCode so callers can tell
// computed layout from stored-fallback paint; the exact tier keeps refusing.
const (
	nativeDiagramLayoutPreviewCode     = "pptx.diagram-layout-approximate-v1"
	nativeDiagramLayoutPolicy          = "diagram-layout-approximate-v1"
	nativeDiagramLayoutTextOmittedCode = "pptx.diagram-layout-text-unavailable"
	nativeDiagramLayoutStyleCode       = "pptx.diagram-layout-style-unavailable"
	nativeDiagramLayoutPartsCode       = "pptx.diagram-layout-parts-unavailable"
	nativeDiagramDrawingEmptyCode      = "pptx.diagram-drawing-empty-unavailable"

	contentTypeDiagramLayout = "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"
	contentTypeDiagramStyle  = "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"
	contentTypeDiagramColors = "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"

	nativeDiagramLayoutGroupMessage = "SmartArt laid out from the diagram parts (" + nativeDiagramLayoutPolicy + "): composite, lin, snake, cycle, hierRoot, hierChild, sp, tx and conn (ECMA-376 §21.4). FIT uniform scale to the frame, centred. PAINT dgm:bg behind every shape; dgm:adj and dgm:shape@rot apply; zero alpha paints nothing, partial opaque; effects, 3D and image fills omitted. LIN children packed along linDir, cross-aligned by nodeVertAlign/nodeHorzAlign (ctr default), spacers counting along that axis only; an overrunning line shrinks by re-solving each child's constrLst over a proportional share (§21.4.2.24), so only that axis follows. RULES an INF extent is elastic: leftover shared pro rata inside each lte ceiling; an elastic node closes on its content. SNAKE lines break where the packed grid best matches the node's aspect, a DEVIATION from ST_BreakpointType. CYCLE shapes round the inscribed ellipse from stAng across spanAng, shrunk until neighbours clear; ctrShpMap fNode hubs the first child; rotPath alongPath turns them. HIERROOT assistants above regular children, bCtrCh default, alignOff a fraction of root w, hierAlign tL/tR laid out as hanging blocks below the root (§21.4.7.36 deviation), trunk gap sibSp/2. HIERCHILD subtrees packed by painted contour, each pair clearing sibSp. CONN joins the pair it separates, right-angle bends to the nearest end site, midpoint bend without bendDist. TX average-advance fitting, 0.5 em/glyph, 1.2 lines. CONSTRAINTS font sizes and margins are points, converted to EMU for extents; a bare constraint declares the default without erasing a value; a reference to nothing is inert; an unmodeled relationship or type refuses; maxDepth is relative to the context point, depth from the first selected, pos/revPos count sibling nodes; siblings run [parTrans, node, sibTrans], the last without one. Budgets 256 points, 32 deep, 2048 nodes, 65536 selections, 4096 constraints. Cached presOf/presParOf skipped; authored presStyleLbl and presLayoutVars honored; the frame is read-only"
	nativeDiagramLayoutChildMessage = "diagram element positioned by the approximate layout evaluation (" + nativeDiagramLayoutPolicy + "); target remains read-only"
)

func nativeDiagramRefusalCode(err error) string {
	if refusal, ok := err.(nativeGraphicFrameProjectionRefusal); ok {
		return refusal.code
	}
	return ""
}

type nativeDiagramLayoutParts struct {
	data, layout, style, colors *nativeXMLNode
	diagramNS                   string
}

// resolveNativeDiagramLayoutParts loads the four diagram parts named by
// dgm:relIds (§21.4.1) with exact relationship types and content types.
func (extractor *nativeExtractor) resolveNativeDiagramLayoutParts(graphic *nativeXMLNode, relationships []nativeExtractRelationship, dialect nativeExtractDialect) (nativeDiagramLayoutParts, error) {
	diagramNS := nativeDiagramURI(dialect)
	parts := nativeDiagramLayoutParts{diagramNS: diagramNS}
	data, err := nativeSingleton(graphic, dialect.drawing, "graphicData", true)
	if err != nil {
		return parts, err
	}
	relIds, err := nativeSingleton(data, diagramNS, "relIds", true)
	if err != nil {
		return parts, err
	}
	for _, item := range []struct {
		attr, relType, contentType, root string
		target                           **nativeXMLNode
	}{
		{"dm", "/diagramData", contentTypeDiagramData, "dataModel", &parts.data},
		{"lo", "/diagramLayout", contentTypeDiagramLayout, "layoutDef", &parts.layout},
		{"qs", "/diagramQuickStyle", contentTypeDiagramStyle, "styleDef", &parts.style},
		{"cs", "/diagramColors", contentTypeDiagramColors, "colorsDef", &parts.colors},
	} {
		id, ok := exactNativeAttr(relIds, dialect.rels, item.attr)
		if !ok || id == "" {
			return parts, refuseNativeDiagram(nativeDiagramLayoutPartsCode, "diagram layout requires the "+item.root+" relationship")
		}
		rel, err := nativeUniqueRelationship(relationships, id)
		if err != nil {
			return parts, err
		}
		if rel == nil || rel.Type != dialect.rels+item.relType || !rel.internal() || rel.Part == "" {
			return parts, refuseNativeDiagram(nativeDiagramLayoutPartsCode, "diagram "+item.root+" relationship is missing, external, or of the wrong type")
		}
		if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(rel.Part), item.contentType) {
			return parts, refuseNativeDiagram(nativeDiagramLayoutPartsCode, "diagram "+item.root+" part has an unexpected content type")
		}
		payload := extractor.pkg.parts[rel.Part]
		if len(payload) == 0 {
			return parts, refuseNativeDiagram(nativeDiagramLayoutPartsCode, "diagram "+item.root+" part is missing or empty")
		}
		root, err := parseNativeXML(payload, rel.Part)
		if err != nil || root.Name != (xml.Name{Space: diagramNS, Local: item.root}) {
			return parts, refuseNativeDiagram(nativeDiagramLayoutPartsCode, "diagram "+item.root+" part is not well-formed")
		}
		*item.target = root
	}
	return parts, nil
}

// nativeDiagramLayoutStyle is the resolved paint of one styleLbl instance.
type nativeDiagramLayoutStyle struct {
	fill      *string
	stroke    *NativeStroke
	fontStyle *nativeXMLNode
}

// nativeDiagramStyleLabels indexes dgm:styleLbl by name for the quick style
// (§21.4.5.10) and color transform (§21.4.4.10) parts.
func nativeDiagramStyleLabels(root *nativeXMLNode, diagramNS string) map[string]*nativeXMLNode {
	labels := map[string]*nativeXMLNode{}
	for _, label := range nativeChildren(root, diagramNS, "styleLbl") {
		if name, ok := exactNativeAttr(label, "", "name"); ok && labels[name] == nil {
			labels[name] = label
		}
	}
	return labels
}

// nativeDiagramColorAt applies ST_ClrAppMethod (§21.4.7.16) to pick the
// color for the index-th node of a style label. span needs interpolation and
// is only accepted for a single color.
func nativeDiagramColorAt(list *nativeXMLNode, index, count int, dialect nativeExtractDialect, theme nativeResolvedTheme) (string, bool, bool, error) {
	if list == nil || len(list.Children) == 0 {
		return "", false, false, nil
	}
	if requireOnlyNativeAttrs(list, xml.Name{Local: "meth"}, xml.Name{Local: "hueDir"}) != nil {
		return "", false, false, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram color list carries unknown attributes")
	}
	colors := list.Children
	method, _ := exactNativeAttr(list, "", "meth")
	position := 0
	switch method {
	case "", "repeat":
		position = index % len(colors)
	case "cycle":
		if len(colors) > 1 {
			period := 2 * (len(colors) - 1)
			position = index % period
			if position >= len(colors) {
				position = period - position
			}
		}
	case "span":
		if len(colors) > 1 {
			return "", false, false, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram spanning color interpolation is not approximated")
		}
	default:
		return "", false, false, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram color application method "+method+" is unknown")
	}
	entry, alpha, err := nativeDiagramOpaqueColor(colors[position], dialect)
	if err != nil {
		return "", false, false, err
	}
	if alpha == 0 {
		// A fully transparent list entry paints nothing. That is a decision
		// the list made, not a gap in it, so the style resolves with no
		// color for that matrix reference instead of refusing the frame.
		return "", false, true, nil
	}
	wrapper := &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "solidFill"}, Children: []*nativeXMLNode{entry}}
	color, err := exactNativeSolidColor(wrapper, dialect, theme)
	if err != nil {
		return "", false, true, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram color list requires exact sRGB or documented theme colors")
	}
	return color, true, true, nil
}

// nativeDiagramOpaqueColor splits a color list entry (§21.4.5) into the color
// without its a:alpha transforms and the resulting opacity in thousandths of
// a percent. The native contract carries an opaque sRGB triple, so a
// partially transparent list entry is painted at full opacity under this
// approximate tier, which the group diagnostic declares; a fully transparent
// one paints nothing, which IS representable and must not become a solid box.
// alphaMod and alphaOff are not folded in and still refuse downstream.
func nativeDiagramOpaqueColor(entry *nativeXMLNode, dialect nativeExtractDialect) (*nativeXMLNode, int64, error) {
	alpha := nativePositiveFixedPct
	kept := make([]*nativeXMLNode, 0, len(entry.Children))
	for _, child := range entry.Children {
		if child.Name != (xml.Name{Space: dialect.drawing, Local: "alpha"}) {
			kept = append(kept, child)
			continue
		}
		value, ok := exactNativeAttr(child, "", "val")
		if !ok {
			return nil, 0, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram color list alpha transform has no value")
		}
		percent, parseErr := strconv.ParseInt(value, 10, 64)
		if parseErr != nil || percent < 0 || percent > nativePositiveFixedPct {
			return nil, 0, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram color list alpha transform is out of range")
		}
		alpha = nativeRoundDiv(alpha*percent, nativePositiveFixedPct)
	}
	if len(kept) == len(entry.Children) {
		return entry, alpha, nil
	}
	return &nativeXMLNode{Name: entry.Name, Attrs: entry.Attrs, Children: kept}, alpha, nil
}

// resolveNativeDiagramLayoutStyle joins the quick style matrix references with
// the color transform lists and the theme format scheme (§21.4.4, §21.4.5,
// §20.1.4.1). Effects and 3D are omitted by policy.
func (extractor *nativeExtractor) resolveNativeDiagramLayoutStyle(label string, index, count int, quickStyles, colorLabels map[string]*nativeXMLNode, diagramNS string, dialect nativeExtractDialect) (nativeDiagramLayoutStyle, error) {
	result := nativeDiagramLayoutStyle{}
	refuse := func(message string) (nativeDiagramLayoutStyle, error) {
		return result, refuseNativeDiagram(nativeDiagramLayoutStyleCode, message)
	}
	if label == "" {
		return refuse("diagram shape has no style label")
	}
	quick, colors := quickStyles[label], colorLabels[label]
	if quick == nil || colors == nil {
		return refuse("diagram style label " + label + " is missing from the quick style or color parts")
	}
	style := nativeChild(quick, diagramNS, "style")
	if style == nil {
		return refuse("diagram quick style label has no style matrix reference")
	}
	refs := map[string]*nativeXMLNode{}
	for _, name := range []string{"lnRef", "fillRef", "effectRef", "fontRef"} {
		ref, err := nativeSingleton(style, dialect.drawing, name, true)
		if err != nil || requireOnlyNativeAttrs(ref, xml.Name{Local: "idx"}) != nil {
			return refuse("diagram quick style matrix references are outside the exact subset")
		}
		refs[name] = ref
	}
	themeRoot := extractor.slideDependencies.themeRoot
	if themeRoot == nil {
		return refuse("diagram layout requires the slide theme")
	}
	elements, err := nativeSingleton(themeRoot, dialect.drawing, "themeElements", true)
	if err != nil {
		return refuse("theme has no themeElements")
	}
	matrix, err := nativeSingleton(elements, dialect.drawing, "fmtScheme", true)
	if err != nil {
		return refuse("theme has no format scheme")
	}
	fillColor, hasFill, fillListed, err := nativeDiagramColorAt(nativeChild(colors, diagramNS, "fillClrLst"), index, count, dialect, extractor.theme)
	if err != nil {
		return result, err
	}
	lineColor, hasLine, lineListed, err := nativeDiagramColorAt(nativeChild(colors, diagramNS, "linClrLst"), index, count, dialect, extractor.theme)
	if err != nil {
		return result, err
	}
	textColor, hasText, _, err := nativeDiagramColorAt(nativeChild(colors, diagramNS, "txFillClrLst"), index, count, dialect, extractor.theme)
	if err != nil {
		return result, err
	}
	paint := &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "spPr"}}
	for _, item := range []struct {
		ref, list string
		color     string
		hasColor  bool
		listed    bool
	}{{"fillRef", "fillStyleLst", fillColor, hasFill, fillListed}, {"lnRef", "lnStyleLst", lineColor, hasLine, lineListed}} {
		value, _ := exactNativeAttr(refs[item.ref], "", "idx")
		if value == "0" {
			continue
		}
		selected, err := parseCanonicalNativeInt(value, 1, 3)
		if err != nil {
			return refuse("diagram style matrix index is outside the bounded first three entries")
		}
		list, err := nativeSingleton(matrix, dialect.drawing, item.list, true)
		if err != nil || selected > int64(len(list.Children)) {
			return refuse("theme style matrix entry is unavailable")
		}
		entry := list.Children[selected-1]
		if !item.hasColor {
			if item.listed {
				// The color transform listed a fully transparent entry for
				// this reference, so it paints nothing.
				continue
			}
			return refuse("diagram color transform defines no color for the " + item.ref + " matrix entry")
		}
		colored, err := nativeStylePlaceholderColor(entry, dialect, item.color)
		if err != nil {
			return refuse("theme style matrix placeholder color transforms are not approximated")
		}
		paint.Children = append(paint.Children, colored)
	}
	if len(paint.Children) != 0 {
		if nativeChild(paint, dialect.drawing, "ln") != nil {
			stroke, err := nativeDiagramDrawingLine(paint, dialect, extractor.theme)
			if err != nil {
				return result, err
			}
			result.stroke = stroke
		}
		fillChild := false
		for _, local := range []string{"solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill"} {
			fillChild = fillChild || nativeChild(paint, dialect.drawing, local) != nil
		}
		if fillChild {
			fill, err := nativeDiagramDrawingFill(paint, dialect, extractor.theme)
			if err != nil {
				return result, err
			}
			result.fill = fill
		}
	}
	fontIndex, _ := exactNativeAttr(refs["fontRef"], "", "idx")
	if fontIndex != "major" && fontIndex != "minor" {
		return refuse("diagram quick style font reference is not major or minor")
	}
	if !hasText {
		if len(refs["fontRef"].Children) == 0 {
			textColor, hasText = "", false
		} else {
			wrapper := &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "solidFill"}, Children: refs["fontRef"].Children}
			textColor, err = exactNativeSolidColor(wrapper, dialect, extractor.theme)
			if err != nil {
				return refuse("diagram quick style font color is not an exact solid color")
			}
			hasText = true
		}
	}
	if hasText {
		result.fontStyle = &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "style"}, Children: []*nativeXMLNode{{
			Name: xml.Name{Space: dialect.drawing, Local: "fontRef"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "idx"}, Value: fontIndex}},
			Children: []*nativeXMLNode{{Name: xml.Name{Space: dialect.drawing, Local: "srgbClr"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: textColor}}}},
		}}}
	}
	return result, nil
}

// nativeDiagramTextWords lists the words of every paragraph of a dgm:t body
// for the fitting model; a:br starts a new measured line.
func nativeDiagramTextWords(text *nativeXMLNode, dialect nativeExtractDialect) [][]string {
	paragraphs := [][]string{}
	for _, paragraph := range nativeChildren(text, dialect.drawing, "p") {
		var builder strings.Builder
		flush := func() {
			paragraphs = append(paragraphs, strings.Fields(builder.String()))
			builder.Reset()
		}
		for _, child := range paragraph.Children {
			switch child.Name.Local {
			case "r", "fld":
				if run := nativeChild(child, dialect.drawing, "t"); run != nil {
					builder.WriteString(run.Text)
				}
			case "br":
				flush()
			}
		}
		flush()
	}
	return paragraphs
}

func nativeDiagramTextHasWords(paragraphs [][]string) bool {
	for _, words := range paragraphs {
		if len(words) != 0 {
			return true
		}
	}
	return false
}

// buildNativeDiagramLayoutTextBody copies dgm:t (through the drawing-fallback
// copier, which already drops editor metadata) as an a:txBody carrying the tx
// algorithm result: margins as insets, middle anchor, centered paragraphs and
// the fitted size on every run that has none of its own.
func buildNativeDiagramLayoutTextBody(text *nativeXMLNode, dialect nativeExtractDialect, insets [4]int64, fontHundredths int64) *nativeXMLNode {
	a := func(local string) xml.Name { return xml.Name{Space: dialect.drawing, Local: local} }
	setAttr := func(node *nativeXMLNode, local, value string, override bool) {
		for index, attr := range node.Attrs {
			if attr.Name.Space == "" && attr.Name.Local == local {
				if override {
					node.Attrs[index].Value = value
				}
				return
			}
		}
		node.Attrs = append(node.Attrs, xml.Attr{Name: xml.Name{Local: local}, Value: value})
	}
	body := nativeDiagramDrawingTextBody(text, dialect)
	body.Name = a("txBody")
	bodyPr := nativeChild(body, dialect.drawing, "bodyPr")
	if bodyPr == nil {
		bodyPr = &nativeXMLNode{Name: a("bodyPr")}
		body.Children = append([]*nativeXMLNode{bodyPr}, body.Children...)
	}
	if nativeChild(body, dialect.drawing, "lstStyle") == nil {
		for index, child := range body.Children {
			if child == bodyPr {
				body.Children = append(body.Children[:index+1], append([]*nativeXMLNode{{Name: a("lstStyle")}}, body.Children[index+1:]...)...)
				break
			}
		}
	}
	for index, local := range []string{"lIns", "tIns", "rIns", "bIns"} {
		setAttr(bodyPr, local, strconv.FormatInt(insets[index], 10), true)
	}
	setAttr(bodyPr, "anchor", "ctr", false)
	setAttr(bodyPr, "wrap", "square", false)
	if nativeChild(bodyPr, dialect.drawing, "noAutofit") == nil && nativeChild(bodyPr, dialect.drawing, "normAutofit") == nil && nativeChild(bodyPr, dialect.drawing, "spAutoFit") == nil {
		bodyPr.Children = append(bodyPr.Children, &nativeXMLNode{Name: a("noAutofit")})
	}
	size := strconv.FormatInt(fontHundredths, 10)
	for _, paragraph := range nativeChildren(body, dialect.drawing, "p") {
		pPr := nativeChild(paragraph, dialect.drawing, "pPr")
		if pPr == nil {
			pPr = &nativeXMLNode{Name: a("pPr")}
			paragraph.Children = append([]*nativeXMLNode{pPr}, paragraph.Children...)
		}
		setAttr(pPr, "algn", "ctr", false)
		for _, run := range paragraph.Children {
			if run.Name != a("r") && run.Name != a("fld") {
				continue
			}
			rPr := nativeChild(run, dialect.drawing, "rPr")
			if rPr == nil {
				rPr = &nativeXMLNode{Name: a("rPr")}
				run.Children = append([]*nativeXMLNode{rPr}, run.Children...)
			}
			setAttr(rPr, "sz", size, false)
		}
	}
	return body
}

type nativeDiagramLayoutItem struct {
	node  *nativeDiagramPresNode
	order int
	index int
	style nativeDiagramLayoutStyle
	rect  nativeDiagramRect
}

// extractNativeDiagramLayoutGraphicFrame lays a SmartArt frame out from its
// diagram parts and projects the result as one preserve-only group.
func (extractor *nativeExtractor) extractNativeDiagramLayoutGraphicFrame(node, graphic *nativeXMLNode, slidePart, slideID string, relationships []nativeExtractRelationship, dialect nativeExtractDialect, objectID, name string, transform NativeTransform) (NativeElement, error) {
	parts, err := extractor.resolveNativeDiagramLayoutParts(graphic, relationships, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	model, err := parseNativeDiagramModel(parts.data, parts.diagramNS, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	evaluator, layoutRoot, err := newNativeDiagramLayoutEvaluator(model, parts.layout, parts.diagramNS)
	if err != nil {
		return NativeElement{}, err
	}
	presRoot, err := evaluator.evaluateLayoutNode(layoutRoot, model.doc, nil, 0)
	if err != nil {
		return NativeElement{}, err
	}
	frameW, frameH := float64(*transform.Cx), float64(*transform.Cy)
	presRoot.vals["w"], presRoot.vals["h"] = frameW, frameH
	if err := evaluator.evaluateConstraints(presRoot); err != nil {
		return NativeElement{}, err
	}
	if err := presRoot.layoutSubtree(frameW, frameH); err != nil {
		return NativeElement{}, err
	}
	fit, err := nativeDiagramFit(presRoot, frameW, frameH)
	if err != nil {
		return NativeElement{}, err
	}

	quickStyles := nativeDiagramStyleLabels(parts.style, parts.diagramNS)
	colorLabels := nativeDiagramStyleLabels(parts.colors, parts.diagramNS)
	items := []*nativeDiagramLayoutItem{}
	counts := map[string]int{}
	var equalize [][]*nativeDiagramPresNode
	var collect func(current *nativeDiagramPresNode)
	collect = func(current *nativeDiagramPresNode) {
		equalize = append(equalize, current.equalize...)
		if current.hasShape && current.shapeType != "" && !current.hideGeom && (current.laidOut || current.isConnector()) {
			items = append(items, &nativeDiagramLayoutItem{node: current, order: len(items), index: counts[current.styleLbl]})
			counts[current.styleLbl]++
		}
		for _, child := range current.children {
			collect(child)
		}
	}
	collect(presRoot)
	if len(items) == 0 {
		return NativeElement{}, refuseNativeDiagram(nativeDiagramLayoutGeometryCode, "diagram layout produced no visible shapes")
	}
	for _, item := range items {
		item.style, err = extractor.resolveNativeDiagramLayoutStyle(item.node.styleLbl, item.index, counts[item.node.styleLbl], quickStyles, colorLabels, parts.diagramNS, dialect)
		if err != nil {
			return NativeElement{}, err
		}
		item.rect = fit.rect(item.node.rect)
	}
	// tx algorithm: fit each text node, then honor primFontSz equalization.
	words := map[*nativeDiagramPresNode][][]string{}
	for _, item := range items {
		current := item.node
		if current.isConnector() || current.alg != "tx" || len(current.presOf) == 0 || current.presOf[0].text == nil {
			continue
		}
		paragraphs := nativeDiagramTextWords(current.presOf[0].text, dialect)
		if !nativeDiagramTextHasWords(paragraphs) {
			continue
		}
		words[current] = paragraphs
		maximum := current.value("primFontSz", nativeDiagramDefaultPrimFontSizePt)
		minimum := maximum
		for _, rule := range current.appliedRules {
			if rule.typ == "primFontSz" && rule.val > 0 {
				minimum = rule.val
			}
		}
		if maximum > nativeDiagramMaxFontSizePt || minimum > nativeDiagramMaxFontSizePt || maximum < 0 || minimum < 0 {
			return NativeElement{}, refuseNativeDiagram(nativeDiagramLayoutConstraintCode, "diagram primary font size constraint is outside the bounded range")
		}
		if minimum < nativeDiagramMinimumFontSizePt {
			minimum = nativeDiagramMinimumFontSizePt
		}
		width := item.rect.w - (current.value("lMarg", 0)+current.value("rMarg", 0))*nativeDiagramPointEMU
		height := item.rect.h - (current.value("tMarg", 0)+current.value("bMarg", 0))*nativeDiagramPointEMU
		current.fontSize = nativeDiagramFitFontSize(paragraphs, maximum, minimum, width, height)
	}
	for _, group := range equalize {
		smallest := int64(0)
		for _, member := range group {
			if member.fontSize > 0 && (smallest == 0 || member.fontSize < smallest) {
				smallest = member.fontSize
			}
		}
		for _, member := range group {
			if member.fontSize > 0 {
				member.fontSize = smallest
			}
		}
	}
	sort.SliceStable(items, func(a, b int) bool {
		if items[a].node.zOrderOff != items[b].node.zOrderOff {
			return items[a].node.zOrderOff < items[b].node.zOrderOff
		}
		return items[a].order < items[b].order
	})

	raw, err := rawNativeNode(extractor.pkg.parts[slidePart], node)
	if err != nil {
		return NativeElement{}, err
	}
	fingerprint := nativeSHA256(raw)
	children := make([]NativeElement, 0, len(items)+1)
	if model.background != nil {
		// The diagram background fills the whole frame behind every shape
		// (§21.4.3.2). It is source-backed paint, not layout.
		color, colorErr := exactNativeSolidColor(model.background, dialect, extractor.theme)
		if colorErr != nil {
			return NativeElement{}, refuseNativeDiagram(nativeDiagramLayoutStyleCode, "diagram background requires one exact sRGB or documented theme color")
		}
		backdrop := &nativeDiagramLayoutItem{node: &nativeDiagramPresNode{name: "background"}, order: -1}
		element, elementErr := extractor.nativeDiagramLayoutElementBase(NativeElementKindShape, slidePart, slideID, objectID, fingerprint, backdrop)
		if elementErr != nil {
			return NativeElement{}, elementErr
		}
		geometry, geometryErr := EvaluateNativePPTXPresetGeometry("rect", *transform.Cx, *transform.Cy, nil)
		if geometryErr != nil || geometry == nil {
			return NativeElement{}, refuseNativeDiagram(nativeDiagramLayoutGeometryCode, "diagram background rectangle is outside the evaluated profile")
		}
		element.Transform = NativeTransform{X: int64Pointer(0), Y: int64Pointer(0), Cx: int64Pointer(*transform.Cx), Cy: int64Pointer(*transform.Cy)}
		element.Geometry = geometry
		element.Fill = stringPointer(color)
		paragraphs := []NativeParagraph{}
		element.Paragraphs = &paragraphs
		children = append(children, element)
	}
	for _, item := range items {
		child, childErr := extractor.emitNativeDiagramLayoutItem(item, fit, words[item.node], slidePart, slideID, objectID, fingerprint, dialect)
		if childErr != nil {
			return NativeElement{}, childErr
		}
		if child != nil {
			children = append(children, *child)
		}
	}
	if len(children) == 0 {
		return NativeElement{}, refuseNativeDiagram(nativeDiagramLayoutGeometryCode, "diagram layout produced no paintable shapes")
	}
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	passthrough, err := extractor.issuePassthrough(slidePart, objectID, fingerprint, raw, nativeDiagramLayoutPreviewCode)
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return NativeElement{}, err
	}
	partName := slidePart
	element := NativeElement{
		Kind: NativeElementKindGroup, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform:      transform,
		ChildTransform: &NativeTransform{X: int64Pointer(0), Y: int64Pointer(0), Cx: int64Pointer(*transform.Cx), Cy: int64Pointer(*transform.Cy)},
		Children:       children,
		Passthrough:    []NativePassthroughRef{passthrough},
		Source:         &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusPreserveOnly, Diagnostics: []NativeDiagnostic{{
			Severity: NativeDiagnosticSeverityWarning, Code: nativeDiagramLayoutPreviewCode, Message: nativeDiagramLayoutGroupMessage,
			Scope: &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		}}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	return element, nil
}

func (extractor *nativeExtractor) nativeDiagramLayoutElementBase(kind NativeElementKind, slidePart, slideID, frameObjectID, frameFingerprint string, item *nativeDiagramLayoutItem) (NativeElement, error) {
	if extractor.elementsEmitted >= nativeMaxTotalElements || extractor.outputNodesEmitted >= nativeMaxNodes {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: element/output node budget exceeded")
	}
	extractor.elementsEmitted++
	extractor.outputNodesEmitted++
	objectID := frameObjectID + "/dgm/" + strconv.Itoa(item.order+1)
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	pointID := ""
	if item.node.point != nil {
		pointID = item.node.point.id
	}
	partName := slidePart
	element := NativeElement{
		Kind: kind, ID: elementID, Provenance: NativeProvenanceParsed,
		Passthrough: []NativePassthroughRef{}, Children: nil,
		Source: &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: nativeSHA256([]byte(frameFingerprint + "\x00" + item.node.name + "\x00" + pointID + "\x00" + objectID))},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusPreserveOnly, Diagnostics: []NativeDiagnostic{{
			Severity: NativeDiagnosticSeverityWarning, Code: nativeDiagramLayoutPreviewCode, Message: nativeDiagramLayoutChildMessage,
			Scope: &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		}}},
	}
	if item.node.name != "" {
		element.Name = stringPointer(item.node.name)
	}
	return element, nil
}

// emitNativeDiagramLayoutItem projects one laid-out shape or connector. A
// connector whose style paints no line, or a shape without extent, yields
// nothing rather than an invented element.
func (extractor *nativeExtractor) emitNativeDiagramLayoutItem(item *nativeDiagramLayoutItem, fit nativeDiagramFitTransform, words [][]string, slidePart, slideID, frameObjectID, frameFingerprint string, dialect nativeExtractDialect) (*NativeElement, error) {
	current := item.node
	if current.isConnector() {
		return extractor.emitNativeDiagramLayoutConnector(item, fit, slidePart, slideID, frameObjectID, frameFingerprint)
	}
	x, y := nativeDiagramRoundEMU(item.rect.x), nativeDiagramRoundEMU(item.rect.y)
	cx, cy := nativeDiagramRoundEMU(item.rect.w), nativeDiagramRoundEMU(item.rect.h)
	if cx <= 0 || cy <= 0 {
		return nil, nil
	}
	geometry, err := EvaluateNativePPTXPresetGeometry(current.shapeType, cx, cy, current.adjust)
	if err != nil || geometry == nil {
		return nil, refuseNativeDiagram(nativeDiagramLayoutGeometryCode, "diagram layout shape preset "+current.shapeType+" is outside the evaluated profile")
	}
	element, err := extractor.nativeDiagramLayoutElementBase(NativeElementKindShape, slidePart, slideID, frameObjectID, frameFingerprint, item)
	if err != nil {
		return nil, err
	}
	element.Transform = NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}
	if current.rotation60000 != 0 {
		element.Transform.RotationAngle = int64Pointer(current.rotation60000)
	}
	element.Geometry = geometry
	element.Fill = item.style.fill
	element.Stroke = item.style.stroke
	paragraphs := []NativeParagraph{}
	element.Paragraphs = &paragraphs
	if len(words) != 0 && current.fontSize > 0 && current.presOf[0].text != nil {
		insets := [4]int64{}
		for index, margin := range []string{"lMarg", "tMarg", "rMarg", "bMarg"} {
			insets[index] = nativeDiagramRoundEMU(current.value(margin, 0) * nativeDiagramPointEMU)
		}
		omitted := ""
		if item.style.fontStyle == nil {
			omitted = "diagram text has no resolvable font reference"
		} else if insets[0]+insets[2] >= cx || insets[1]+insets[3] >= cy {
			omitted = "diagram text margins leave no text area"
		} else {
			body := buildNativeDiagramLayoutTextBody(current.presOf[0].text, dialect, insets, current.fontSize*100)
			parsed, layout, reason, inherited, textErr := extractor.extractNativeDiagramDrawingText(body, item.style.fontStyle, element.Transform, dialect)
			if textErr != nil {
				return nil, textErr
			}
			if reason != "" {
				omitted = reason
			} else {
				paragraphs = parsed
				element.Paragraphs = &paragraphs
				element.TextBody = layout
				if inherited {
					nativeMarkInheritedTextPreview(&element)
				}
			}
		}
		if omitted != "" {
			element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
				Severity: NativeDiagnosticSeverityWarning, Code: nativeDiagramLayoutTextOmittedCode,
				Message: "diagram shape geometry retained; text omitted: " + omitted,
				Scope:   element.Compatibility.Diagnostics[0].Scope,
			})
		}
	}
	return &element, nil
}

// nativeDiagramConnectorEnds finds the begin and end shapes of a conn node.
// A connector joins the two presentation nodes its own position separates
// (§21.4.7.1 conn), and srcNode/dstNode name the layout node inside each of
// them to attach to:
//
//   - under a hierChild whose parent is a hierRoot, it runs from that
//     hierRoot's own root shape down to the hierarchy member that follows it,
//     which is the parent-to-child edge of an org chart;
//   - inside a linear node it is the sibTrans between two packed siblings, so
//     it runs from the sibling before it to the sibling after it. The linear
//     layouts give that sibTrans its own width, so the gap the connector is
//     drawn in is already reserved in the packed row.
func nativeDiagramConnectorEnds(conn *nativeDiagramPresNode) (*nativeDiagramPresNode, *nativeDiagramPresNode, error) {
	parent := conn.parent
	hierarchical := parent != nil && parent.alg == "hierChild" && parent.parent != nil && parent.parent.alg == "hierRoot"
	linear := parent != nil && parent.alg == "lin"
	if !hierarchical && !linear {
		return nil, nil, refuseNativeDiagram(nativeDiagramLayoutAlgorithmCode, "diagram connectors are only modeled inside linear nodes and between hierRoot parents and their hierChild members")
	}
	rootShape := func(container *nativeDiagramPresNode) *nativeDiagramPresNode {
		if container.alg != "hierRoot" {
			return container
		}
		for _, child := range container.children {
			if !child.isConnector() && child.alg != "hierChild" {
				return child
			}
		}
		return nil
	}
	named := func(container *nativeDiagramPresNode, name string) *nativeDiagramPresNode {
		if name == "" {
			return container
		}
		var found *nativeDiagramPresNode
		var walk func(current *nativeDiagramPresNode)
		walk = func(current *nativeDiagramPresNode) {
			if found == nil && current.name == name && current.hasShape {
				found = current
			}
			for _, child := range current.children {
				walk(child)
			}
		}
		walk(container)
		return found
	}
	position := -1
	for index, sibling := range parent.children {
		if sibling == conn {
			position = index
			break
		}
	}
	var previous, next *nativeDiagramPresNode
	for index := position - 1; index >= 0; index-- {
		if !parent.children[index].isConnector() {
			previous = parent.children[index]
			break
		}
	}
	if position >= 0 {
		for _, candidate := range parent.children[position+1:] {
			if !candidate.isConnector() {
				next = candidate
				break
			}
		}
	}
	source := previous
	if hierarchical {
		source = rootShape(parent.parent)
	}
	if source == nil || next == nil {
		return nil, nil, refuseNativeDiagram(nativeDiagramLayoutAlgorithmCode, "diagram connector has no begin or end shape")
	}
	begin := named(source, conn.param("srcNode", ""))
	if begin == nil {
		return nil, nil, refuseNativeDiagram(nativeDiagramLayoutAlgorithmCode, "diagram connector has no begin or end shape")
	}
	end := named(rootShape(next), conn.param("dstNode", ""))
	if end == nil || !begin.laidOut || !end.laidOut {
		return nil, nil, refuseNativeDiagram(nativeDiagramLayoutAlgorithmCode, "diagram connector end shape was not laid out")
	}
	return begin, end, nil
}

func (extractor *nativeExtractor) emitNativeDiagramLayoutConnector(item *nativeDiagramLayoutItem, fit nativeDiagramFitTransform, slidePart, slideID, frameObjectID, frameFingerprint string) (*NativeElement, error) {
	current := item.node
	if current.shapeType != "conn" {
		return nil, refuseNativeDiagram(nativeDiagramLayoutAlgorithmCode, "diagram conn algorithm requires the conn shape type")
	}
	if item.style.stroke == nil {
		return nil, nil
	}
	begin, end, err := nativeDiagramConnectorEnds(current)
	if err != nil {
		return nil, err
	}
	bendDist, hasBendDist := current.vals["bendDist"]
	points, err := routeNativeDiagramConnector(current, fit.rect(begin.rect), fit.rect(end.rect), bendDist*fit.scale, hasBendDist, current.value("begPad", 0)*fit.scale, current.value("endPad", 0)*fit.scale)
	if err != nil {
		return nil, err
	}
	minX, minY, maxX, maxY := points[0].x, points[0].y, points[0].x, points[0].y
	for _, point := range points[1:] {
		minX, minY = minFloat(minX, point.x), minFloat(minY, point.y)
		maxX, maxY = maxFloat(maxX, point.x), maxFloat(maxY, point.y)
	}
	x, y := nativeDiagramRoundEMU(minX), nativeDiagramRoundEMU(minY)
	cx, cy := nativeDiagramRoundEMU(maxX)-x, nativeDiagramRoundEMU(maxY)-y
	if cx < 1 {
		cx = 1
	}
	if cy < 1 {
		cy = 1
	}
	commands := make([]NativeGeometryCommand, 0, len(points))
	for index, point := range points {
		kind := "lineTo"
		if index == 0 {
			kind = "moveTo"
		}
		commands = append(commands, NativeGeometryCommand{Kind: kind, X: int64Pointer(nativeDiagramRoundEMU(point.x) - x), Y: int64Pointer(nativeDiagramRoundEMU(point.y) - y)})
	}
	element, err := extractor.nativeDiagramLayoutElementBase(NativeElementKindConnector, slidePart, slideID, frameObjectID, frameFingerprint, item)
	if err != nil {
		return nil, err
	}
	element.Transform = NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}
	if current.rotation60000 != 0 {
		element.Transform.RotationAngle = int64Pointer(current.rotation60000)
	}
	element.Geometry = &NativeEvaluatedGeometry{
		Profile:  "drawingml-paths-v1",
		TextRect: NativeGeometryTextRect{X: 0, Y: 0, CX: cx, CY: cy},
		Paths:    []NativeGeometryPath{{FillMode: "none", Stroke: true, Commands: commands}},
	}
	element.Stroke = item.style.stroke
	return &element, nil
}

func minFloat(a, b float64) float64 {
	if a < b {
		return a
	}
	return b
}

func maxFloat(a, b float64) float64 {
	if a > b {
		return a
	}
	return b
}
