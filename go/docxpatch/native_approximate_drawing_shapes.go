package docxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"
	"unicode/utf16"
)

// Read-only sidecar for the explicitly labeled approximate page preview.
//
// The strict extractor keeps refusing wps:wsp shapes (UNMODELED_RUN_CONTENT /
// UNMODELED_DRAWING and the picture refusals). This inspection joins those very
// diagnostics back to their source nodes and describes the bounded subset the
// approximate preview can paint: prstGeom rect or line with no adjust values,
// solid or theme-referenced fills, solid outlines and wp:inline / wp:anchor
// placement. Everything else stays omitted with a declared reason. It never
// changes source bytes, native extraction, editing authority or pagination.
const NativeApproximateDrawingShapesProtocol = "injoffice.docx.approximate-drawing-shapes"
const NativeApproximateDrawingShapePolicy = "docx.approximate-drawing-shape-preview-v1"
const nativeApproximateDrawingShapeLimit = 64
const nativeApproximateTextboxParagraphLimit = 256
const nativeApproximateTextboxTextLimit = 100000
const nativeMarkupCompatibilityNS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

type NativeApproximateShapeLineV1 struct {
	RGB      string `json:"rgb"`
	WidthEMU int64  `json:"width_emu"`
	Dash     string `json:"dash"`
}

type NativeApproximateTextboxV1 struct {
	LinkID             string                      `json:"link_id,omitempty"`
	LinkSeq            int64                       `json:"link_seq"`
	InsetsEMU          [4]int64                    `json:"insets_emu"` // left, top, right, bottom
	VerticalAnchor     string                      `json:"vertical_anchor"`
	Wrap               string                      `json:"wrap"`
	Paragraphs         []NativeParagraphV1         `json:"paragraphs"`
	ResolvedParagraphs []NativeResolvedParagraphV1 `json:"resolved_paragraphs"`
	ResolvedRuns       []NativeResolvedRunV1       `json:"resolved_runs"`
	OmittedRuns        int                         `json:"omitted_runs"`
	OmittedBlocks      int                         `json:"omitted_blocks"`
}

type NativeApproximateDrawingShapeV1 struct {
	ID              string                        `json:"id"`
	ParagraphID     string                        `json:"paragraph_id"`
	DiagnosticIDs   []string                      `json:"diagnostic_ids"`
	Anchor          NativeSourceAnchorV1          `json:"anchor"`
	RunAnchor       NativeSourceAnchorV1          `json:"run_anchor"`
	Status          string                        `json:"status"`
	Reason          string                        `json:"reason,omitempty"`
	Placement       string                        `json:"placement,omitempty"`
	Preset          string                        `json:"preset,omitempty"`
	WidthEMU        int64                         `json:"width_emu"`
	HeightEMU       int64                         `json:"height_emu"`
	RotationDegrees int64                         `json:"rotation_degrees"`
	FlipHorizontal  bool                          `json:"flip_horizontal"`
	FlipVertical    bool                          `json:"flip_vertical"`
	FillRGB         *string                       `json:"fill_rgb,omitempty"`
	Line            *NativeApproximateShapeLineV1 `json:"line,omitempty"`
	PageAnchor      *NativeTextboxPageAnchorV1    `json:"page_anchor,omitempty"`
	Wrap            string                        `json:"wrap,omitempty"`
	Textbox         *NativeApproximateTextboxV1   `json:"textbox,omitempty"`
	Notes           []string                      `json:"notes,omitempty"`
}

type NativeApproximateDrawingShapesV1 struct {
	Protocol      string                            `json:"protocol"`
	Version       int                               `json:"version"`
	Policy        string                            `json:"policy"`
	PackageSHA256 string                            `json:"package_sha256"`
	PartSHA256    string                            `json:"part_sha256"`
	Items         []NativeApproximateDrawingShapeV1 `json:"items"`
	OmittedCount  int                               `json:"omitted_count"`
}

type nativeApproximateShapeContext struct {
	resolver  *nativeLayoutResolver
	ns, wp, a string
	main      string
	raw       []byte
	theme     *nativeApproximateTheme
	textUnits int
	shapeSeen map[string]int
}

// InspectNativeApproximateDrawingShapesV1 returns nil when the body has no
// refused wps shapes. Callers must treat every item as approximate evidence.
func InspectNativeApproximateDrawingShapesV1(data []byte) (*NativeApproximateDrawingShapesV1, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxPackageBytes {
		return nil, fmt.Errorf("approximate drawing shapes package size must be 1..%d bytes", NativeDOCXMaxPackageBytes)
	}
	data = append([]byte(nil), data...)
	resolver, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	doc := resolver.doc
	ns, main := resolver.wordNS, resolver.mainPart
	wp, a := wordDrawingTransitional, drawingMLTransitional
	if ns == wordMLStrict {
		wp, a = wordDrawingStrict, drawingMLStrict
	}
	raw := resolver.pkg.files[main]
	context := &nativeApproximateShapeContext{resolver: resolver, ns: ns, wp: wp, a: a, main: main, raw: raw, theme: nativeApproximateThemeFromPackage(resolver.pkg, main, ns), shapeSeen: map[string]int{}}
	nodes := map[string]*nativeXMLNode{}
	var visit func(*nativeXMLNode)
	visit = func(n *nativeXMLNode) {
		nodes[n.Path] = n
		for _, c := range n.Children {
			visit(c)
		}
	}
	visit(resolver.mainRoot)
	diagnostics := map[string][]NativeUnsupportedCapabilityV1{}
	for _, d := range doc.Unsupported {
		if (d.Capability == "drawings" || d.Capability == "runs") && d.Anchor != nil && d.Anchor.PartName == main {
			diagnostics[d.ScopeID] = append(diagnostics[d.ScopeID], d)
		}
	}
	out := &NativeApproximateDrawingShapesV1{Protocol: NativeApproximateDrawingShapesProtocol, Version: 1, Policy: NativeApproximateDrawingShapePolicy, PackageSHA256: doc.Source.PackageSHA256, PartSHA256: nativeSHA(raw), Items: []NativeApproximateDrawingShapeV1{}}
	for _, block := range doc.Body.Blocks {
		p := block.Paragraph
		if p == nil || len(diagnostics[p.ID]) == 0 {
			continue
		}
		owner := nodes[p.Anchor.Path]
		if owner == nil || owner.Name != (xml.Name{Space: ns, Local: "p"}) {
			continue
		}
		for _, run := range owner.Children {
			if run.Name != (xml.Name{Space: ns, Local: "r"}) {
				continue
			}
			for _, child := range run.Children {
				drawing := context.drawingNode(child)
				if drawing == nil {
					continue
				}
				joined := []string{}
				for _, d := range diagnostics[p.ID] {
					if d.Anchor.Path != child.Path && !strings.HasPrefix(d.Anchor.Path, child.Path+"/") {
						continue
					}
					n := nodes[d.Anchor.Path]
					if n == nil || nativeSHA(raw[n.Start:n.End]) != d.Anchor.XMLSHA256 {
						continue
					}
					joined = append(joined, d.ID)
				}
				if len(joined) == 0 {
					// Modeled or otherwise owned content is never re-described here.
					continue
				}
				if len(out.Items) >= nativeApproximateDrawingShapeLimit {
					out.OmittedCount++
					continue
				}
				out.Items = append(out.Items, context.describe(p.ID, joined, run, drawing))
			}
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}

// drawingNode returns the w:drawing under a run child: either directly or the
// wps Choice of a markup-compatibility alternate. VML-only fallbacks are not
// read while a wps twin exists; a lone w:pict stays with its source refusal.
func (context *nativeApproximateShapeContext) drawingNode(child *nativeXMLNode) *nativeXMLNode {
	if child.Name == (xml.Name{Space: context.ns, Local: "drawing"}) {
		return child
	}
	if child.Name != (xml.Name{Space: nativeMarkupCompatibilityNS, Local: "AlternateContent"}) {
		return nil
	}
	for _, choice := range child.Children {
		if choice.Name != (xml.Name{Space: nativeMarkupCompatibilityNS, Local: "Choice"}) {
			continue
		}
		requires, _ := nativeUnqualifiedAttr(choice, "Requires")
		if !nativeApproximateRequiresWPS(requires) {
			continue
		}
		drawings := directNativeChildren(choice, context.ns, "drawing")
		if len(drawings) == 1 && len(choice.Children) == 1 {
			return drawings[0]
		}
		return nil
	}
	return nil
}

func nativeApproximateRequiresWPS(requires string) bool {
	for _, token := range strings.Fields(requires) {
		if token == "wps" {
			return true
		}
	}
	return false
}

func (context *nativeApproximateShapeContext) anchor(n *nativeXMLNode) NativeSourceAnchorV1 {
	return nativeTextboxSourceAnchor(n, context.main, context.raw)
}

func (context *nativeApproximateShapeContext) describe(paragraphID string, diagnosticIDs []string, run, drawing *nativeXMLNode) NativeApproximateDrawingShapeV1 {
	digest := strings.TrimPrefix(nativeSHA(context.raw[drawing.Start:drawing.End]), "sha256:")[:16]
	context.shapeSeen[digest]++
	item := NativeApproximateDrawingShapeV1{ID: fmt.Sprintf("approximate-drawing-shape:%s:%d", digest, context.shapeSeen[digest]), ParagraphID: paragraphID, DiagnosticIDs: diagnosticIDs, Anchor: context.anchor(drawing), RunAnchor: context.anchor(run), Status: "omitted"}
	omit := func(reason string) NativeApproximateDrawingShapeV1 {
		item.Reason = reason
		return item
	}
	wp, a := context.wp, context.a
	var container *nativeXMLNode
	for _, c := range drawing.Children {
		if c.Name.Space == wp && (c.Name.Local == "inline" || c.Name.Local == "anchor") {
			if container != nil {
				return omit("ambiguous-drawing-container")
			}
			container = c
		}
	}
	if container == nil {
		return omit("missing-drawing-container")
	}
	extent := firstDirectNativeChild(container, wp, "extent")
	if extent == nil {
		return omit("missing-extent")
	}
	width, okW := nativePositiveInt64Attr(extent, "", "cx")
	height, okH := nativePositiveInt64Attr(extent, "", "cy")
	if !okW || !okH || width > 127000000 || height > 127000000 {
		return omit("invalid-extent")
	}
	item.WidthEMU, item.HeightEMU = width, height
	graphic := firstDirectNativeChild(container, a, "graphic")
	if graphic == nil {
		return omit("missing-graphic")
	}
	graphicData := firstDirectNativeChild(graphic, a, "graphicData")
	if graphicData == nil {
		return omit("missing-graphic-data")
	}
	uri, _ := nativeUnqualifiedAttr(graphicData, "uri")
	if uri != nativeTextboxWPS {
		local := "unknown"
		if len(graphicData.Children) == 1 {
			local = graphicData.Children[0].Name.Local
		}
		return omit("unsupported-graphic:" + local)
	}
	shapes := directNativeChildren(graphicData, nativeTextboxWPS, "wsp")
	if len(shapes) != 1 || len(graphicData.Children) != 1 {
		return omit("unsupported-graphic:group-or-multiple")
	}
	shape := shapes[0]
	spPr := firstDirectNativeChild(shape, nativeTextboxWPS, "spPr")
	if spPr == nil {
		return omit("missing-shape-properties")
	}
	style := firstDirectNativeChild(shape, nativeTextboxWPS, "style")
	geometry := firstDirectNativeChild(spPr, a, "prstGeom")
	if geometry == nil {
		if firstDirectNativeChild(spPr, a, "custGeom") != nil {
			return omit("custom-geometry")
		}
		return omit("missing-geometry")
	}
	preset, _ := nativeUnqualifiedAttr(geometry, "prst")
	if preset != "rect" && preset != "line" {
		return omit("unsupported-preset:" + preset)
	}
	for _, c := range geometry.Children {
		if c.Name == (xml.Name{Space: a, Local: "avLst"}) && len(c.Children) > 0 {
			return omit("adjust-values")
		}
	}
	item.Preset = preset
	if xfrm := firstDirectNativeChild(spPr, a, "xfrm"); xfrm != nil {
		if rot, ok := nativeUnqualifiedAttr(xfrm, "rot"); ok {
			value, err := strconv.ParseInt(rot, 10, 64)
			if err != nil || value < -21600000*100 || value > 21600000*100 {
				return omit("invalid-rotation")
			}
			degrees := (value / 60000) % 360
			if degrees < 0 {
				degrees += 360
			}
			item.RotationDegrees = degrees
		}
		item.FlipHorizontal = nativeApproximateFlag(xfrm, "flipH")
		item.FlipVertical = nativeApproximateFlag(xfrm, "flipV")
	}
	if preset == "rect" && item.RotationDegrees%90 != 0 {
		return omit("rotation-unsupported")
	}
	if preset == "rect" && (item.RotationDegrees == 90 || item.RotationDegrees == 270) {
		item.Notes = append(item.Notes, "quarter-turn rectangle painted as its rotated bounding box")
	}
	fill, fillNotes, fillOK := context.shapeFill(spPr, style)
	if !fillOK {
		return omit("unsupported-fill")
	}
	item.FillRGB = fill
	item.Notes = append(item.Notes, fillNotes...)
	line, lineNotes, lineOK := context.shapeLine(spPr, style)
	if !lineOK {
		return omit("unsupported-outline")
	}
	item.Line = line
	item.Notes = append(item.Notes, lineNotes...)
	if item.Line != nil && (item.Line.WidthEMU >= width && item.Line.WidthEMU >= height) {
		return omit("outline-exceeds-shape")
	}
	if container.Name.Local == "inline" {
		item.Placement = "inline"
	} else {
		item.Placement = "anchored"
		pageAnchor, wrap, reason := context.pageAnchor(container)
		if reason != "" {
			return omit(reason)
		}
		item.PageAnchor, item.Wrap = pageAnchor, wrap
		if wrap != "none" {
			item.Notes = append(item.Notes, "body text wrapping around the shape is not applied")
		}
	}
	textbox, reason := context.textbox(shape, item.ID)
	if reason != "" {
		item.Notes = append(item.Notes, "textbox content omitted: "+reason)
	} else {
		item.Textbox = textbox
	}
	item.Status = "supported"
	return item
}

func nativeApproximateFlag(n *nativeXMLNode, local string) bool {
	value, ok := nativeUnqualifiedAttr(n, local)
	return ok && (value == "1" || value == "true")
}

// pageAnchor reads wp:anchor placement leniently: authored offsets or alignment
// per axis plus stacking. Wrap geometry is recorded, not applied.
func (context *nativeApproximateShapeContext) pageAnchor(container *nativeXMLNode) (*NativeTextboxPageAnchorV1, string, string) {
	wp := context.wp
	if nativeApproximateFlag(container, "simplePos") {
		return nil, "", "simple-position-unsupported"
	}
	positionH := firstDirectNativeChild(container, wp, "positionH")
	positionV := firstDirectNativeChild(container, wp, "positionV")
	if positionH == nil || positionV == nil {
		return nil, "", "missing-position"
	}
	horizontal := map[string]bool{"page": true, "margin": true, "column": true, "character": true, "leftMargin": true, "rightMargin": true, "insideMargin": true, "outsideMargin": true}
	vertical := map[string]bool{"page": true, "margin": true, "paragraph": true, "line": true, "topMargin": true, "bottomMargin": true, "insideMargin": true, "outsideMargin": true}
	x, hRelative, hAlign, ok := nativeApproximatePosition(positionH, wp, map[string]bool{"left": true, "center": true, "right": true, "inside": true, "outside": true})
	if !ok || !horizontal[hRelative] || (hAlign != "" && hRelative == "character") {
		return nil, "", "unsupported-horizontal-position"
	}
	y, vRelative, vAlign, ok := nativeApproximatePosition(positionV, wp, map[string]bool{"top": true, "center": true, "bottom": true, "inside": true, "outside": true})
	if !ok || !vertical[vRelative] || (vAlign != "" && (vRelative == "paragraph" || vRelative == "line")) {
		return nil, "", "unsupported-vertical-position"
	}
	relativeHeight, _ := nativeNonnegativeInt64Attr(container, "", "relativeHeight")
	stacking := uint32(math.MaxUint32)
	if relativeHeight >= 0 && relativeHeight <= math.MaxUint32 {
		stacking = uint32(relativeHeight)
	}
	anchor := &NativeTextboxPageAnchorV1{
		Stacking:         &NativeTextboxStackingV1{BehindDoc: nativeApproximateFlag(container, "behindDoc"), RelativeHeight: stacking},
		Policy:           "relative-position-no-wrap-v2",
		SourceAnchor:     context.anchor(container),
		HorizontalAnchor: context.anchor(positionH), VerticalAnchor: context.anchor(positionV),
		XEMU: x, YEMU: y, HorizontalRelative: hRelative, VerticalRelative: vRelative, HorizontalAlign: hAlign, VerticalAlign: vAlign,
	}
	wraps := map[string]string{"wrapNone": "none", "wrapSquare": "square", "wrapTight": "tight", "wrapThrough": "through", "wrapTopAndBottom": "top-and-bottom"}
	wrap := "none"
	for _, child := range container.Children {
		if value, ok := wraps[child.Name.Local]; ok && child.Name.Space == wp {
			wrap = value
		}
	}
	return anchor, wrap, ""
}

func nativeApproximatePosition(node *nativeXMLNode, wp string, aligns map[string]bool) (int64, string, string, bool) {
	relativeFrom, _ := nativeUnqualifiedAttr(node, "relativeFrom")
	if len(node.Children) != 1 {
		return 0, "", "", false
	}
	child := node.Children[0]
	if child.Name.Space != wp {
		return 0, "", "", false
	}
	switch child.Name.Local {
	case "posOffset":
		value, err := strconv.ParseInt(strings.TrimSpace(child.Text), 10, 64)
		if err != nil || value < -127000000 || value > 127000000 {
			return 0, "", "", false
		}
		return value, relativeFrom, "", true
	case "align":
		value := strings.TrimSpace(child.Text)
		if !aligns[value] {
			return 0, "", "", false
		}
		return 0, relativeFrom, value, true
	}
	return 0, "", "", false
}

// shapeFill returns nil for no fill. Gradient, pattern and picture fills are
// not approximated; they omit the fill and record the omission.
func (context *nativeApproximateShapeContext) shapeFill(spPr, style *nativeXMLNode) (*string, []string, bool) {
	a := context.a
	notes := []string{}
	for _, child := range spPr.Children {
		if child.Name.Space != a {
			continue
		}
		switch child.Name.Local {
		case "noFill":
			return nil, notes, true
		case "solidFill":
			rgb, colorNotes, ok := context.solidColor(child, "")
			if !ok {
				return nil, nil, false
			}
			return nativeString(rgb), append(notes, colorNotes...), true
		case "gradFill", "pattFill", "blipFill", "grpFill":
			return nil, append(notes, child.Name.Local+" is not approximated; fill omitted"), true
		}
	}
	if style == nil {
		return nil, notes, true
	}
	reference := firstDirectNativeChild(style, a, "fillRef")
	if reference == nil {
		return nil, notes, true
	}
	index, _ := nativeNonnegativeInt64Attr(reference, "", "idx")
	if index == 0 {
		return nil, notes, true
	}
	phClr, colorNotes, ok := context.referenceColor(reference)
	if !ok {
		return nil, nil, false
	}
	notes = append(notes, colorNotes...)
	if context.theme == nil || index < 1 || index > int64(len(context.theme.fills)) {
		return nil, append(notes, "theme fill style unavailable; fill omitted"), true
	}
	fillStyle := context.theme.fills[int(index)-1]
	if fillStyle.Name != (xml.Name{Space: context.theme.ns, Local: "solidFill"}) {
		return nil, append(notes, "theme "+fillStyle.Name.Local+" is not approximated; fill omitted"), true
	}
	rgb, themeNotes, ok := context.solidColor(fillStyle, phClr)
	if !ok {
		return nil, nil, false
	}
	return nativeString(rgb), append(append(notes, "fill resolved from theme fill style"), themeNotes...), true
}

func (context *nativeApproximateShapeContext) shapeLine(spPr, style *nativeXMLNode) (*NativeApproximateShapeLineV1, []string, bool) {
	a := context.a
	notes := []string{}
	ln := firstDirectNativeChild(spPr, a, "ln")
	var themeLine *nativeXMLNode
	phClr := ""
	if style != nil {
		if reference := firstDirectNativeChild(style, a, "lnRef"); reference != nil {
			index, _ := nativeNonnegativeInt64Attr(reference, "", "idx")
			if index >= 1 && context.theme != nil && index <= int64(len(context.theme.lines)) {
				color, colorNotes, ok := context.referenceColor(reference)
				if !ok {
					return nil, nil, false
				}
				phClr = color
				notes = append(notes, colorNotes...)
				themeLine = context.theme.lines[int(index)-1]
			}
		}
	}
	if ln == nil && themeLine == nil {
		return nil, notes, true
	}
	width := int64(-1)
	rgb := ""
	dash := ""
	hasFill := false
	read := func(node *nativeXMLNode, placeholder string) bool {
		if w, ok := nativeNonnegativeInt64Attr(node, "", "w"); ok && width < 0 {
			width = w
		}
		for _, child := range node.Children {
			if child.Name.Space != a {
				continue
			}
			switch child.Name.Local {
			case "noFill":
				if !hasFill {
					hasFill = true
					rgb = ""
				}
			case "solidFill":
				if !hasFill {
					color, colorNotes, ok := context.solidColor(child, placeholder)
					if !ok {
						return false
					}
					notes = append(notes, colorNotes...)
					hasFill = true
					rgb = color
				}
			case "gradFill", "pattFill":
				if !hasFill {
					hasFill = true
					rgb = ""
					notes = append(notes, "outline "+child.Name.Local+" is not approximated; outline omitted")
				}
			case "prstDash":
				if dash == "" {
					dash, _ = nativeUnqualifiedAttr(child, "val")
				}
			}
		}
		return true
	}
	if ln != nil && !read(ln, "") {
		return nil, nil, false
	}
	if themeLine != nil && !read(themeLine, phClr) {
		return nil, nil, false
	}
	if !hasFill {
		if ln != nil && themeLine == nil {
			// An a:ln without explicit fill inherits nothing modeled here.
			return nil, append(notes, "outline without explicit fill omitted"), true
		}
		return nil, notes, true
	}
	if rgb == "" {
		return nil, notes, true
	}
	if width < 0 {
		width = 9525
		notes = append(notes, "outline width defaulted to 9525 EMU (0.75pt)")
	}
	if width == 0 {
		width = 9525
		notes = append(notes, "hairline outline painted at 0.75pt")
	}
	if width > 12700000 {
		return nil, nil, false
	}
	if dash == "" {
		dash = "solid"
	}
	if themeLine != nil {
		notes = append(notes, "outline resolved from theme line style")
	}
	return &NativeApproximateShapeLineV1{RGB: rgb, WidthEMU: width, Dash: dash}, notes, true
}

func (context *nativeApproximateShapeContext) referenceColor(reference *nativeXMLNode) (string, []string, bool) {
	for _, child := range reference.Children {
		if child.Name.Space == context.a {
			return context.color(child, "")
		}
	}
	return "", nil, false
}

// solidColor reads the single color child of an a:solidFill.
func (context *nativeApproximateShapeContext) solidColor(fill *nativeXMLNode, phClr string) (string, []string, bool) {
	for _, child := range fill.Children {
		if child.Name.Space == context.a {
			return context.color(child, phClr)
		}
	}
	return "", nil, false
}

// color resolves srgbClr, sysClr (lastClr) and schemeClr through the package
// theme. Shade, tint, lumMod and lumOff transforms are approximated in sRGB.
func (context *nativeApproximateShapeContext) color(node *nativeXMLNode, phClr string) (string, []string, bool) {
	a := context.a
	notes := []string{}
	var rgb string
	switch node.Name.Local {
	case "srgbClr":
		value, _ := nativeUnqualifiedAttr(node, "val")
		exact, ok := nativeExactRGB(value)
		if !ok {
			return "", nil, false
		}
		rgb = exact
	case "sysClr":
		value, _ := nativeUnqualifiedAttr(node, "lastClr")
		exact, ok := nativeExactRGB(value)
		if !ok {
			return "", nil, false
		}
		rgb = exact
	case "schemeClr":
		value, _ := nativeUnqualifiedAttr(node, "val")
		if value == "phClr" {
			if phClr == "" {
				return "", nil, false
			}
			rgb = phClr
		} else {
			slot := map[string]string{"tx1": "dk1", "bg1": "lt1", "tx2": "dk2", "bg2": "lt2"}[value]
			if slot == "" {
				slot = value
			}
			if context.theme == nil || context.theme.colors[slot] == "" {
				return "", nil, false
			}
			rgb = context.theme.colors[slot]
			notes = append(notes, "scheme color "+value+" resolved from the package theme")
		}
	default:
		return "", nil, false
	}
	for _, transform := range node.Children {
		if transform.Name.Space != a {
			continue
		}
		value, ok := nativeNonnegativeInt64Attr(transform, "", "val")
		if !ok || value > 10000000 {
			continue
		}
		switch transform.Name.Local {
		case "shade":
			rgb = nativeApproximateScaleRGB(rgb, float64(value)/100000, 0)
			notes = append(notes, "shade transform approximated in sRGB")
		case "tint":
			rgb = nativeApproximateScaleRGB(rgb, float64(value)/100000, 1)
			notes = append(notes, "tint transform approximated in sRGB")
		case "lumMod":
			rgb = nativeApproximateLuminance(rgb, float64(value)/100000, 0)
			notes = append(notes, "luminance transform approximated in sRGB")
		case "lumOff":
			rgb = nativeApproximateLuminance(rgb, 1, float64(value)/100000)
			notes = append(notes, "luminance transform approximated in sRGB")
		case "alpha":
			notes = append(notes, "alpha transform ignored; painted opaque")
		}
	}
	return rgb, notes, true
}

func nativeApproximateRGBComponents(rgb string) [3]float64 {
	var out [3]float64
	for i := 0; i < 3; i++ {
		value, _ := strconv.ParseUint(rgb[i*2:i*2+2], 16, 8)
		out[i] = float64(value) / 255
	}
	return out
}

func nativeApproximateRGBString(components [3]float64) string {
	out := ""
	for _, value := range components {
		if value < 0 {
			value = 0
		}
		if value > 1 {
			value = 1
		}
		out += fmt.Sprintf("%02X", int(math.Round(value*255)))
	}
	return out
}

// toward 0 scales toward black (shade); toward 1 scales toward white (tint).
func nativeApproximateScaleRGB(rgb string, factor float64, toward float64) string {
	components := nativeApproximateRGBComponents(rgb)
	for i := range components {
		components[i] = toward + (components[i]-toward)*factor
	}
	return nativeApproximateRGBString(components)
}

func nativeApproximateLuminance(rgb string, mod, off float64) string {
	c := nativeApproximateRGBComponents(rgb)
	maxC := math.Max(c[0], math.Max(c[1], c[2]))
	minC := math.Min(c[0], math.Min(c[1], c[2]))
	l := (maxC + minC) / 2
	var h, s float64
	if maxC != minC {
		d := maxC - minC
		if l > 0.5 {
			s = d / (2 - maxC - minC)
		} else {
			s = d / (maxC + minC)
		}
		switch maxC {
		case c[0]:
			h = (c[1] - c[2]) / d
			if c[1] < c[2] {
				h += 6
			}
		case c[1]:
			h = (c[2]-c[0])/d + 2
		default:
			h = (c[0]-c[1])/d + 4
		}
		h /= 6
	}
	l = l*mod + off
	if l < 0 {
		l = 0
	}
	if l > 1 {
		l = 1
	}
	if s == 0 {
		return nativeApproximateRGBString([3]float64{l, l, l})
	}
	var q float64
	if l < 0.5 {
		q = l * (1 + s)
	} else {
		q = l + s - l*s
	}
	p := 2*l - q
	hue := func(t float64) float64 {
		if t < 0 {
			t++
		}
		if t > 1 {
			t--
		}
		switch {
		case t < 1.0/6:
			return p + (q-p)*6*t
		case t < 0.5:
			return q
		case t < 2.0/3:
			return p + (q-p)*(2.0/3-t)*6
		}
		return p
	}
	return nativeApproximateRGBString([3]float64{hue(h + 1.0/3), hue(h), hue(h - 1.0/3)})
}

type nativeApproximateTheme struct {
	ns     string
	colors map[string]string
	fills  []*nativeXMLNode
	lines  []*nativeXMLNode
}

func nativeApproximateThemeFromPackage(pkg *nativePackage, mainPart, wordNS string) *nativeApproximateTheme {
	relBase, drawingNS := relBaseTransitional, drawingMLTransitional
	if wordNS == wordMLStrict {
		relBase, drawingNS = relBaseStrict, drawingMLStrict
	}
	partName, count := "", 0
	for _, rel := range pkg.rels[mainPart] {
		if rel.Type != relBase+"theme" {
			continue
		}
		if rel.External || rel.PartName == "" {
			return nil
		}
		count++
		partName = rel.PartName
	}
	if count != 1 || !nativeASCIIEqual(pkg.contentTypes[partName], "application/vnd.openxmlformats-officedocument.theme+xml") {
		return nil
	}
	root, err := parseNativeXML(partName, pkg.files[partName])
	if err != nil || root.Name != (xml.Name{Space: drawingNS, Local: "theme"}) {
		return nil
	}
	theme := &nativeApproximateTheme{ns: drawingNS, colors: nativeParseThemeSrgbColors(root, drawingNS)}
	elements := firstDirectNativeChild(root, drawingNS, "themeElements")
	if elements == nil {
		return theme
	}
	if scheme := firstDirectNativeChild(elements, drawingNS, "fmtScheme"); scheme != nil {
		if fills := firstDirectNativeChild(scheme, drawingNS, "fillStyleLst"); fills != nil {
			for _, child := range fills.Children {
				if child.Name.Space == drawingNS && len(theme.fills) < 16 {
					theme.fills = append(theme.fills, child)
				}
			}
		}
		if lines := firstDirectNativeChild(scheme, drawingNS, "lnStyleLst"); lines != nil {
			for _, child := range lines.Children {
				if child.Name == (xml.Name{Space: drawingNS, Local: "ln"}) && len(theme.lines) < 16 {
					theme.lines = append(theme.lines, child)
				}
			}
		}
	}
	return theme
}

// textbox reads wps:txbx / wps:linkedTxbx content through the ordinary
// paragraph extractor and style resolver so the approximate preview shapes the
// same run model as body text. Nested tables, drawings and references inside
// the box are counted as omitted rather than flattened.
func (context *nativeApproximateShapeContext) textbox(shape *nativeXMLNode, shapeID string) (*NativeApproximateTextboxV1, string) {
	ns, a := context.ns, context.a
	box := firstDirectNativeChild(shape, nativeTextboxWPS, "txbx")
	linked := firstDirectNativeChild(shape, nativeTextboxWPS, "linkedTxbx")
	if box == nil && linked == nil {
		return nil, ""
	}
	textbox := &NativeApproximateTextboxV1{InsetsEMU: [4]int64{91440, 45720, 91440, 45720}, VerticalAnchor: "t", Wrap: "square", Paragraphs: []NativeParagraphV1{}, ResolvedParagraphs: []NativeResolvedParagraphV1{}, ResolvedRuns: []NativeResolvedRunV1{}}
	if body := firstDirectNativeChild(shape, nativeTextboxWPS, "bodyPr"); body != nil {
		for i, key := range []string{"lIns", "tIns", "rIns", "bIns"} {
			if value, ok := nativeNonnegativeInt64Attr(body, "", key); ok && value <= 127000000 {
				textbox.InsetsEMU[i] = value
			}
		}
		if value, ok := nativeUnqualifiedAttr(body, "anchor"); ok && (value == "t" || value == "ctr" || value == "b") {
			textbox.VerticalAnchor = value
		}
		if value, ok := nativeUnqualifiedAttr(body, "wrap"); ok && value == "none" {
			textbox.Wrap = "none"
		}
		if value, ok := nativeUnqualifiedAttr(body, "vert"); ok && value != "horz" {
			return nil, "vertical-text"
		}
		if firstDirectNativeChild(body, a, "spAutoFit") != nil {
			// Autofit changes the shape extent in Word; the authored extent is kept.
		}
	}
	if linked != nil {
		if box != nil {
			return nil, "ambiguous-textbox"
		}
		id, _ := nativeUnqualifiedAttr(linked, "id")
		seq, ok := nativeNonnegativeInt64Attr(linked, "", "seq")
		if id == "" || !ok || seq == 0 || len(id) > 64 {
			return nil, "invalid-linked-textbox"
		}
		textbox.LinkID, textbox.LinkSeq = id, seq
		return textbox, ""
	}
	if id, ok := nativeUnqualifiedAttr(box, "id"); ok && id != "" && len(id) <= 64 {
		textbox.LinkID = id
	}
	contents := directNativeChildren(box, ns, "txbxContent")
	if len(contents) != 1 {
		return nil, "ambiguous-textbox-content"
	}
	content := contents[0]
	if len(content.Children) > nativeApproximateTextboxParagraphLimit {
		return nil, "paragraph-limit"
	}
	extractor := context.contentExtractor()
	localResolver := *context.resolver
	localResolver.diagnostics = append([]NativeResolutionDiagnosticV1{}, context.resolver.diagnostics...)
	localResolver.diagnosticSet = map[string]bool{}
	for key, value := range context.resolver.diagnosticSet {
		localResolver.diagnosticSet[key] = value
	}
	result := &NativeResolvedLayoutInputV1{}
	numberingState := newNativeNumberingState()
	safeID := strings.NewReplacer("/", "-").Replace(shapeID)
	for index, child := range content.Children {
		if child.Name != (xml.Name{Space: ns, Local: "p"}) {
			textbox.OmittedBlocks++
			continue
		}
		refusalsBefore := len(extractor.unsupported)
		paragraph, err := extractor.extractParagraph(context.main, child)
		if err != nil {
			return nil, "unsupported-textbox-content"
		}
		// Content the ordinary extractor refuses (references, fields, nested
		// markup) is dropped from the preview and counted, never merged.
		textbox.OmittedRuns += len(extractor.unsupported) - refusalsBefore
		paragraph.ID = fmt.Sprintf("%s:p%d", safeID, index)
		paragraph.EditPolicy = nativeReadOnlyPolicy("APPROXIMATE_TEXTBOX_PREVIEW", "Textbox paragraphs are read-only approximate preview evidence")
		runs := []NativeRunV1{}
		for runIndex := range paragraph.Runs {
			run := paragraph.Runs[runIndex]
			if run.Kind != "text" && run.Kind != "control" {
				textbox.OmittedRuns++
				continue
			}
			if run.Kind == "control" && run.Control != "tab" && run.Control != "line-break" {
				textbox.OmittedRuns++
				continue
			}
			if run.Text != nil {
				context.textUnits += len(utf16.Encode([]rune(*run.Text)))
				if context.textUnits > nativeApproximateTextboxTextLimit {
					return nil, "text-limit"
				}
			}
			run.ID = fmt.Sprintf("%s:r%d", paragraph.ID, runIndex)
			runs = append(runs, run)
		}
		paragraph.Runs = runs
		localResolver.resolveParagraph(&paragraph, result, numberingState, nil)
		textbox.Paragraphs = append(textbox.Paragraphs, paragraph)
	}
	if localResolver.diagnosticOverflow {
		return nil, "unsupported-textbox-content"
	}
	textbox.ResolvedParagraphs = result.Paragraphs
	textbox.ResolvedRuns = result.Runs
	if textbox.ResolvedParagraphs == nil {
		textbox.ResolvedParagraphs = []NativeResolvedParagraphV1{}
	}
	if textbox.ResolvedRuns == nil {
		textbox.ResolvedRuns = []NativeResolvedRunV1{}
	}
	return textbox, ""
}

// contentExtractor is a private extractor over the same package. Its
// diagnostics are counted as omissions and never merged into the document.
func (context *nativeApproximateShapeContext) contentExtractor() *nativeExtractor {
	relNS := relNSTransitional
	if context.ns == wordMLStrict {
		relNS = relNSStrict
	}
	return &nativeExtractor{
		pkg: context.resolver.pkg, wordNS: context.ns, relNS: relNS, mainPart: context.main, mainRoot: context.resolver.mainRoot,
		modeledParts: map[string]bool{context.main: true}, unsupportedSet: map[string]bool{}, storyByRel: map[string]string{}, storyByNative: map[string]string{}, commentByNative: map[string]string{},
		unsupported: []NativeUnsupportedCapabilityV1{}, previousIDs: map[string][]string{}, previousUsed: map[string]int{}, previousNative: map[string]string{}, previousPath: map[string]string{}, reservedIDs: map[string]bool{}, allocatedIDs: map[string]bool{}, seenParaIDs: map[string]string{},
	}
}
