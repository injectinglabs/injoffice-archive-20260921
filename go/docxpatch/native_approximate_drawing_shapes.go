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
// straight-edge polygon presets at their documented default adjust values,
// solid or theme-referenced fills, solid outlines and wp:inline / wp:anchor
// placement. Everything else stays omitted with a declared reason. It never
// changes source bytes, native extraction, editing authority or pagination.
const NativeApproximateDrawingShapesProtocol = "injoffice.docx.approximate-drawing-shapes"
const NativeApproximateDrawingShapePolicy = "docx.approximate-drawing-shape-preview-v1"
const nativeApproximateDrawingShapeLimit = 64
const nativeApproximateTextboxParagraphLimit = 256
const nativeApproximateTextboxTextLimit = 100000
const nativeMarkupCompatibilityNS = "http://schemas.openxmlformats.org/markup-compatibility/2006"

// nativeApproximateWPG is the wordprocessingGroup namespace, which is both the
// graphicData uri and the element namespace of a DrawingML group shape.
const nativeApproximateWPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"

// nativeApproximateEMULimit bounds every EMU coordinate the sidecar reports, so
// group child mapping cannot overflow int64 and stays inside the wire's range.
const nativeApproximateEMULimit = 127000000

func nativeApproximateWithinEMU(value int64) bool {
	return value >= -nativeApproximateEMULimit && value <= nativeApproximateEMULimit
}

type NativeApproximateShapeLineV1 struct {
	RGB      string `json:"rgb"`
	WidthEMU int64  `json:"width_emu"`
	Dash     string `json:"dash"`
}

// nativeApproximateGradientStopLimit bounds an a:gsLst the preview projects.
const nativeApproximateGradientStopLimit = 16

// NativeApproximateShapeGradientStopV1 is one a:gs. Position is 1/1000 of a
// percent along the gradient axis, as a:gs/@pos states it.
type NativeApproximateShapeGradientStopV1 struct {
	PositionPct int64  `json:"position_pct"`
	RGB         string `json:"rgb"`
}

// NativeApproximateShapeGradientV1 is an a:gradFill whose direction is an
// a:lin. Angle is 1/60000 of a degree, clockwise from the positive x axis, and
// stops are ordered by strictly increasing position.
// NativeApproximateShapeBlipFillV1 is one a:blipFill resolved to the image part
// its blip embeds, with the source rectangle that survives the fill's own crop
// and stretch in one-hundred-thousandths. It carries identity only; the page
// paint compiler joins it to the preserved part and transports the bytes.
type NativeApproximateShapeBlipFillV1 struct {
	RelationshipID string               `json:"relationship_id"`
	MediaPart      string               `json:"media_part"`
	ContentType    string               `json:"content_type"`
	SourceCrop     *NativeDrawingCropV1 `json:"source_crop,omitempty"`
}

// nativeApproximateFillRectLimit bounds an a:fillRect overhang: 1000 times the
// shape on one side is already far beyond anything Word writes, and it keeps the
// crop composition inside int64.
const nativeApproximateFillRectLimit = 100000000

type NativeApproximateShapeGradientV1 struct {
	Angle int64                                  `json:"angle_60000ths"`
	Stops []NativeApproximateShapeGradientStopV1 `json:"stops"`
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
	ID              string                            `json:"id"`
	ParagraphID     string                            `json:"paragraph_id"`
	DiagnosticIDs   []string                          `json:"diagnostic_ids"`
	Anchor          NativeSourceAnchorV1              `json:"anchor"`
	RunAnchor       NativeSourceAnchorV1              `json:"run_anchor"`
	Status          string                            `json:"status"`
	Reason          string                            `json:"reason,omitempty"`
	Placement       string                            `json:"placement,omitempty"`
	Preset          string                            `json:"preset,omitempty"`
	WidthEMU        int64                             `json:"width_emu"`
	HeightEMU       int64                             `json:"height_emu"`
	RotationDegrees int64                             `json:"rotation_degrees"`
	FlipHorizontal  bool                              `json:"flip_horizontal"`
	FlipVertical    bool                              `json:"flip_vertical"`
	FillRGB         *string                           `json:"fill_rgb,omitempty"`
	FillGradient    *NativeApproximateShapeGradientV1 `json:"fill_gradient,omitempty"`
	BlipFill        *NativeApproximateShapeBlipFillV1 `json:"blip_fill,omitempty"`
	Line            *NativeApproximateShapeLineV1     `json:"line,omitempty"`
	PageAnchor      *NativeTextboxPageAnchorV1        `json:"page_anchor,omitempty"`
	Wrap            string                            `json:"wrap,omitempty"`
	Textbox         *NativeApproximateTextboxV1       `json:"textbox,omitempty"`
	Notes           []string                          `json:"notes,omitempty"`
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
				drawing, alternateReason := context.drawingAlternate(child)
				if drawing == nil && alternateReason == "" {
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
				if drawing == nil {
					// The alternate selected no readable branch. Describe the
					// omission so the dropped source content is disclosed.
					item := context.newItem(p.ID, joined, run, child)
					item.Reason = alternateReason
					if len(out.Items) >= nativeApproximateDrawingShapeLimit {
						out.OmittedCount++
						continue
					}
					out.Items = append(out.Items, item)
					continue
				}
				if uri := nativeApproximateGraphicURI(drawing, wp, a); uri == nativeChartNSTransitional || uri == nativeChartNSStrict {
					// Charts belong to InspectNativeApproximateDrawingChartsV1.
					continue
				}
				// A group shape describes one item per child, so the budget is
				// spent per described shape rather than per drawing.
				for _, described := range context.describe(p.ID, joined, run, drawing) {
					if len(out.Items) >= nativeApproximateDrawingShapeLimit {
						out.OmittedCount++
						continue
					}
					out.Items = append(out.Items, described)
				}
			}
		}
	}
	if len(out.Items) == 0 && out.OmittedCount == 0 {
		return nil, nil
	}
	return out, nil
}

// nativeApproximateUnderstoodNamespaces names every namespace this preview can
// actually read out of a markup-compatibility branch: the DrawingML shape and
// group-shape vocabularies. A mc:Choice requiring anything else — wp14 relative
// positioning, a14 drawing extensions, VML — is not understood, so the Part 3
// rule sends the reader to the mc:Fallback instead.
var nativeApproximateUnderstoodNamespaces = map[string]bool{
	nativeTextboxWPS:     true,
	nativeApproximateWPG: true,
}

// drawingNode is the chart sidecar's view of drawingAlternate: the w:drawing a
// run child contributes, without the reason an alternate contributed none.
func (context *nativeApproximateShapeContext) drawingNode(child *nativeXMLNode) *nativeXMLNode {
	drawing, _ := context.drawingAlternate(child)
	return drawing
}

// drawingAlternate returns the w:drawing a run child contributes: either
// directly, or through the markup-compatibility alternate Word writes around a
// wps/wpg shape and its VML twin. The branch is selected by the Part 3 rule, so
// a Fallback is read only when no Choice named an understood namespace, and a
// VML-only fallback yields no drawing.
//
// A nil drawing with an empty reason means the child is not drawing markup at
// all and keeps its own source refusal. A nil drawing with a reason means an
// alternate WAS present and contributed no shape; that must be disclosed.
func (context *nativeApproximateShapeContext) drawingAlternate(child *nativeXMLNode) (*nativeXMLNode, string) {
	if child.Name == (xml.Name{Space: context.ns, Local: "drawing"}) {
		return child, ""
	}
	if !nativeMCIsAlternate(child) {
		return nil, ""
	}
	branch, outcome := nativeMCSelectAlternate(child, nativeApproximateUnderstoodNamespaces)
	switch outcome {
	case nativeMCInvalid:
		return nil, nativeMCReasonInvalid
	case nativeMCUnselected:
		return nil, nativeMCReasonUnselected
	}
	selected, reason := nativeMCResolvedChildren(branch, nativeApproximateUnderstoodNamespaces)
	if len(selected) == 0 {
		// A branch that contributes no element drops source content. It is
		// disclosed, never treated as "nothing to draw".
		if reason == "" {
			reason = nativeMCReasonEmptyBranch
		}
		return nil, reason
	}
	if len(selected) != 1 || selected[0].Name != (xml.Name{Space: context.ns, Local: "drawing"}) {
		// A fallback's w:pict, or any branch this preview cannot read, keeps
		// the source refusal it already carries.
		return nil, ""
	}
	return selected[0], ""
}

func (context *nativeApproximateShapeContext) anchor(n *nativeXMLNode) NativeSourceAnchorV1 {
	return nativeTextboxSourceAnchor(n, context.main, context.raw)
}

// newItem allocates the deterministic id and the source joins every shape
// described for one drawing shares. A group shape allocates one per child.
func (context *nativeApproximateShapeContext) newItem(paragraphID string, diagnosticIDs []string, run, drawing *nativeXMLNode) NativeApproximateDrawingShapeV1 {
	digest := strings.TrimPrefix(nativeSHA(context.raw[drawing.Start:drawing.End]), "sha256:")[:16]
	context.shapeSeen[digest]++
	return NativeApproximateDrawingShapeV1{ID: fmt.Sprintf("approximate-drawing-shape:%s:%d", digest, context.shapeSeen[digest]), ParagraphID: paragraphID, DiagnosticIDs: diagnosticIDs, Anchor: context.anchor(drawing), RunAnchor: context.anchor(run), Status: "omitted"}
}

// describe returns every approximate shape one drawing contributes: exactly one
// for a standalone wps:wsp, and one per child for a wpg:wgp group shape.
func (context *nativeApproximateShapeContext) describe(paragraphID string, diagnosticIDs []string, run, drawing *nativeXMLNode) []NativeApproximateDrawingShapeV1 {
	item := context.newItem(paragraphID, diagnosticIDs, run, drawing)
	omit := func(reason string) []NativeApproximateDrawingShapeV1 {
		item.Reason = reason
		return []NativeApproximateDrawingShapeV1{item}
	}
	// A run that also carries text (or anything but its properties) would leave
	// modeled runs nested inside the shape's run anchor; only the drawing may
	// share the w:r with w:rPr.
	for _, sibling := range run.Children {
		if sibling.Name == (xml.Name{Space: context.ns, Local: "rPr"}) || sibling == drawing || (sibling.Name == (xml.Name{Space: nativeMarkupCompatibilityNS, Local: "AlternateContent"}) && context.drawingNode(sibling) == drawing) {
			continue
		}
		return omit("shared-run")
	}
	wp, a := context.wp, context.a
	drawingChildren, reason := nativeMCResolvedChildren(drawing, nativeApproximateUnderstoodNamespaces)
	if reason != "" {
		return omit(reason)
	}
	var container *nativeXMLNode
	for _, c := range drawingChildren {
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
	// A markup-compatibility alternate can wrap the container's own children,
	// so resolve them before reading the extent and the graphic out of them.
	containerChildren, reason := nativeMCResolvedChildren(container, nativeApproximateUnderstoodNamespaces)
	if reason != "" {
		return omit(reason)
	}
	extent := nativeMCFirstChild(containerChildren, wp, "extent")
	if extent == nil {
		return omit("missing-extent")
	}
	width, okW := nativePositiveInt64Attr(extent, "", "cx")
	height, okH := nativePositiveInt64Attr(extent, "", "cy")
	if !okW || !okH || width > 127000000 || height > 127000000 {
		return omit("invalid-extent")
	}
	item.WidthEMU, item.HeightEMU = width, height
	graphic := nativeMCFirstChild(containerChildren, a, "graphic")
	if graphic == nil {
		return omit("missing-graphic")
	}
	graphicData := firstDirectNativeChild(graphic, a, "graphicData")
	if graphicData == nil {
		return omit("missing-graphic-data")
	}
	uri, _ := nativeUnqualifiedAttr(graphicData, "uri")
	if uri == nativeApproximateWPG {
		return context.describeGroup(item, paragraphID, diagnosticIDs, run, drawing, container, graphicData)
	}
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
	if reason := context.describeShape(&item, shape); reason != "" {
		return omit(reason)
	}
	if container.Name.Local == "inline" {
		item.Placement = "inline"
	} else {
		item.Placement = "anchored"
		pageAnchor, wrap, anchorReason := context.pageAnchor(container)
		if anchorReason != "" {
			return omit(anchorReason)
		}
		item.PageAnchor, item.Wrap = pageAnchor, wrap
		if wrap != "none" {
			item.Notes = append(item.Notes, "body text wrapping around the shape is not applied")
		}
	}
	textbox, textboxReason := context.textbox(shape, item.ID)
	if textboxReason != "" {
		item.Notes = append(item.Notes, "textbox content omitted: "+textboxReason)
	} else {
		item.Textbox = textbox
	}
	item.Status = "supported"
	return []NativeApproximateDrawingShapeV1{item}
}

// nativeApproximatePolygonPresets are the straight-edge prstGeom presets whose
// ECMA-376 preset geometry default adjust values resolve to one exact closed polygon
// in the shape's own extent. They are admitted only at those defaults (the
// adjust-values refusal below still rejects any authored adjustment), so the
// preview never interpolates a guide it has not implemented. The painter
// derives each outline from the same default guides; anything else keeps its
// unsupported-preset refusal.
var nativeApproximatePolygonPresets = map[string]bool{
	"downArrow":  true,
	"upArrow":    true,
	"leftArrow":  true,
	"rightArrow": true,
	"star5":      true,
}

// nativeApproximatePresetDefaultAdjustments are the a:avLst values ECMA-376
// preset geometry already declares for each admitted preset. Word writes the
// preset defaults out in full on a shape nobody has adjusted, so an a:avLst
// that only restates them describes exactly the outline the painter derives
// from those same defaults and is admitted; a value that differs by any amount
// is a real adjustment this tier does not implement and still refuses. A preset
// absent from this map (rect, line) admits no a:gd at all.
var nativeApproximatePresetDefaultAdjustments = map[string]map[string]int64{
	"downArrow":  {"adj1": 50000, "adj2": 50000},
	"upArrow":    {"adj1": 50000, "adj2": 50000},
	"leftArrow":  {"adj1": 50000, "adj2": 50000},
	"rightArrow": {"adj1": 50000, "adj2": 50000},
	"star5":      {"adj": 19098, "hf": 105146, "vf": 110557},
}

// nativeApproximateDefaultAdjustments reports whether every child of an a:avLst
// is an a:gd naming an adjustment of this preset and restating its documented
// default exactly, as `fmla="val N"`.
func nativeApproximateDefaultAdjustments(avLst *nativeXMLNode, a, preset string) bool {
	defaults := nativeApproximatePresetDefaultAdjustments[preset]
	for _, gd := range avLst.Children {
		if gd.Name != (xml.Name{Space: a, Local: "gd"}) {
			return false
		}
		name, ok := nativeUnqualifiedAttr(gd, "name")
		if !ok {
			return false
		}
		expected, known := defaults[name]
		if !known {
			return false
		}
		formula, ok := nativeUnqualifiedAttr(gd, "fmla")
		if !ok || !strings.HasPrefix(formula, "val ") {
			return false
		}
		value, err := strconv.ParseInt(strings.TrimSpace(strings.TrimPrefix(formula, "val ")), 10, 64)
		if err != nil || value != expected {
			return false
		}
	}
	return true
}

// describeShape reads one wps:wsp into an item whose WidthEMU/HeightEMU are
// already its placed extent, and returns the omission reason or "".
func (context *nativeApproximateShapeContext) describeShape(item *NativeApproximateDrawingShapeV1, shape *nativeXMLNode) string {
	a := context.a
	spPr := firstDirectNativeChild(shape, nativeTextboxWPS, "spPr")
	if spPr == nil {
		return "missing-shape-properties"
	}
	style := firstDirectNativeChild(shape, nativeTextboxWPS, "style")
	geometry := firstDirectNativeChild(spPr, a, "prstGeom")
	if geometry == nil {
		if firstDirectNativeChild(spPr, a, "custGeom") != nil {
			return "custom-geometry"
		}
		return "missing-geometry"
	}
	preset, _ := nativeUnqualifiedAttr(geometry, "prst")
	if preset != "rect" && preset != "line" && !nativeApproximatePolygonPresets[preset] {
		return "unsupported-preset:" + preset
	}
	for _, c := range geometry.Children {
		if c.Name == (xml.Name{Space: a, Local: "avLst"}) && len(c.Children) > 0 && !nativeApproximateDefaultAdjustments(c, a, preset) {
			return "adjust-values"
		}
	}
	item.Preset = preset
	if xfrm := firstDirectNativeChild(spPr, a, "xfrm"); xfrm != nil {
		if rot, ok := nativeUnqualifiedAttr(xfrm, "rot"); ok {
			value, err := strconv.ParseInt(rot, 10, 64)
			if err != nil || value < -21600000*100 || value > 21600000*100 {
				return "invalid-rotation"
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
		return "rotation-unsupported"
	}
	if preset == "rect" && (item.RotationDegrees == 90 || item.RotationDegrees == 270) {
		item.Notes = append(item.Notes, "quarter-turn rectangle painted as its rotated bounding box")
	}
	// A polygon preset is admitted only at its authored extent: this preview
	// paints the default-adjust outline itself, so a rotation or a flip it would
	// have to apply to every vertex stays omitted instead of being guessed.
	if nativeApproximatePolygonPresets[preset] {
		if item.RotationDegrees != 0 {
			return "rotation-unsupported"
		}
		if item.FlipHorizontal || item.FlipVertical {
			return "flip-unsupported"
		}
	}
	fill, fillNotes, fillOK := context.shapeFill(spPr, style)
	if !fillOK {
		return "unsupported-fill"
	}
	item.FillRGB, item.FillGradient, item.BlipFill = fill.RGB, fill.Gradient, fill.Blip
	// The picture paints axis-aligned in the shape's placed box, so only a
	// rectangle the source neither rotates nor flips carries it.
	if item.BlipFill != nil && (preset != "rect" || item.RotationDegrees != 0 || item.FlipHorizontal || item.FlipVertical) {
		item.BlipFill = nil
		fillNotes = append(fillNotes, "blipFill on a rotated, flipped or non-rectangular shape is not approximated; fill omitted")
	}
	item.Notes = append(item.Notes, fillNotes...)
	line, lineNotes, lineOK := context.shapeLine(spPr, style)
	if !lineOK {
		return "unsupported-outline"
	}
	item.Line = line
	item.Notes = append(item.Notes, lineNotes...)
	if item.Line != nil && (item.Line.WidthEMU >= item.WidthEMU && item.Line.WidthEMU >= item.HeightEMU) {
		return "outline-exceeds-shape"
	}
	return ""
}

// nativeApproximateScaleEMU maps a child coordinate onto the group's placed
// extent: value * placed / span, rounded half away from zero. Inputs are bounded
// to EMU page coordinates by the caller so the product cannot overflow int64.
func nativeApproximateScaleEMU(value, placed, span int64) int64 {
	if span <= 0 {
		return 0
	}
	product := value * placed
	if product < 0 {
		return -((-product*2 + span) / (2 * span))
	}
	return (product*2 + span) / (2 * span)
}

// describeGroup maps every wps:wsp child of a wpg:wgp into the group's placed
// extent and describes each as its own approximate anchored shape. The group's
// a:xfrm defines the child coordinate space: a child's placed offset is
// (a:off - a:chOff) scaled by a:ext / a:chExt, and its placed extent is a:ext
// scaled the same way. Nested groups, children whose transform cannot be mapped
// exactly, and group transforms this preview cannot reproduce stay omitted with
// their own reason instead of being guessed.
func (context *nativeApproximateShapeContext) describeGroup(base NativeApproximateDrawingShapeV1, paragraphID string, diagnosticIDs []string, run, drawing, container, graphicData *nativeXMLNode) []NativeApproximateDrawingShapeV1 {
	a := context.a
	omit := func(reason string) []NativeApproximateDrawingShapeV1 {
		base.Reason = reason
		return []NativeApproximateDrawingShapeV1{base}
	}
	groups := directNativeChildren(graphicData, nativeApproximateWPG, "wgp")
	if len(groups) != 1 || len(graphicData.Children) != 1 {
		return omit("unsupported-graphic:group-or-multiple")
	}
	group := groups[0]
	if container.Name.Local != "anchor" {
		// An inline group would have to reserve one line atom for the whole
		// group; this preview only reserves per shape, so it stays omitted.
		return omit("inline-group-unsupported")
	}
	anchor, wrap, reason := context.pageAnchor(container)
	if reason != "" {
		return omit(reason)
	}
	if anchor.HorizontalAlign != "" || anchor.VerticalAlign != "" {
		// Alignment places the group's own extent, which the sidecar would have
		// to resolve against the page before it could place children inside it.
		return omit("aligned-group-position-unsupported")
	}
	properties := firstDirectNativeChild(group, nativeApproximateWPG, "grpSpPr")
	if properties == nil {
		return omit("missing-group-properties")
	}
	xfrm := firstDirectNativeChild(properties, a, "xfrm")
	if xfrm == nil {
		return omit("missing-child-coordinate-space")
	}
	if rot, ok := nativeUnqualifiedAttr(xfrm, "rot"); ok && strings.TrimSpace(rot) != "0" {
		return omit("group-rotation-unsupported")
	}
	if nativeApproximateFlag(xfrm, "flipH") || nativeApproximateFlag(xfrm, "flipV") {
		return omit("group-flip-unsupported")
	}
	if offset := firstDirectNativeChild(xfrm, a, "off"); offset != nil {
		// A top-level group is placed by wp:anchor; a non-zero a:off would move
		// it again by an amount this preview does not model.
		x, okX := nativeInt64Attr(offset, "", "x")
		y, okY := nativeInt64Attr(offset, "", "y")
		if !okX || !okY || x != 0 || y != 0 {
			return omit("group-offset-unsupported")
		}
	}
	extent, childExtent, childOffset := firstDirectNativeChild(xfrm, a, "ext"), firstDirectNativeChild(xfrm, a, "chExt"), firstDirectNativeChild(xfrm, a, "chOff")
	if extent == nil {
		return omit("missing-child-coordinate-space")
	}
	groupCX, okCX := nativePositiveInt64Attr(extent, "", "cx")
	groupCY, okCY := nativePositiveInt64Attr(extent, "", "cy")
	if !okCX || !okCY || groupCX > nativeApproximateEMULimit || groupCY > nativeApproximateEMULimit {
		return omit("invalid-child-coordinate-space")
	}
	placedCX, placedCY := groupCX, groupCY
	var spanCX, spanCY, originX, originY int64
	if childExtent != nil || childOffset != nil {
		// A stated child coordinate space maps children onto the object's
		// displayed size, so wp:extent and the group's own a:ext must agree
		// before either can be chosen as that size.
		if childExtent == nil || childOffset == nil {
			return omit("missing-child-coordinate-space")
		}
		if groupCX != base.WidthEMU || groupCY != base.HeightEMU {
			return omit("group-extent-mismatch")
		}
		placedCX, placedCY = base.WidthEMU, base.HeightEMU
		var okSX, okSY, okOX, okOY bool
		spanCX, okSX = nativePositiveInt64Attr(childExtent, "", "cx")
		spanCY, okSY = nativePositiveInt64Attr(childExtent, "", "cy")
		originX, okOX = nativeInt64Attr(childOffset, "", "x")
		originY, okOY = nativeInt64Attr(childOffset, "", "y")
		if !okSX || !okSY || !okOX || !okOY || spanCX > nativeApproximateEMULimit || spanCY > nativeApproximateEMULimit || !nativeApproximateWithinEMU(originX) || !nativeApproximateWithinEMU(originY) {
			return omit("invalid-child-coordinate-space")
		}
	} else {
		// ECMA-376 Part 1 §20.1.7.6 makes a:chOff and a:chExt optional. A group
		// that states neither leaves its children in the group's own coordinate
		// space: the child origin is (0, 0) and the child span is a:ext, so the
		// mapping below is the identity and wp:extent is only the wrap box.
		// Word paints such a group at a:ext — checked against its own PDF export
		// of dml-groupshape-childposition, whose 193680x9125640 EMU child fills
		// exactly 15.2504x718.554 pt with no wp:extent rescale. Stating one of
		// the pair without the other leaves half the mapping unattested, so that
		// keeps refusing above.
		spanCX, spanCY = groupCX, groupCY
		if groupCX != base.WidthEMU || groupCY != base.HeightEMU {
			base.Notes = append(base.Notes, "group has no child coordinate space and is painted at its own a:ext; wp:extent differs and bounds only the wrap region")
		}
	}
	scaled := placedCX != spanCX || placedCY != spanCY
	pending := &base
	nextItem := func() *NativeApproximateDrawingShapeV1 {
		if pending != nil {
			item := pending
			pending = nil
			item.WidthEMU, item.HeightEMU = 0, 0
			return item
		}
		fresh := context.newItem(paragraphID, diagnosticIDs, run, drawing)
		return &fresh
	}
	items := []NativeApproximateDrawingShapeV1{}
	for _, child := range group.Children {
		if child == properties || child.Name == (xml.Name{Space: nativeApproximateWPG, Local: "cNvGrpSpPr"}) {
			continue
		}
		item := nextItem()
		if child.Name == (xml.Name{Space: nativeApproximateWPG, Local: "grpSp"}) {
			item.Reason = "nested-group"
			items = append(items, *item)
			continue
		}
		if child.Name != (xml.Name{Space: nativeTextboxWPS, Local: "wsp"}) {
			item.Reason = "unsupported-group-child:" + child.Name.Local
			items = append(items, *item)
			continue
		}
		spPr := firstDirectNativeChild(child, nativeTextboxWPS, "spPr")
		if spPr == nil {
			item.Reason = "missing-shape-properties"
			items = append(items, *item)
			continue
		}
		childXfrm := firstDirectNativeChild(spPr, a, "xfrm")
		var offset, size *nativeXMLNode
		if childXfrm != nil {
			offset, size = firstDirectNativeChild(childXfrm, a, "off"), firstDirectNativeChild(childXfrm, a, "ext")
		}
		if offset == nil || size == nil {
			item.Reason = "missing-child-transform"
			items = append(items, *item)
			continue
		}
		if rot, ok := nativeUnqualifiedAttr(childXfrm, "rot"); ok && strings.TrimSpace(rot) != "0" {
			// A child rotation composes with the group's own scale; the preview
			// paints axis-aligned boxes, so a rotated child is not reproduced.
			item.Reason = "child-rotation-unsupported"
			items = append(items, *item)
			continue
		}
		x, okX := nativeInt64Attr(offset, "", "x")
		y, okY := nativeInt64Attr(offset, "", "y")
		cx, okCX := nativePositiveInt64Attr(size, "", "cx")
		cy, okCY := nativePositiveInt64Attr(size, "", "cy")
		if !okX || !okY || !okCX || !okCY || !nativeApproximateWithinEMU(x) || !nativeApproximateWithinEMU(y) || cx > nativeApproximateEMULimit || cy > nativeApproximateEMULimit {
			item.Reason = "invalid-child-transform"
			items = append(items, *item)
			continue
		}
		width, height := nativeApproximateScaleEMU(cx, placedCX, spanCX), nativeApproximateScaleEMU(cy, placedCY, spanCY)
		placedX := anchor.XEMU + nativeApproximateScaleEMU(x-originX, placedCX, spanCX)
		placedY := anchor.YEMU + nativeApproximateScaleEMU(y-originY, placedCY, spanCY)
		if width <= 0 || height <= 0 {
			item.Reason = "degenerate-child-extent"
			items = append(items, *item)
			continue
		}
		if !nativeApproximateWithinEMU(placedX) || !nativeApproximateWithinEMU(placedY) {
			item.Reason = "child-outside-coordinate-range"
			items = append(items, *item)
			continue
		}
		item.WidthEMU, item.HeightEMU = width, height
		if reason := context.describeShape(item, child); reason != "" {
			item.Reason = reason
			item.Preset, item.FillRGB, item.BlipFill, item.Line, item.Notes = "", nil, nil, nil, nil
			item.WidthEMU, item.HeightEMU, item.RotationDegrees, item.FlipHorizontal, item.FlipVertical = 0, 0, 0, false, false
			items = append(items, *item)
			continue
		}
		childAnchor := *anchor
		childAnchor.XEMU, childAnchor.YEMU = placedX, placedY
		item.Placement, item.PageAnchor, item.Wrap = "anchored", &childAnchor, wrap
		if wrap != "none" {
			item.Notes = append(item.Notes, "body text wrapping around the shape is not applied")
		}
		item.Notes = append(item.Notes, "group shape child placed by mapping its child coordinates into the group's declared extent")
		if scaled {
			item.Notes = append(item.Notes, "group child coordinates are scaled onto the group extent; text box insets and font sizes inside the group are not scaled")
		}
		textbox, reason := context.textbox(child, item.ID)
		if reason != "" {
			item.Notes = append(item.Notes, "textbox content omitted: "+reason)
		} else {
			item.Textbox = textbox
		}
		item.Status = "supported"
		items = append(items, *item)
	}
	if len(items) == 0 {
		return omit("empty-group")
	}
	return items
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
	// Word writes wp14 relative positioning as a markup-compatibility alternate
	// around wp:positionH/wp:positionV. wp14 is not understood here, so the
	// Part 3 rule reads the mc:Fallback's plain wp:posOffset placement instead.
	children, reason := nativeMCResolvedChildren(container, nativeApproximateUnderstoodNamespaces)
	if reason != "" {
		return nil, "", reason
	}
	positionH := nativeMCFirstChild(children, wp, "positionH")
	positionV := nativeMCFirstChild(children, wp, "positionV")
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
	for _, child := range children {
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

// nativeApproximateFill is one resolved shape fill: at most one of RGB and
// Gradient is set, and both nil means the fill is absent or omitted.
type nativeApproximateFill struct {
	RGB      *string
	Gradient *NativeApproximateShapeGradientV1
	Blip     *NativeApproximateShapeBlipFillV1
}

// shapeFill returns an empty fill for no fill. A linear a:gradFill is projected
// as its stop list; pattern and picture fills are not approximated, and they
// omit the fill and record the omission.
func (context *nativeApproximateShapeContext) shapeFill(spPr, style *nativeXMLNode) (nativeApproximateFill, []string, bool) {
	a := context.a
	none := nativeApproximateFill{}
	notes := []string{}
	for _, child := range spPr.Children {
		if child.Name.Space != a {
			continue
		}
		switch child.Name.Local {
		case "noFill":
			return none, notes, true
		case "solidFill":
			rgb, colorNotes, ok := context.solidColor(child, "")
			if !ok {
				return none, nil, false
			}
			return nativeApproximateFill{RGB: nativeString(rgb)}, append(notes, colorNotes...), true
		case "gradFill":
			gradient, gradientNotes, reason := context.linearGradientFill(child, "")
			notes = append(notes, gradientNotes...)
			if gradient == nil {
				return none, append(notes, "gradFill "+reason+"; fill omitted"), true
			}
			return nativeApproximateFill{Gradient: gradient}, notes, true
		case "blipFill":
			blip, reason := context.blipFill(child)
			if blip == nil {
				return none, append(notes, "blipFill "+reason+"; fill omitted"), true
			}
			return nativeApproximateFill{Blip: blip}, notes, true
		case "pattFill", "grpFill":
			return none, append(notes, child.Name.Local+" is not approximated; fill omitted"), true
		}
	}
	if style == nil {
		return none, notes, true
	}
	reference := firstDirectNativeChild(style, a, "fillRef")
	if reference == nil {
		return none, notes, true
	}
	index, _ := nativeNonnegativeInt64Attr(reference, "", "idx")
	if index == 0 {
		return none, notes, true
	}
	phClr, colorNotes, ok := context.referenceColor(reference)
	if !ok {
		return none, nil, false
	}
	notes = append(notes, colorNotes...)
	fillStyle := context.theme.styleAt(context.theme.fills, index)
	if fillStyle == nil {
		return none, append(notes, "theme fill style unavailable; fill omitted"), true
	}
	if fillStyle.Name == (xml.Name{Space: context.theme.ns, Local: "gradFill"}) {
		gradient, gradientNotes, reason := context.linearGradientFill(fillStyle, phClr)
		notes = append(notes, gradientNotes...)
		if gradient == nil {
			return none, append(notes, "theme gradFill "+reason+"; fill omitted"), true
		}
		return nativeApproximateFill{Gradient: gradient}, append(notes, "fill resolved from theme fill style"), true
	}
	if fillStyle.Name != (xml.Name{Space: context.theme.ns, Local: "solidFill"}) {
		return none, append(notes, "theme "+fillStyle.Name.Local+" is not approximated; fill omitted"), true
	}
	rgb, themeNotes, ok := context.solidColor(fillStyle, phClr)
	if !ok {
		return none, nil, false
	}
	return nativeApproximateFill{RGB: nativeString(rgb)}, append(append(notes, "fill resolved from theme fill style"), themeNotes...), true
}

// linearGradientFill projects an a:gradFill whose direction is an a:lin.
//
// ECMA-376 Part 1 defines a:gradFill (§20.1.8.33), the a:gsLst stop list
// (§20.1.8.36) and the a:lin linear direction (§20.1.8.41), whose @ang is
// 1/60000 of a degree clockwise from the positive x axis. Only that linear
// form is projected: a:path shades along a rectangle or the shape outline and
// a:tileRect and @flip change the mapping, so each one omits the fill instead
// of painting a straight interpolation that is not the authored one.
//
// @scaled selects whether the angle is measured in the shape's own scaled
// space or unscaled. Those two readings coincide only where the axis is
// parallel to a box edge, so a scaled gradient is projected only at a multiple
// of 90 degrees and otherwise omits the fill rather than guessing the
// shape-dependent skew. Stop alpha is dropped, as everywhere else in this
// tier: Word paints these stops opaque.
//
// It returns (nil, notes, reason) when the fill is outside that subset.
func (context *nativeApproximateShapeContext) linearGradientFill(node *nativeXMLNode, phClr string) (*NativeApproximateShapeGradientV1, []string, string) {
	a := context.a
	if _, flipped := nativeUnqualifiedAttr(node, "flip"); flipped {
		return nil, nil, "declares a tile flip"
	}
	direction := firstDirectNativeChild(node, a, "lin")
	if direction == nil {
		return nil, nil, "has no a:lin linear direction"
	}
	for _, child := range node.Children {
		if child.Name.Space != a || (child.Name.Local != "gsLst" && child.Name.Local != "lin") {
			return nil, nil, "carries an unmodeled " + child.Name.Local
		}
	}
	angle, ok := nativeNonnegativeInt64Attr(direction, "", "ang")
	if !ok || angle >= 21600000 {
		return nil, nil, "declares no usable a:lin angle"
	}
	if scaled, _ := nativeUnqualifiedAttr(direction, "scaled"); (scaled == "1" || scaled == "true") && angle%5400000 != 0 {
		return nil, nil, "is scaled to a non-axis-aligned angle"
	}
	stopList := firstDirectNativeChild(node, a, "gsLst")
	if stopList == nil {
		return nil, nil, "has no a:gsLst stop list"
	}
	if len(stopList.Children) < 2 || len(stopList.Children) > nativeApproximateGradientStopLimit {
		return nil, nil, fmt.Sprintf("must hold 2..%d stops", nativeApproximateGradientStopLimit)
	}
	notes := []string{}
	stops := make([]NativeApproximateShapeGradientStopV1, 0, len(stopList.Children))
	previous := int64(-1)
	for _, stop := range stopList.Children {
		if stop.Name != (xml.Name{Space: a, Local: "gs"}) {
			return nil, nil, "stop list carries unmodeled markup"
		}
		position, ok := nativeNonnegativeInt64Attr(stop, "", "pos")
		if !ok || position > 100000 || position <= previous {
			return nil, nil, "stop positions must increase through 0..100000"
		}
		previous = position
		color := (*nativeXMLNode)(nil)
		for _, child := range stop.Children {
			if child.Name.Space == a {
				color = child
				break
			}
		}
		if color == nil {
			return nil, nil, "stop carries no colour"
		}
		rgb, colorNotes, ok := context.color(color, phClr)
		if !ok {
			return nil, nil, "stop colour is outside the approximated subset"
		}
		notes = append(notes, colorNotes...)
		stops = append(stops, NativeApproximateShapeGradientStopV1{PositionPct: position, RGB: rgb})
	}
	return &NativeApproximateShapeGradientV1{Angle: angle, Stops: stops}, notes, ""
}

// blipFill resolves one a:blipFill (ECMA-376 20.1.8.14) to the image part its
// blip embeds and the source rectangle that is painted, or returns the reason
// it is not approximated. a:srcRect crops the source; a:stretch/a:fillRect then
// places the cropped source in the shape, and a negative offset says the
// picture overhangs the shape on that side, which the shape's outline clips.
// Both compose into one crop of the source, in one-hundred-thousandths: the
// cropped span is stretched over the shape plus its overhangs, so the fraction
// left of the shape is span * overhang / (shape + overhangs). A tile, a
// positive inset (the picture would not cover the shape), a linked or
// effect-bearing blip, and a relationship that is not an internal image part
// stay omitted and say so.
func (context *nativeApproximateShapeContext) blipFill(node *nativeXMLNode) (*NativeApproximateShapeBlipFillV1, string) {
	a := context.a
	relNS, relBase := relNSTransitional, relBaseTransitional
	if context.ns == wordMLStrict {
		relNS, relBase = relNSStrict, relBaseStrict
	}
	if !nativeExactContainer(node, xml.Name{Local: "dpi"}, xml.Name{Local: "rotWithShape"}) {
		return nil, "carries unmodeled attributes"
	}
	rectNames := []xml.Name{{Local: "l"}, {Local: "t"}, {Local: "r"}, {Local: "b"}}
	var blip, stretch *nativeXMLNode
	crop := [4]int64{}
	for _, child := range node.Children {
		if child.Name.Space != a {
			return nil, "carries unmodeled children"
		}
		switch child.Name.Local {
		case "blip":
			if blip != nil {
				return nil, "must embed exactly one blip"
			}
			blip = child
		case "srcRect":
			if !nativeExactLeaf(child, rectNames...) {
				return nil, "srcRect carries unmodeled attributes"
			}
			for index, name := range []string{"l", "t", "r", "b"} {
				if _, present := nativeUnqualifiedAttr(child, name); !present {
					continue
				}
				value, ok := nativeInt64Attr(child, "", name)
				if !ok || value < 0 || value > 99000 {
					return nil, "srcRect is out of range"
				}
				crop[index] = value
			}
		case "stretch":
			if !nativeExactContainer(child) || len(child.Children) > 1 {
				return nil, "stretch is not a single fill rectangle"
			}
			if len(child.Children) == 1 {
				stretch = child.Children[0]
				if stretch.Name != (xml.Name{Space: a, Local: "fillRect"}) || !nativeExactLeaf(stretch, rectNames...) {
					return nil, "stretch is not a fill rectangle"
				}
			}
		case "tile":
			return nil, "tiles the picture"
		default:
			return nil, "carries unmodeled children"
		}
	}
	if blip == nil {
		return nil, "must embed exactly one blip"
	}
	if !nativeExactContainer(blip, xml.Name{Space: relNS, Local: "embed"}, xml.Name{Local: "cstate"}) || !nativeInertBlipExtensions(blip, a) {
		return nil, "blip is linked or carries effects"
	}
	relID, ok := nativeAttr(blip, relNS, "embed")
	if !ok || relID == "" {
		return nil, "blip embeds no relationship"
	}
	var rel *nativeRelationship
	for i := range context.resolver.pkg.rels[context.main] {
		if context.resolver.pkg.rels[context.main][i].ID == relID {
			rel = &context.resolver.pkg.rels[context.main][i]
		}
	}
	if rel == nil || rel.External || rel.Type != relBase+"image" || rel.PartName == "" {
		return nil, "relationship is not an internal image part"
	}
	contentType := context.resolver.pkg.contentTypes[rel.PartName]
	if !strings.HasPrefix(nativeASCIIFold(contentType), "image/") {
		return nil, "image part has no image content type"
	}
	if crop[0]+crop[2] > 99000 || crop[1]+crop[3] > 99000 {
		return nil, "srcRect leaves less than one percent of the picture"
	}
	if stretch != nil {
		over := [4]int64{}
		for index, name := range []string{"l", "t", "r", "b"} {
			if _, present := nativeUnqualifiedAttr(stretch, name); !present {
				continue
			}
			value, ok := nativeInt64Attr(stretch, "", name)
			if !ok || value < -nativeApproximateFillRectLimit {
				return nil, "fillRect is out of range"
			}
			if value > 0 {
				return nil, "fillRect insets the picture inside the shape"
			}
			over[index] = -value
		}
		const scale = int64(100000)
		for _, axis := range [][2]int{{0, 2}, {1, 3}} {
			lo, hi := axis[0], axis[1]
			span, total := scale-crop[lo]-crop[hi], scale+over[lo]+over[hi]
			crop[lo] += (2*span*over[lo] + total) / (2 * total)
			crop[hi] += (2*span*over[hi] + total) / (2 * total)
		}
		if crop[0]+crop[2] > 99000 || crop[1]+crop[3] > 99000 {
			return nil, "fillRect overhang leaves less than one percent of the picture"
		}
	}
	fill := &NativeApproximateShapeBlipFillV1{RelationshipID: relID, MediaPart: rel.PartName, ContentType: contentType}
	if crop != [4]int64{} {
		fill.SourceCrop = &NativeDrawingCropV1{Left: nativeInt64(crop[0]), Top: nativeInt64(crop[1]), Right: nativeInt64(crop[2]), Bottom: nativeInt64(crop[3])}
	}
	return fill, ""
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
			if styled := context.theme.styleAt(context.theme.lines, index); styled != nil {
				color, colorNotes, ok := context.referenceColor(reference)
				if !ok {
					return nil, nil, false
				}
				phClr = color
				notes = append(notes, colorNotes...)
				themeLine = styled
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
			rgb = nativeApproximateHSL(rgb, float64(value)/100000, 0, 1)
			notes = append(notes, "luminance transform approximated in sRGB")
		case "lumOff":
			rgb = nativeApproximateHSL(rgb, 1, float64(value)/100000, 1)
			notes = append(notes, "luminance transform approximated in sRGB")
		case "satMod":
			rgb = nativeApproximateHSL(rgb, 1, 0, float64(value)/100000)
			notes = append(notes, "saturation transform approximated in sRGB")
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

// nativeApproximateHSL applies the DrawingML luminance and saturation
// modulations in the same sRGB-derived HSL space. DrawingML does not specify
// the colour space for lumMod/lumOff/satMod, so the result is approximate and
// every caller discloses it.
func nativeApproximateHSL(rgb string, mod, off, satMod float64) string {
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
	s *= satMod
	if s > 1 {
		s = 1
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

// styleAt selects the 1-based fmtScheme entry without converting the parsed
// index to a narrower integer type; idx 0 and out-of-range values select nothing.
func (theme *nativeApproximateTheme) styleAt(styles []*nativeXMLNode, index int64) *nativeXMLNode {
	if theme == nil || index < 1 {
		return nil
	}
	for position, style := range styles {
		if int64(position)+1 == index {
			return style
		}
	}
	return nil
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
