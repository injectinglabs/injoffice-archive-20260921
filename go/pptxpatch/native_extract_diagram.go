package pptxpatch

import (
	"encoding/xml"
	"errors"
	"fmt"
	"strconv"
	"strings"
)

// SmartArt (DrawingML diagram) graphic frames reference four diagram parts
// (data, layout, quick style, colors). Evaluating the layout algorithm is
// out of scope for native PPTX v1. PowerPoint additionally persists a
// pre-laid-out fallback drawing (dsp:drawing, related from the slide through
// the diagramDrawing relationship and named by dsp:dataModelExt inside the
// data model). Like LibreOffice, native extract paints that drawing verbatim:
// each dsp:sp becomes an ordinary read-only shape inside one group anchored
// at the graphic frame. Nothing is laid out, invented, or rewritten; when the
// drawing part is missing, empty, or outside the exact subset the whole frame
// is refused and preserved as opaque bytes.
const (
	nativeDiagramURITransitional = "http://schemas.openxmlformats.org/drawingml/2006/diagram"
	nativeDiagramURIStrict       = "http://purl.oclc.org/ooxml/drawingml/diagram"
	nsDiagramDrawing             = "http://schemas.microsoft.com/office/drawing/2008/diagram"
	relDiagramDrawing            = "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
	contentTypeDiagramData       = "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"
	contentTypeDiagramDrawing    = "application/vnd.ms-office.drawingml.diagramDrawing+xml"

	// nativeDiagramDrawingPreviewCode labels every element projected from a
	// diagram drawing part. Mutation guards treat it as preview-only.
	nativeDiagramDrawingPreviewCode = "pptx.diagram-drawing-fallback-preview"
	// nativeDiagramDrawingPolicy names the declared read-only preview policy.
	nativeDiagramDrawingPolicy          = "diagram-drawing-fallback-v1"
	nativeDiagramDrawingTextOmittedCode = "pptx.diagram-drawing-text-unavailable"
	nativeMaxDiagramDrawingShapes       = 2048
)

func nativeDiagramURI(dialect nativeExtractDialect) string {
	if dialect.drawing == nsDrawingStrict {
		return nativeDiagramURIStrict
	}
	return nativeDiagramURITransitional
}

func refuseNativeDiagram(code, message string) error {
	return refuseNativeGraphicFrame(code, message)
}

// extractNativeDiagramGraphicFrame projects a SmartArt graphic frame from its
// PowerPoint drawing fallback as one preserve-only group of shapes.
func (extractor *nativeExtractor) extractNativeDiagramGraphicFrame(node *nativeXMLNode, slidePart, slideID string, relationships []nativeExtractRelationship, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "graphicFrame"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid diagram graphic frame root")
	}
	if err := requireOnlyNativeAttrs(node); err != nil {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram graphic frame root attributes are outside the exact native subset")
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvGraphicFramePr"},
		xml.Name{Space: dialect.presentation, Local: "xfrm"},
		xml.Name{Space: dialect.drawing, Local: "graphic"}); err != nil {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram graphic frame contains unmodeled markup or direct text")
	}
	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvGraphicFramePr", true)
	if err != nil {
		return NativeElement{}, err
	}
	transformNode, err := nativeSingleton(node, dialect.presentation, "xfrm", true)
	if err != nil {
		return NativeElement{}, err
	}
	graphic, err := nativeSingleton(node, dialect.drawing, "graphic", true)
	if err != nil {
		return NativeElement{}, err
	}
	if len(node.Children) != 3 || node.Children[0] != nonVisual || node.Children[1] != transformNode || node.Children[2] != graphic {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: malformed diagram graphic frame child order")
	}
	objectID, name, err := validateNativeDiagramNonVisual(nonVisual, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	transform, err := validateNativeTableTransform(transformNode, dialect)
	if err != nil {
		var refusal nativeGraphicFrameProjectionRefusal
		if errors.As(err, &refusal) {
			return NativeElement{}, refuseNativeDiagram("pptx.diagram-transform-unavailable", "diagram frame transform is rotated, flipped, or outside the exact native subset")
		}
		return NativeElement{}, err
	}
	dataRelationshipID, err := extractNativeDiagramDataRelationshipID(graphic, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	drawingRel, err := extractor.resolveNativeDiagramDrawingRelationship(relationships, dataRelationshipID, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	shapes, drawingPayload, err := extractor.parseNativeDiagramDrawingShapes(drawingRel.Part)
	if err != nil {
		// Without stored fallback shapes the opt-in approximate tier lays the
		// diagram out from its parts (native_diagram_layout.go). The exact
		// tier and the populated-fallback path are unchanged.
		if extractor.options.AllowInheritedTextPreview && nativeDiagramRefusalCode(err) == nativeDiagramDrawingEmptyCode {
			return extractor.extractNativeDiagramLayoutGraphicFrame(node, graphic, slidePart, slideID, relationships, dialect, objectID, name, transform)
		}
		return NativeElement{}, err
	}
	children := make([]NativeElement, 0, len(shapes))
	for index, shape := range shapes {
		child, childErr := extractor.extractNativeDiagramDrawingShape(shape, index, drawingRel.Part, drawingPayload, slidePart, slideID, objectID, dialect)
		if childErr != nil {
			return NativeElement{}, childErr
		}
		children = append(children, child)
	}

	raw, err := rawNativeNode(extractor.pkg.parts[slidePart], node)
	if err != nil {
		return NativeElement{}, err
	}
	fingerprint := nativeSHA256(raw)
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	passthrough, err := extractor.issuePassthrough(slidePart, objectID, fingerprint, raw, nativeDiagramDrawingPreviewCode)
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return NativeElement{}, err
	}
	partName := drawingRel.Part
	element := NativeElement{
		Kind: NativeElementKindGroup, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform:      transform,
		ChildTransform: &NativeTransform{X: int64Pointer(0), Y: int64Pointer(0), Cx: int64Pointer(*transform.Cx), Cy: int64Pointer(*transform.Cy)},
		Children:       children,
		Passthrough:    []NativePassthroughRef{passthrough},
		Source:         &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusPreserveOnly, Diagnostics: []NativeDiagnostic{{
			Severity: NativeDiagnosticSeverityWarning, Code: nativeDiagramDrawingPreviewCode,
			Message: "SmartArt painted from the PowerPoint pre-laid-out diagram drawing part (" + nativeDiagramDrawingPolicy + "): shapes, solid paint, and text come from dsp:sp source markup; diagram layout is not evaluated, effects are omitted, and the frame remains read-only",
			Scope:   &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		}}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	return element, nil
}

func validateNativeDiagramNonVisual(node *nativeXMLNode, dialect nativeExtractDialect) (string, string, error) {
	if requireOnlyNativeAttrs(node) != nil || requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvGraphicFramePr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}) != nil {
		return "", "", refuseNativeDiagram("pptx.diagram-nonvisual-unavailable", "diagram nonvisual properties are outside the exact subset")
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	for _, local := range []string{"cNvGraphicFramePr", "nvPr"} {
		if _, err := nativeSingleton(node, dialect.presentation, local, false); err != nil {
			return "", "", err
		}
	}
	// hidden, hyperlink, or unknown identity attributes are not modeled; a
	// hidden frame must never paint.
	if requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}, xml.Name{Local: "descr"}, xml.Name{Local: "title"}) != nil ||
		requireOnlyNativeChildren(cNvPr, xml.Name{Space: dialect.drawing, Local: "extLst"}) != nil || !onlyNativeXMLSpace(cNvPr.Text) {
		return "", "", refuseNativeDiagram("pptx.diagram-nonvisual-unavailable", "diagram frame visibility, hyperlink, or unknown identity metadata is not modeled")
	}
	nativeID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", fmt.Errorf("pptxpatch: native extract: diagram name exceeds contract bound")
	}
	return "cNvPr-" + nativeID, name, nil
}

func extractNativeDiagramDataRelationshipID(graphic *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	if requireOnlyNativeAttrs(graphic) != nil || requireOnlyNativeChildren(graphic, xml.Name{Space: dialect.drawing, Local: "graphicData"}) != nil {
		return "", refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram graphic container is outside the exact native subset")
	}
	data, err := nativeSingleton(graphic, dialect.drawing, "graphicData", true)
	if err != nil {
		return "", err
	}
	if requireOnlyNativeAttrs(data, xml.Name{Local: "uri"}) != nil {
		return "", refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram graphic data has unmodeled type metadata")
	}
	uri, ok := exactNativeAttr(data, "", "uri")
	if !ok || uri != nativeDiagramURI(dialect) {
		return "", refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram graphic data is not the exact DrawingML diagram URI")
	}
	if requireOnlyNativeChildren(data, xml.Name{Space: uri, Local: "relIds"}) != nil {
		return "", refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram graphic data contains unknown markup")
	}
	relIds, err := nativeSingleton(data, uri, "relIds", true)
	if err != nil {
		return "", err
	}
	if requireOnlyNativeAttrs(relIds,
		xml.Name{Space: dialect.rels, Local: "dm"}, xml.Name{Space: dialect.rels, Local: "lo"},
		xml.Name{Space: dialect.rels, Local: "qs"}, xml.Name{Space: dialect.rels, Local: "cs"}) != nil || requireOnlyNativeChildren(relIds) != nil {
		return "", refuseNativeDiagram("pptx.diagram-markup-unavailable", "diagram relationship markup is outside the exact native subset")
	}
	id, ok := exactNativeAttr(relIds, dialect.rels, "dm")
	if !ok || id == "" || !nativeIDPattern.MatchString(id) {
		return "", refuseNativeDiagram("pptx.diagram-data-unavailable", "diagram requires one valid data-model relationship")
	}
	return id, nil
}

func nativeUniqueRelationship(relationships []nativeExtractRelationship, id string) (*nativeExtractRelationship, error) {
	var selected *nativeExtractRelationship
	for index := range relationships {
		if relationships[index].ID != id {
			continue
		}
		if selected != nil {
			return nil, fmt.Errorf("pptxpatch: native extract OPC: ambiguous duplicate diagram relationship id")
		}
		selected = &relationships[index]
	}
	return selected, nil
}

// resolveNativeDiagramDrawingRelationship finds the slide relationship of the
// PowerPoint drawing fallback. The data model names it through
// dsp:dataModelExt/@relId; without that extension exactly one diagramDrawing
// relationship on the slide is accepted, anything else is refused.
func (extractor *nativeExtractor) resolveNativeDiagramDrawingRelationship(relationships []nativeExtractRelationship, dataRelationshipID string, dialect nativeExtractDialect) (nativeExtractRelationship, error) {
	dataRel, err := nativeUniqueRelationship(relationships, dataRelationshipID)
	if err != nil {
		return nativeExtractRelationship{}, err
	}
	if dataRel == nil || dataRel.Type != dialect.rels+"/diagramData" || !dataRel.internal() || dataRel.Part == "" {
		return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-data-unavailable", "diagram data relationship is missing, external, or not the exact diagram data type")
	}
	if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(dataRel.Part), contentTypeDiagramData) {
		return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-data-unavailable", "diagram data part has an unexpected content type")
	}
	dataPayload := extractor.pkg.parts[dataRel.Part]
	if len(dataPayload) == 0 {
		return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-data-unavailable", "diagram data part is missing or empty")
	}
	dataRoot, err := parseNativeXML(dataPayload, dataRel.Part)
	if err != nil || dataRoot.Name != (xml.Name{Space: nativeDiagramURI(dialect), Local: "dataModel"}) {
		return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-data-unavailable", "diagram data part is not a well-formed dgm:dataModel")
	}
	drawingRelationshipID := nativeDiagramDrawingRelationshipID(dataRoot, dialect)
	var selected *nativeExtractRelationship
	if drawingRelationshipID != "" {
		selected, err = nativeUniqueRelationship(relationships, drawingRelationshipID)
		if err != nil {
			return nativeExtractRelationship{}, err
		}
		if selected == nil || selected.Type != relDiagramDrawing {
			return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-drawing-unavailable", "diagram data names a drawing relationship that is missing or not a diagram drawing; diagram layout is not evaluated")
		}
	} else {
		for index := range relationships {
			if relationships[index].Type != relDiagramDrawing {
				continue
			}
			if selected != nil {
				return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-drawing-unavailable", "diagram has no unambiguous pre-laid-out drawing part; diagram layout is not evaluated")
			}
			selected = &relationships[index]
		}
		if selected == nil {
			return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-drawing-unavailable", "diagram has no pre-laid-out drawing part; diagram layout is not evaluated")
		}
	}
	if !selected.internal() || selected.Part == "" {
		return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-drawing-unavailable", "diagram drawing relationship is external")
	}
	if !asciiEqualFoldNative(extractor.pkg.contentTypes.forPart(selected.Part), contentTypeDiagramDrawing) {
		return nativeExtractRelationship{}, refuseNativeDiagram("pptx.diagram-drawing-unavailable", "diagram drawing part has an unexpected content type")
	}
	return *selected, nil
}

func nativeDiagramDrawingRelationshipID(dataRoot *nativeXMLNode, dialect nativeExtractDialect) string {
	diagramURI := nativeDiagramURI(dialect)
	for _, extLst := range nativeChildren(dataRoot, diagramURI, "extLst") {
		for _, ext := range nativeChildren(extLst, dialect.drawing, "ext") {
			if uri, ok := exactNativeAttr(ext, "", "uri"); !ok || uri != nsDiagramDrawing {
				continue
			}
			for _, modelExt := range nativeChildren(ext, nsDiagramDrawing, "dataModelExt") {
				if relID, ok := exactNativeAttr(modelExt, "", "relId"); ok && relID != "" && nativeIDPattern.MatchString(relID) {
					return relID
				}
			}
		}
	}
	return ""
}

func (extractor *nativeExtractor) parseNativeDiagramDrawingShapes(part string) ([]*nativeXMLNode, []byte, error) {
	payload := extractor.pkg.parts[part]
	if len(payload) == 0 {
		return nil, nil, refuseNativeDiagram("pptx.diagram-drawing-unavailable", "diagram drawing part is missing or empty")
	}
	root, err := parseNativeXML(payload, part)
	if err != nil || root.Name != (xml.Name{Space: nsDiagramDrawing, Local: "drawing"}) {
		return nil, nil, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing part is not a well-formed dsp:drawing")
	}
	if requireOnlyNativeAttrs(root) != nil || requireOnlyNativeChildren(root, xml.Name{Space: nsDiagramDrawing, Local: "spTree"}) != nil {
		return nil, nil, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing root contains unknown markup")
	}
	spTree, err := nativeSingleton(root, nsDiagramDrawing, "spTree", true)
	if err != nil {
		return nil, nil, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing requires one shape tree")
	}
	if requireOnlyNativeAttrs(spTree) != nil || requireOnlyNativeChildren(spTree,
		xml.Name{Space: nsDiagramDrawing, Local: "nvGrpSpPr"},
		xml.Name{Space: nsDiagramDrawing, Local: "grpSpPr"},
		xml.Name{Space: nsDiagramDrawing, Local: "sp"}) != nil {
		return nil, nil, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing contains nested groups or unknown markup")
	}
	shapes := nativeChildren(spTree, nsDiagramDrawing, "sp")
	if len(shapes) == 0 {
		return nil, nil, refuseNativeDiagram(nativeDiagramDrawingEmptyCode, "diagram drawing part has no pre-laid-out shapes; diagram layout is not evaluated")
	}
	if len(shapes) > nativeMaxDiagramDrawingShapes {
		return nil, nil, refuseNativeDiagram("pptx.diagram-drawing-budget-unavailable", "diagram drawing exceeds the bounded shape budget")
	}
	return shapes, payload, nil
}

// extractNativeDiagramDrawingShape projects one dsp:sp. Its xfrm is already
// expressed in the graphic frame's coordinate space, so the child transform
// is used verbatim under the group's identity child mapping.
func (extractor *nativeExtractor) extractNativeDiagramDrawingShape(node *nativeXMLNode, index int, drawingPart string, drawingPayload []byte, slidePart, slideID, frameObjectID string, dialect nativeExtractDialect) (NativeElement, error) {
	dsp := func(local string) xml.Name { return xml.Name{Space: nsDiagramDrawing, Local: local} }
	if requireOnlyNativeAttrs(node, xml.Name{Local: "modelId"}) != nil || requireOnlyNativeChildren(node, dsp("nvSpPr"), dsp("spPr"), dsp("style"), dsp("txBody"), dsp("txXfrm")) != nil {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing shape contains unknown markup")
	}
	singletons := map[string]*nativeXMLNode{}
	for _, local := range []string{"nvSpPr", "spPr", "style", "txBody", "txXfrm"} {
		child, err := nativeSingleton(node, nsDiagramDrawing, local, local == "nvSpPr" || local == "spPr")
		if err != nil {
			return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing shape has duplicate or missing "+local)
		}
		singletons[local] = child
	}
	cNvPr, err := nativeSingleton(singletons["nvSpPr"], nsDiagramDrawing, "cNvPr", true)
	if err != nil || requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}, xml.Name{Local: "descr"}, xml.Name{Local: "title"}) != nil {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "hidden, linked, or unknown diagram drawing shape identity is not modeled")
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		name = ""
	}

	spPr := singletons["spPr"]
	allowed := []xml.Name{
		{Space: dialect.drawing, Local: "xfrm"}, {Space: dialect.drawing, Local: "prstGeom"}, {Space: dialect.drawing, Local: "custGeom"},
		{Space: dialect.drawing, Local: "noFill"}, {Space: dialect.drawing, Local: "solidFill"}, {Space: dialect.drawing, Local: "gradFill"},
		{Space: dialect.drawing, Local: "pattFill"}, {Space: dialect.drawing, Local: "blipFill"}, {Space: dialect.drawing, Local: "grpFill"},
		{Space: dialect.drawing, Local: "ln"}, {Space: dialect.drawing, Local: "effectLst"}, {Space: dialect.drawing, Local: "effectDag"},
		{Space: dialect.drawing, Local: "scene3d"}, {Space: dialect.drawing, Local: "sp3d"}, {Space: dialect.drawing, Local: "extLst"},
	}
	if requireOnlyNativeAttrs(spPr) != nil || requireOnlyNativeChildren(spPr, allowed...) != nil {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing shape properties contain unknown markup")
	}
	for _, allowedName := range allowed {
		if _, err := nativeSingleton(spPr, allowedName.Space, allowedName.Local, false); err != nil {
			return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-markup-unavailable", "diagram drawing shape properties contain duplicate markup")
		}
	}
	for _, local := range []string{"scene3d", "sp3d", "effectDag"} {
		if nativeChild(spPr, dialect.drawing, local) != nil {
			return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-effects-unavailable", "diagram drawing 3D or effect-graph markup is not approximated")
		}
	}
	xfrm := nativeChild(spPr, dialect.drawing, "xfrm")
	if xfrm == nil {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-transform-unavailable", "diagram drawing shape lacks an explicit transform")
	}
	transformGaps := nativeShapeGapSet{}
	transform, err := validateNativeAutoShapeTransform(xfrm, dialect, &transformGaps)
	if err != nil || transformGaps.refused() {
		return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-transform-unavailable", "diagram drawing shape transform is outside the exact native subset")
	}
	geometry, err := nativeDiagramDrawingGeometry(spPr, transform, dialect)
	if err != nil {
		return NativeElement{}, err
	}
	// The style matrix supplies fill/outline only where dsp:spPr omits them.
	// When it cannot be resolved exactly, the shape may still paint from its
	// own explicit fill AND outline; otherwise the frame is refused rather
	// than painting an outline-less or fill-less shape (mirrors AutoShape
	// pptx.autoshape-theme-style-unavailable).
	paint := spPr
	if style := singletons["style"]; style != nil {
		resolved, styleErr := resolveNativeShapeStyle(spPr, style, extractor.slideDependencies.themeRoot, dialect, extractor.theme)
		if styleErr == nil {
			paint = resolved
		} else if !nativeDiagramDrawingHasExplicitPaint(spPr, dialect) {
			return NativeElement{}, refuseNativeDiagram("pptx.diagram-drawing-style-unavailable", "diagram drawing shape inherits fill or outline from an unresolvable style matrix reference")
		}
	}
	fill, err := nativeDiagramDrawingFill(paint, dialect, extractor.theme)
	if err != nil {
		return NativeElement{}, err
	}
	stroke, err := nativeDiagramDrawingLine(paint, dialect, extractor.theme)
	if err != nil {
		return NativeElement{}, err
	}

	paragraphs := []NativeParagraph{}
	var layout *NativeTextBodyLayout
	textOmitted := ""
	inheritedText := false
	if txBody := singletons["txBody"]; txBody != nil {
		if reason := nativeDiagramDrawingTextFrameMismatch(singletons["txXfrm"], transform, geometry, dialect); reason != "" {
			textOmitted = reason
		} else {
			parsed, parsedLayout, omitted, inherited, textErr := extractor.extractNativeDiagramDrawingText(txBody, singletons["style"], transform, dialect)
			if textErr != nil {
				return NativeElement{}, textErr
			}
			if omitted != "" {
				textOmitted = omitted
			} else {
				paragraphs = parsed
				layout = parsedLayout
				inheritedText = inherited
			}
		}
	}

	raw, err := rawNativeNode(drawingPayload, node)
	if err != nil {
		return NativeElement{}, err
	}
	if extractor.elementsEmitted >= nativeMaxTotalElements || extractor.outputNodesEmitted >= nativeMaxNodes {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: element/output node budget exceeded")
	}
	extractor.elementsEmitted++
	extractor.outputNodesEmitted++
	objectID := frameObjectID + "/dsp/" + strconv.Itoa(index+1)
	elementID := extractor.identities.elements[nativeIdentityKey(slidePart, objectID)]
	if elementID == "" {
		elementID = stableNativeID("element", extractor.documentID+"\x00"+slidePart, objectID)
	}
	partName := drawingPart
	element := NativeElement{
		Kind: NativeElementKindShape, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, Geometry: geometry, Fill: fill, Stroke: stroke, Paragraphs: &paragraphs, TextBody: layout,
		Passthrough: []NativePassthroughRef{}, Children: nil,
		Source: &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: nativeSHA256(raw)},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusPreserveOnly, Diagnostics: []NativeDiagnostic{{
			Severity: NativeDiagnosticSeverityWarning, Code: nativeDiagramDrawingPreviewCode,
			Message: "diagram shape painted verbatim from the pre-laid-out drawing part (" + nativeDiagramDrawingPolicy + "); target remains read-only",
			Scope:   &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		}}},
	}
	// Rotated or flipped dsp:sp carry the same affine preview declaration
	// every other native producer emits (quarter-turn or bounded affine).
	for _, gap := range transformGaps.values {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning, Code: gap.code, Message: gap.message,
			Scope: &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		})
	}
	if textOmitted != "" {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning, Code: nativeDiagramDrawingTextOmittedCode,
			Message: "diagram shape geometry retained; text omitted: " + textOmitted,
			Scope:   &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		})
	}
	if inheritedText {
		nativeMarkInheritedTextPreview(&element)
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	return element, nil
}

// nativeDiagramDrawingHasExplicitPaint reports whether dsp:spPr declares both
// its own fill and its own outline, so an unresolvable style matrix cannot
// change what is painted.
func nativeDiagramDrawingHasExplicitPaint(spPr *nativeXMLNode, dialect nativeExtractDialect) bool {
	fill := false
	for _, local := range []string{"noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill"} {
		fill = fill || nativeChild(spPr, dialect.drawing, local) != nil
	}
	return fill && nativeChild(spPr, dialect.drawing, "ln") != nil
}

// nativeDiagramDrawingTextFrameMismatchToleranceEMU absorbs PowerPoint's own
// integer rounding of the preset text rectangle it writes into dsp:txXfrm.
const nativeDiagramDrawingTextFrameMismatchToleranceEMU = int64(2)

// nativeDiagramDrawingTextFrameMismatch returns a non-empty reason when
// dsp:txXfrm places the text somewhere other than the evaluated preset text
// rectangle of the shape (which is where the shape's text body is laid out).
// Without a separate text-frame model the text is omitted instead of being
// silently laid out in the wrong rectangle.
func nativeDiagramDrawingTextFrameMismatch(txXfrm *nativeXMLNode, transform NativeTransform, geometry *NativeEvaluatedGeometry, dialect nativeExtractDialect) string {
	if txXfrm == nil {
		return ""
	}
	if geometry == nil || requireOnlyNativeAttrs(txXfrm) != nil || requireOnlyNativeChildren(txXfrm, xml.Name{Space: dialect.drawing, Local: "off"}, xml.Name{Space: dialect.drawing, Local: "ext"}) != nil {
		return "diagram text frame (dsp:txXfrm) is rotated, flipped, or outside the exact subset"
	}
	off, offErr := nativeSingleton(txXfrm, dialect.drawing, "off", true)
	ext, extErr := nativeSingleton(txXfrm, dialect.drawing, "ext", true)
	if offErr != nil || extErr != nil || requireOnlyNativeAttrs(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}) != nil || requireOnlyNativeAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}) != nil {
		return "diagram text frame (dsp:txXfrm) is malformed"
	}
	values := map[string]int64{}
	for _, field := range []struct {
		node  *nativeXMLNode
		local string
		min   int64
	}{{off, "x", -nativeMaxSafeInteger}, {off, "y", -nativeMaxSafeInteger}, {ext, "cx", 0}, {ext, "cy", 0}} {
		raw, ok := exactNativeAttr(field.node, "", field.local)
		parsed, parseErr := parseCanonicalNativeInt(raw, field.min, nativeMaxSafeInteger)
		if !ok || parseErr != nil {
			return "diagram text frame (dsp:txXfrm) is non-canonical"
		}
		values[field.local] = parsed
	}
	expected := map[string]int64{
		"x": *transform.X + geometry.TextRect.X, "y": *transform.Y + geometry.TextRect.Y,
		"cx": geometry.TextRect.CX, "cy": geometry.TextRect.CY,
	}
	for key, want := range expected {
		delta := values[key] - want
		if delta < -nativeDiagramDrawingTextFrameMismatchToleranceEMU || delta > nativeDiagramDrawingTextFrameMismatchToleranceEMU {
			return "diagram text frame (dsp:txXfrm) differs from the preset text rectangle; a separate text frame is not modeled"
		}
	}
	return ""
}

func nativeDiagramDrawingGeometry(spPr *nativeXMLNode, transform NativeTransform, dialect nativeExtractDialect) (*NativeEvaluatedGeometry, error) {
	preset := nativeChild(spPr, dialect.drawing, "prstGeom")
	custom := nativeChild(spPr, dialect.drawing, "custGeom")
	var geometry *NativeEvaluatedGeometry
	var err error
	switch {
	case preset != nil && custom == nil:
		geometry, err = evaluateNativePresetSource(preset, dialect.drawing, *transform.Cx, *transform.Cy)
	case custom != nil && preset == nil:
		geometry, err = evaluateNativeCustomGeometry(custom, dialect.drawing, *transform.Cx, *transform.Cy)
	default:
		return nil, refuseNativeDiagram("pptx.diagram-drawing-geometry-unavailable", "diagram drawing shape requires exactly one preset or custom geometry")
	}
	if err != nil || geometry == nil {
		return nil, refuseNativeDiagram("pptx.diagram-drawing-geometry-unavailable", "diagram drawing shape geometry is outside the evaluated profile")
	}
	return geometry, nil
}

// nativeDiagramDrawingIdentityColor drops provably no-op color transforms
// that PowerPoint writes into drawing fallbacks (zero offsets, 100% mods) so
// the exact solid-color resolver sees the authored base color. Any other
// transform is left in place and must be handled or refused there.
func nativeDiagramDrawingIdentityColor(fill *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	if fill == nil {
		return nil
	}
	result := *fill
	result.Children = make([]*nativeXMLNode, 0, len(fill.Children))
	for _, color := range fill.Children {
		if color.Name != (xml.Name{Space: dialect.drawing, Local: "schemeClr"}) && color.Name != (xml.Name{Space: dialect.drawing, Local: "srgbClr"}) {
			result.Children = append(result.Children, color)
			continue
		}
		copied := *color
		copied.Children = make([]*nativeXMLNode, 0, len(color.Children))
		for _, transform := range color.Children {
			value, _ := exactNativeAttr(transform, "", "val")
			identity := false
			if transform.Name.Space == dialect.drawing && len(transform.Children) == 0 && len(transform.Attrs) == 1 {
				switch transform.Name.Local {
				case "hueOff", "satOff", "lumOff", "alphaOff":
					identity = value == "0"
				case "hueMod", "satMod", "lumMod", "alphaMod":
					identity = value == "100000"
				}
			}
			if !identity {
				copied.Children = append(copied.Children, transform)
			}
		}
		result.Children = append(result.Children, &copied)
	}
	return &result
}

func nativeDiagramDrawingFill(paint *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*string, error) {
	for _, local := range []string{"gradFill", "pattFill", "blipFill", "grpFill"} {
		if nativeChild(paint, dialect.drawing, local) != nil {
			return nil, refuseNativeDiagram("pptx.diagram-drawing-fill-unavailable", "diagram drawing gradient, pattern, picture, or group fills are not approximated")
		}
	}
	noFill := nativeChild(paint, dialect.drawing, "noFill")
	solidFill := nativeChild(paint, dialect.drawing, "solidFill")
	if (noFill == nil) == (solidFill == nil) {
		return nil, refuseNativeDiagram("pptx.diagram-drawing-fill-unavailable", "diagram drawing shape fill is inherited, missing, or conflicting")
	}
	if noFill != nil {
		if requireEmptyNativeElement(noFill) != nil {
			return nil, refuseNativeDiagram("pptx.diagram-drawing-fill-unavailable", "diagram drawing no-fill markup is not exact")
		}
		return nil, nil
	}
	color, err := exactNativeSolidColor(nativeDiagramDrawingIdentityColor(solidFill, dialect), dialect, theme)
	if err != nil {
		return nil, refuseNativeDiagram("pptx.diagram-drawing-fill-unavailable", "diagram drawing fill requires an exact sRGB or documented theme solid color")
	}
	return &color, nil
}

func nativeDiagramDrawingLine(paint *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*NativeStroke, error) {
	line := nativeChild(paint, dialect.drawing, "ln")
	if line == nil {
		return nil, nil
	}
	refuse := func(message string) (*NativeStroke, error) {
		return nil, refuseNativeDiagram("pptx.diagram-drawing-line-unavailable", message)
	}
	if requireOnlyNativeAttrs(line, xml.Name{Local: "w"}, xml.Name{Local: "cap"}, xml.Name{Local: "cmpd"}, xml.Name{Local: "algn"}) != nil {
		return refuse("diagram drawing outline attributes are outside the native subset")
	}
	children := []xml.Name{
		{Space: dialect.drawing, Local: "noFill"}, {Space: dialect.drawing, Local: "solidFill"}, {Space: dialect.drawing, Local: "prstDash"},
		{Space: dialect.drawing, Local: "round"}, {Space: dialect.drawing, Local: "bevel"}, {Space: dialect.drawing, Local: "miter"},
		{Space: dialect.drawing, Local: "headEnd"}, {Space: dialect.drawing, Local: "tailEnd"}, {Space: dialect.drawing, Local: "extLst"},
	}
	if requireOnlyNativeChildren(line, children...) != nil {
		return refuse("diagram drawing outline gradient, custom dash, or unknown markup is not approximated")
	}
	for _, child := range children {
		if _, err := nativeSingleton(line, child.Space, child.Local, false); err != nil {
			return refuse("diagram drawing outline contains duplicate markup")
		}
	}
	noFill := nativeChild(line, dialect.drawing, "noFill")
	solidFill := nativeChild(line, dialect.drawing, "solidFill")
	if (noFill == nil) == (solidFill == nil) {
		return refuse("diagram drawing outline fill is inherited, missing, or conflicting")
	}
	if noFill != nil {
		if requireEmptyNativeElement(noFill) != nil {
			return refuse("diagram drawing outline no-fill markup is not exact")
		}
		return nil, nil
	}
	if compound, ok := exactNativeAttr(line, "", "cmpd"); ok && compound != "sng" {
		return refuse("compound diagram drawing outlines are not approximated")
	}
	for _, local := range []string{"headEnd", "tailEnd"} {
		end := nativeChild(line, dialect.drawing, local)
		if end == nil {
			continue
		}
		if endType, ok := exactNativeAttr(end, "", "type"); ok && endType != "none" {
			return refuse("diagram drawing outline arrowheads are not modeled on shapes")
		}
	}
	widthValue, widthOK := exactNativeAttr(line, "", "w")
	width, widthErr := parseCanonicalNativeInt(widthValue, 0, nativeMaxLineWidthEmu)
	if !widthOK || widthErr != nil {
		return refuse("diagram drawing outline width is missing or non-canonical")
	}
	stroke := &NativeStroke{WidthEMU: int64Pointer(width)}
	// The contract carries cap/join/dash all-or-none. Only when the source
	// declares all three are they retained; otherwise the painter's existing
	// legacy-stroke defaults apply and nothing is invented here.
	var dash *NativeStrokeDash
	if dashNode := nativeChild(line, dialect.drawing, "prstDash"); dashNode != nil {
		dashValue, _ := exactNativeAttr(dashNode, "", "val")
		if requireOnlyNativeAttrs(dashNode, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(dashNode) != nil || dashValue != "solid" {
			return refuse("dashed diagram drawing outlines are not approximated")
		}
		solid := NativeStrokeDashSolid
		dash = &solid
	}
	var cap *NativeStrokeCap
	if capValue, ok := exactNativeAttr(line, "", "cap"); ok {
		var parsed NativeStrokeCap
		switch capValue {
		case "flat":
			parsed = NativeStrokeCapFlat
		case "rnd":
			parsed = NativeStrokeCapRound
		case "sq":
			parsed = NativeStrokeCapSquare
		default:
			return refuse("diagram drawing outline cap is not a native cap")
		}
		cap = &parsed
	}
	var join *NativeStrokeJoin
	var miterLimit *int64
	joins := 0
	if round := nativeChild(line, dialect.drawing, "round"); round != nil {
		if requireEmptyNativeElement(round) != nil {
			return refuse("diagram drawing round join is malformed")
		}
		parsed := NativeStrokeJoinRound
		join = &parsed
		joins++
	}
	if bevel := nativeChild(line, dialect.drawing, "bevel"); bevel != nil {
		if requireEmptyNativeElement(bevel) != nil {
			return refuse("diagram drawing bevel join is malformed")
		}
		parsed := NativeStrokeJoinBevel
		join = &parsed
		joins++
	}
	if miter := nativeChild(line, dialect.drawing, "miter"); miter != nil {
		if requireOnlyNativeAttrs(miter, xml.Name{Local: "lim"}) != nil || requireOnlyNativeChildren(miter) != nil {
			return refuse("diagram drawing miter join is malformed")
		}
		if limitValue, ok := exactNativeAttr(miter, "", "lim"); ok {
			limit, limitErr := parseCanonicalNativeInt(limitValue, 0, nativeMaxDrawingPercentage)
			if limitErr != nil {
				return refuse("diagram drawing miter limit is non-canonical")
			}
			miterLimit = int64Pointer(limit)
		}
		parsed := NativeStrokeJoinMiter
		join = &parsed
		joins++
	}
	if joins > 1 {
		return refuse("diagram drawing outline declares conflicting joins")
	}
	if cap != nil && join != nil && dash != nil {
		stroke.Cap, stroke.Join, stroke.Dash, stroke.MiterLimit = cap, join, dash, miterLimit
	}
	color, err := exactNativeSolidColor(nativeDiagramDrawingIdentityColor(solidFill, dialect), dialect, theme)
	if err != nil {
		return refuse("diagram drawing outline requires an exact sRGB or documented theme solid color")
	}
	stroke.Color = color
	return stroke, nil
}

// nativeDiagramDrawingTextBody copies a dsp:txBody, dropping only non-visual
// run metadata (proofing flags, disabled kerning) and provably no-op spacing
// (0 before/after, 100% line spacing) so the shared exact text extractor can
// qualify the remainder. Anything else stays and is refused there.
func nativeDiagramDrawingTextBody(txBody *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	var copyNode func(node *nativeXMLNode, depth int) *nativeXMLNode
	copyNode = func(node *nativeXMLNode, depth int) *nativeXMLNode {
		result := *node
		result.Attrs = make([]xml.Attr, 0, len(node.Attrs))
		for _, attr := range node.Attrs {
			if node.Name.Space == dialect.drawing && attr.Name.Space == "" {
				switch node.Name.Local {
				case "rPr", "endParaRPr":
					// Editor/proofing state has no paint. A kerning threshold does
					// (PowerPoint kerns pairs at or above it), so only the disabled
					// value kern="0" is a no-op; other thresholds stay and route the
					// text to the omitted path exactly like AutoShape runs.
					if attr.Name.Local == "dirty" || attr.Name.Local == "smtClean" || attr.Name.Local == "err" || attr.Name.Local == "noProof" || (attr.Name.Local == "kern" && attr.Value == "0") {
						continue
					}
				case "pPr":
					if attr.Name.Local == "defTabSz" {
						continue
					}
				}
			}
			result.Attrs = append(result.Attrs, attr)
		}
		result.Children = make([]*nativeXMLNode, 0, len(node.Children))
		for _, child := range node.Children {
			if node.Name == (xml.Name{Space: dialect.drawing, Local: "pPr"}) && nativeDiagramDrawingIdentitySpacing(child, dialect) {
				continue
			}
			if depth >= nativeMaxDepth {
				result.Children = append(result.Children, child)
				continue
			}
			result.Children = append(result.Children, copyNode(child, depth+1))
		}
		return &result
	}
	result := copyNode(txBody, 0)
	// Single-column bodies cannot show column spacing, and first/last
	// paragraph spacing is a no-op once no paragraph declares any spacing.
	if bodyPr := nativeChild(result, dialect.drawing, "bodyPr"); bodyPr != nil {
		columns, hasColumns := exactNativeAttr(bodyPr, "", "numCol")
		singleColumn := !hasColumns || columns == "1"
		spacing := false
		for _, paragraph := range nativeChildren(result, dialect.drawing, "p") {
			if pPr := nativeChild(paragraph, dialect.drawing, "pPr"); pPr != nil && (nativeChild(pPr, dialect.drawing, "spcBef") != nil || nativeChild(pPr, dialect.drawing, "spcAft") != nil) {
				spacing = true
			}
		}
		filtered := make([]xml.Attr, 0, len(bodyPr.Attrs))
		for _, attr := range bodyPr.Attrs {
			if attr.Name.Space == "" && ((attr.Name.Local == "spcCol" && singleColumn) || (attr.Name.Local == "spcFirstLastPara" && !spacing)) {
				continue
			}
			filtered = append(filtered, attr)
		}
		bodyPr.Attrs = filtered
	}
	return result
}

func nativeDiagramDrawingIdentitySpacing(node *nativeXMLNode, dialect nativeExtractDialect) bool {
	if node.Name.Space != dialect.drawing || len(node.Attrs) != 0 || len(node.Children) != 1 || !onlyNativeXMLSpace(node.Text) {
		return false
	}
	value := node.Children[0]
	if value.Name.Space != dialect.drawing || len(value.Children) != 0 || len(value.Attrs) != 1 {
		return false
	}
	amount, _ := exactNativeAttr(value, "", "val")
	switch node.Name.Local {
	case "spcBef", "spcAft":
		return (value.Name.Local == "spcPct" || value.Name.Local == "spcPts") && amount == "0"
	case "lnSpc":
		return value.Name.Local == "spcPct" && amount == "100000"
	}
	return false
}

// extractNativeDiagramDrawingText returns the qualified paragraphs and layout,
// or a non-empty omission reason when the text is outside the exact subset.
// With the opt-in inherited text preview the shared
// source-latin-inheritance-approximate-v1 resolver supplies the declared
// defaults, and inherited reports true so the element is labeled the same
// way as AutoShape text. Duplicate-singleton and budget errors are hard errors.
func (extractor *nativeExtractor) extractNativeDiagramDrawingText(txBody, style *nativeXMLNode, transform NativeTransform, dialect nativeExtractDialect) (paragraphs []NativeParagraph, layout *NativeTextBodyLayout, omitted string, inherited bool, err error) {
	body := nativeDiagramDrawingTextBody(txBody, dialect)
	if err := extractor.reserveNativeTextOutput(body, dialect); err != nil {
		return nil, nil, "", false, err
	}
	layout, err = extractNativeTextBodyLayoutPolicy(body, dialect, false)
	if err != nil {
		var duplicate nativeDuplicateSingletonError
		if isNativeDuplicateSingleton(err, &duplicate) {
			return nil, nil, "", false, err
		}
		return nil, nil, err.Error(), false, nil
	}
	if err := validateNativeTextBodyBounds(layout, transform); err != nil {
		return nil, nil, err.Error(), false, nil
	}
	paintText := body
	if extractor.options.AllowInheritedTextPreview {
		resolved, omissions, spacing, previewErr := extractor.inheritedTextPreview(body, style, true, dialect)
		if previewErr != nil {
			var duplicate nativeDuplicateSingletonError
			if isNativeDuplicateSingleton(previewErr, &duplicate) {
				return nil, nil, "", false, previewErr
			}
			return nil, nil, "diagram shape inherited text preview is unavailable: " + previewErr.Error(), false, nil
		}
		// The diagram fallback has no per-shape disclosure yet, so it cannot
		// carry the approximate paragraph spacing either: every authored slot
		// stays an omission and validated-but-unmodeled properties keep this
		// path fail-closed.
		for _, source := range spacing {
			nativeDiscloseParagraphSpacing(source, omissions)
		}
		if omitted := omissions.names(); len(omitted) != 0 {
			return nil, nil, "diagram shape inherited text omits unmodeled properties: " + strings.Join(omitted, ", "), false, nil
		}
		paintText = resolved
		inherited = true
	} else if style != nil {
		resolved, _, styleErr := resolveNativeShapeTextFontReference(body, style, dialect, extractor.theme)
		if styleErr != nil {
			return nil, nil, "diagram shape font reference is outside the exact subset", false, nil
		}
		paintText = resolved
	}
	paragraphs, err = extractor.extractNativeParagraphs(paintText, dialect)
	if err != nil {
		var duplicate nativeDuplicateSingletonError
		if isNativeDuplicateSingleton(err, &duplicate) {
			return nil, nil, "", false, err
		}
		return nil, nil, "diagram shape text content or formatting is outside the exact subset: " + err.Error(), false, nil
	}
	return paragraphs, layout, "", inherited, nil
}
