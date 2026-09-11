package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

type nativeConnectorGap struct {
	code    string
	message string
	refusal bool
}

type nativeConnectorGapSet struct {
	values []nativeConnectorGap
	seen   map[string]bool
}

func (gaps *nativeConnectorGapSet) add(code, message string, refusal bool) {
	if gaps.seen == nil {
		gaps.seen = map[string]bool{}
	}
	key := code + fmt.Sprint(refusal)
	if gaps.seen[key] {
		return
	}
	gaps.seen[key] = true
	gaps.values = append(gaps.values, nativeConnectorGap{code: code, message: message, refusal: refusal})
}

func (gaps nativeConnectorGapSet) refused() bool {
	for _, gap := range gaps.values {
		if gap.refusal {
			return true
		}
	}
	return false
}

// extractConnector projects only a straight, unrotated DrawingML connector
// whose line paint is explicit. Named endpoint arrows are currently presence
// flags plus additive typed source descriptors. A renderer must not treat a true flag
// as an exact triangle or infer an Office arrow size. Unknown named types,
// sizes, transformed/theme-unresolved paint, non-solid dash,
// custom/bent geometry, and unknown rendering markup remain an exact
// capability-backed refusal rather than a nearby line approximation.
func (extractor *nativeExtractor) extractConnector(node *nativeXMLNode, slidePart, slideID string, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "cxnSp"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid connector root")
	}
	if err := rejectNativeConnectorDialectMix(node, dialect); err != nil {
		return NativeElement{}, err
	}
	gaps := nativeConnectorGapSet{}
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.connector-markup-unavailable", "connector root attributes are outside the exact native subset", true)
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvCxnSpPr"},
		xml.Name{Space: dialect.presentation, Local: "spPr"},
		xml.Name{Space: dialect.presentation, Local: "style"}); err != nil {
		gaps.add("pptx.connector-markup-unavailable", "connector contains unmodeled rendering markup", true)
	}

	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvCxnSpPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	shapeProperties, err := nativeSingleton(node, dialect.presentation, "spPr", true)
	if err != nil {
		return NativeElement{}, err
	}
	style, err := nativeSingleton(node, dialect.presentation, "style", false)
	if err != nil {
		return NativeElement{}, err
	}

	objectID, name, err := validateNativeConnectorNonVisual(nonVisual, dialect, &gaps)
	if err != nil {
		return NativeElement{}, err
	}
	transform, antiDiagonal, stroke, headArrow, tailArrow, err := validateNativeConnectorProperties(shapeProperties, dialect, extractor.theme, &gaps)
	if err != nil {
		return NativeElement{}, err
	}
	if style != nil {
		gaps.add("pptx.connector-theme-style-unavailable", "connector theme and style-matrix references are preserved but not resolved", true)
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
	element := NativeElement{
		Kind: NativeElementKindConnector, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, Stroke: stroke, Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	if antiDiagonal {
		element.FlipH = boolPointer(true)
	}
	if headArrow {
		element.HeadArrow = boolPointer(true)
	}
	if tailArrow {
		element.TailArrow = boolPointer(true)
	}
	if !gaps.refused() {
		line, _ := nativeSingleton(shapeProperties, dialect.drawing, "ln", false)
		if line != nil {
			head, _ := nativeSingleton(line, dialect.drawing, "headEnd", false)
			tail, _ := nativeSingleton(line, dialect.drawing, "tailEnd", false)
			element.HeadEnd = nativeConnectorArrowDescriptor(head)
			element.TailEnd = nativeConnectorArrowDescriptor(tail)
		}
	}
	if len(gaps.values) == 0 {
		return element, nil
	}

	status := NativeCompatibilityStatusPreserveOnly
	reason := "pptx.connector-preserve-only"
	if gaps.refused() {
		status = NativeCompatibilityStatusRefused
		reason = "pptx.connector-refused"
	}
	passthrough, err := extractor.issuePassthrough(slidePart, objectID, fingerprint, raw, reason)
	if err != nil {
		return NativeElement{}, err
	}
	if err := extractor.reserveNativePassthroughReference(); err != nil {
		return NativeElement{}, err
	}
	element.Passthrough = append(element.Passthrough, passthrough)
	element.Compatibility.Status = status
	partName := slidePart
	for _, gap := range gaps.values {
		severity := NativeDiagnosticSeverityWarning
		if gap.refusal {
			severity = NativeDiagnosticSeverityRefusal
		}
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: severity, Code: gap.code, Message: gap.message,
			Scope: &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &partName},
		})
	}
	return element, nil
}

func validateNativeConnectorNonVisual(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeConnectorGapSet) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.connector-nonvisual-unavailable", "connector nonvisual properties contain unmodeled attributes", true)
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvCxnSpPr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		gaps.add("pptx.connector-nonvisual-unavailable", "connector nonvisual properties are outside the exact native subset", true)
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	cNvCxnSpPr, err := nativeSingleton(node, dialect.presentation, "cNvCxnSpPr", true)
	if err != nil {
		return "", "", err
	}
	nvPr, err := nativeSingleton(node, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", "", err
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil || len(cNvPr.Children) != 0 || !onlyNativeXMLSpace(cNvPr.Text) {
		gaps.add("pptx.connector-nonvisual-unavailable", "connector visibility, hyperlink, or extension metadata is not modeled", true)
	}
	nativeID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", fmt.Errorf("pptxpatch: native extract: connector name exceeds contract bound")
	}
	if err := requireEmptyNativeElement(cNvCxnSpPr); err != nil {
		gaps.add("pptx.connector-connection-unavailable", "connector endpoint attachment metadata is preserved but not editable", false)
	}
	if err := requireEmptyNativeElement(nvPr); err != nil {
		gaps.add("pptx.connector-inheritance-unavailable", "connector placeholder or inherited nonvisual properties are not resolved", true)
	}
	return "cNvPr-" + nativeID, name, nil
}

func validateNativeConnectorProperties(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme, gaps *nativeConnectorGapSet) (NativeTransform, bool, *NativeStroke, bool, bool, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.connector-properties-unavailable", "connector shape properties contain unmodeled attributes", true)
	}
	allowed := []xml.Name{
		{Space: dialect.drawing, Local: "xfrm"},
		{Space: dialect.drawing, Local: "prstGeom"},
		{Space: dialect.drawing, Local: "custGeom"},
		{Space: dialect.drawing, Local: "ln"},
		{Space: dialect.drawing, Local: "effectLst"},
		{Space: dialect.drawing, Local: "effectDag"},
		{Space: dialect.drawing, Local: "scene3d"},
		{Space: dialect.drawing, Local: "sp3d"},
		{Space: dialect.drawing, Local: "extLst"},
	}
	if err := requireOnlyNativeChildren(node, allowed...); err != nil {
		gaps.add("pptx.connector-properties-unavailable", "connector properties contain unknown rendering markup", true)
	}
	for _, name := range allowed {
		if _, err := nativeSingleton(node, name.Space, name.Local, false); err != nil {
			return NativeTransform{}, false, nil, false, false, err
		}
	}
	xfrm, err := nativeSingleton(node, dialect.drawing, "xfrm", true)
	if err != nil {
		return NativeTransform{}, false, nil, false, false, err
	}
	transform, antiDiagonal, err := validateNativeConnectorTransform(xfrm, dialect, gaps)
	if err != nil {
		return NativeTransform{}, false, nil, false, false, err
	}
	validateNativeConnectorGeometry(node, dialect, gaps)
	lineGaps := nativeShapeGapSet{}
	stroke, err := validateNativeAutoShapeLine(node, dialect, theme, true, &lineGaps)
	if err != nil {
		return NativeTransform{}, false, nil, false, false, err
	}
	for _, gap := range lineGaps.values {
		code := "pptx.connector-line-unavailable"
		if strings.Contains(gap.code, "dash") {
			code = "pptx.connector-dash-unavailable"
		}
		gaps.add(code, strings.ReplaceAll(gap.message, "outline", "connector line"), true)
	}
	headArrow, tailArrow := false, false
	if line, _ := nativeSingleton(node, dialect.drawing, "ln", false); line != nil {
		head, tail, arrowErr := validateNativeConnectorLineEnds(line, dialect, gaps)
		if arrowErr != nil {
			return NativeTransform{}, false, nil, false, false, arrowErr
		}
		headArrow, tailArrow = head, tail
	}
	for _, name := range []string{"effectLst", "effectDag", "scene3d", "sp3d", "extLst"} {
		if child, _ := nativeSingleton(node, dialect.drawing, name, false); child != nil {
			gaps.add("pptx.connector-effects-unavailable", "connector effects, 3D, or extension markup is preserved but not approximated", true)
		}
	}
	return transform, antiDiagonal, stroke, headArrow, tailArrow, nil
}

func validateNativeConnectorLineEnds(line *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeConnectorGapSet) (bool, bool, error) {
	head, err := nativeSingleton(line, dialect.drawing, "headEnd", false)
	if err != nil {
		return false, false, err
	}
	tail, err := nativeSingleton(line, dialect.drawing, "tailEnd", false)
	if err != nil {
		return false, false, err
	}
	headArrow, headErr := exactNativeConnectorArrow(head)
	if headErr != nil {
		gaps.add("pptx.connector-line-unavailable", "connector arrowhead size or type is outside the exact named subset", true)
		return false, false, nil
	}
	tailArrow, tailErr := exactNativeConnectorArrow(tail)
	if tailErr != nil {
		gaps.add("pptx.connector-line-unavailable", "connector arrowhead size or type is outside the exact named subset", true)
		return false, false, nil
	}
	return headArrow, tailArrow, nil
}

func exactNativeConnectorArrow(node *nativeXMLNode) (bool, error) {
	if node == nil {
		return false, nil
	}
	if requireOnlyNativeAttrs(node, xml.Name{Local: "type"}, xml.Name{Local: "w"}, xml.Name{Local: "len"}) != nil || requireOnlyNativeChildren(node) != nil {
		return false, fmt.Errorf("arrow markup is not exact")
	}
	if width, ok := exactNativeAttr(node, "", "w"); ok && !nativeConnectorArrowSizeIsExact(width) {
		return false, fmt.Errorf("arrow width does not map to an exact native EMU size")
	}
	if length, ok := exactNativeAttr(node, "", "len"); ok && !nativeConnectorArrowSizeIsExact(length) {
		return false, fmt.Errorf("arrow length does not map to an exact native EMU size")
	}
	value, ok := exactNativeAttr(node, "", "type")
	if !ok || value == "" || value == "none" {
		return false, nil
	}
	switch value {
	case "triangle", "arrow", "stealth", "diamond", "oval":
		return true, nil
	default:
		return false, fmt.Errorf("unsupported arrow type")
	}
}

func nativeConnectorArrowSizeIsExact(value string) bool {
	switch value {
	case "sm", "med", "lg":
		// Preserve named source dimensions; pixel sizes are renderer-policy-owned.
		return true
	default:
		return false
	}
}

func nativeConnectorArrowDescriptor(node *nativeXMLNode) *NativeArrowEnd {
	if node == nil {
		return nil
	}
	kind, _ := exactNativeAttr(node, "", "type")
	if kind == "" {
		kind = "none"
	}
	end := &NativeArrowEnd{Type: kind}
	if width, ok := exactNativeAttr(node, "", "w"); ok {
		end.W = stringPointer(width)
	}
	if length, ok := exactNativeAttr(node, "", "len"); ok {
		end.Len = stringPointer(length)
	}
	return end
}

func validateNativeConnectorTransform(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeConnectorGapSet) (NativeTransform, bool, error) {
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "rot"}, xml.Name{Local: "flipH"}, xml.Name{Local: "flipV"}); err != nil {
		gaps.add("pptx.connector-transform-unavailable", "connector transform contains unmodeled attributes", true)
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "off"},
		xml.Name{Space: dialect.drawing, Local: "ext"}); err != nil {
		return NativeTransform{}, false, fmt.Errorf("pptxpatch: native extract: invalid connector transform children: %w", err)
	}
	off, err := nativeSingleton(node, dialect.drawing, "off", true)
	if err != nil {
		return NativeTransform{}, false, err
	}
	ext, err := nativeSingleton(node, dialect.drawing, "ext", true)
	if err != nil {
		return NativeTransform{}, false, err
	}
	if err := requireOnlyNativeAttrs(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil || requireOnlyNativeChildren(off) != nil {
		return NativeTransform{}, false, fmt.Errorf("pptxpatch: native extract: invalid connector offset")
	}
	if err := requireOnlyNativeAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}); err != nil || requireOnlyNativeChildren(ext) != nil {
		return NativeTransform{}, false, fmt.Errorf("pptxpatch: native extract: invalid connector extent")
	}
	x, err := requiredCanonicalNativeConnectorInt(off, "x", -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, false, err
	}
	y, err := requiredCanonicalNativeConnectorInt(off, "y", -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, false, err
	}
	cx, err := requiredCanonicalNativeConnectorInt(ext, "cx", 1, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, false, err
	}
	cy, err := requiredCanonicalNativeConnectorInt(ext, "cy", 1, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, false, err
	}
	if value, ok := exactNativeAttr(node, "", "rot"); ok {
		rotation, parseErr := parseCanonicalNativeInt(value, -nativeMaxSafeInteger, nativeMaxSafeInteger)
		if parseErr != nil {
			return NativeTransform{}, false, fmt.Errorf("pptxpatch: native extract: invalid connector rotation")
		}
		if rotation != 0 {
			gaps.add("pptx.connector-transform-unavailable", "rotated connectors are preserved but not approximated", true)
		}
	}
	flipH, flipV := false, false
	for _, value := range []struct {
		name   string
		target *bool
	}{{name: "flipH", target: &flipH}, {name: "flipV", target: &flipV}} {
		if raw, ok := exactNativeAttr(node, "", value.name); ok {
			parsed, parseErr := nativeBool(raw)
			if parseErr != nil {
				return NativeTransform{}, false, parseErr
			}
			*value.target = parsed
		}
	}
	return NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}, flipH != flipV, nil
}

func requiredCanonicalNativeConnectorInt(node *nativeXMLNode, local string, minimum, maximum int64) (int64, error) {
	value, ok := exactNativeAttr(node, "", local)
	if !ok {
		return 0, fmt.Errorf("pptxpatch: native extract: missing connector %s", local)
	}
	parsed, err := parseCanonicalNativeInt(value, minimum, maximum)
	if err != nil {
		return 0, fmt.Errorf("pptxpatch: native extract: invalid connector %s", local)
	}
	return parsed, nil
}

func validateNativeConnectorGeometry(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeConnectorGapSet) {
	presetGeometry, _ := nativeSingleton(node, dialect.drawing, "prstGeom", false)
	customGeometry, _ := nativeSingleton(node, dialect.drawing, "custGeom", false)
	if presetGeometry == nil || customGeometry != nil {
		gaps.add("pptx.connector-geometry-unavailable", "custom or missing connector geometry is preserved but not approximated", true)
		return
	}
	if err := requireOnlyNativeAttrs(presetGeometry, xml.Name{Local: "prst"}); err != nil {
		gaps.add("pptx.connector-geometry-unavailable", "connector preset geometry attributes are outside the exact subset", true)
		return
	}
	adjustments, err := nativeSingleton(presetGeometry, dialect.drawing, "avLst", true)
	if err != nil || requireOnlyNativeChildren(presetGeometry, xml.Name{Space: dialect.drawing, Local: "avLst"}) != nil || requireEmptyNativeElement(adjustments) != nil {
		gaps.add("pptx.connector-geometry-unavailable", "connector geometry adjustments are preserved but not approximated", true)
		return
	}
	preset, ok := exactNativeAttr(presetGeometry, "", "prst")
	if !ok || (preset != "line" && preset != "straightConnector1") {
		gaps.add("pptx.connector-geometry-unavailable", "only straight line connector geometry is represented exactly", true)
	}
}

func rejectNativeConnectorDialectMix(node *nativeXMLNode, dialect nativeExtractDialect) error {
	opposite := map[string]bool{}
	if dialect.presentation == nsPresentationTransitional {
		opposite[nsPresentationStrict] = true
		opposite[nsDrawingStrict] = true
		opposite[nsOfficeRelsStrict] = true
	} else {
		opposite[nsPresentationTransitional] = true
		opposite[nsDrawingTransitional] = true
		opposite[nsOfficeRelsTransitional] = true
	}
	var walk func(*nativeXMLNode) error
	walk = func(current *nativeXMLNode) error {
		if opposite[current.Name.Space] {
			return fmt.Errorf("pptxpatch: native extract: connector mixes Strict and Transitional namespaces")
		}
		for _, attr := range current.Attrs {
			if attr.Name.Space != "xmlns" && attr.Name.Local != "xmlns" && opposite[attr.Name.Space] {
				return fmt.Errorf("pptxpatch: native extract: connector mixes Strict and Transitional namespaces")
			}
		}
		for _, child := range current.Children {
			if err := walk(child); err != nil {
				return err
			}
		}
		return nil
	}
	return walk(node)
}
