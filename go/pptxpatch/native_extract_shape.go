package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

type nativeShapeGap struct {
	code    string
	message string
	refusal bool
}

type nativeShapeGapSet struct {
	values []nativeShapeGap
	seen   map[string]bool
}

func (gaps *nativeShapeGapSet) add(code, message string, refusal bool) {
	if gaps.seen == nil {
		gaps.seen = map[string]bool{}
	}
	key := code + fmt.Sprint(refusal)
	if gaps.seen[key] {
		return
	}
	gaps.seen[key] = true
	gaps.values = append(gaps.values, nativeShapeGap{code: code, message: message, refusal: refusal})
}

func (gaps nativeShapeGapSet) refused() bool {
	for _, gap := range gaps.values {
		if gap.refusal {
			return true
		}
	}
	return false
}

func nativeShapeIsTextBox(node *nativeXMLNode, dialect nativeExtractDialect) (bool, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "sp"}) {
		return false, fmt.Errorf("pptxpatch: native extract: invalid shape root")
	}
	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvSpPr", true)
	if err != nil {
		return false, err
	}
	nonVisualShape, err := nativeSingleton(nonVisual, dialect.presentation, "cNvSpPr", true)
	if err != nil {
		return false, err
	}
	value, ok := exactNativeAttr(nonVisualShape, "", "txBox")
	if !ok {
		placeholder, err := nativeTextPlaceholder(node, dialect)
		return placeholder != nil, err
	}
	parsed, err := nativeBool(value)
	if err != nil {
		return false, err
	}
	return parsed, nil
}

func (extractor *nativeExtractor) extractAutoShape(node *nativeXMLNode, slidePart, slideID string, dialect nativeExtractDialect) (NativeElement, error) {
	if node == nil || node.Name != (xml.Name{Space: dialect.presentation, Local: "sp"}) {
		return NativeElement{}, fmt.Errorf("pptxpatch: native extract: invalid AutoShape root")
	}
	if err := rejectNativeAutoShapeDialectMix(node, dialect); err != nil {
		return NativeElement{}, err
	}
	gaps := nativeShapeGapSet{}
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.autoshape-markup-unavailable", "shape root attributes are not representable in native PPTX v1", true)
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "nvSpPr"},
		xml.Name{Space: dialect.presentation, Local: "spPr"},
		xml.Name{Space: dialect.presentation, Local: "style"},
		xml.Name{Space: dialect.presentation, Local: "txBody"}); err != nil {
		gaps.add("pptx.autoshape-markup-unavailable", "shape contains unmodeled rendering markup", true)
	}

	nonVisual, err := nativeSingleton(node, dialect.presentation, "nvSpPr", true)
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
	textBody, err := nativeSingleton(node, dialect.presentation, "txBody", false)
	if err != nil {
		return NativeElement{}, err
	}

	objectID, name, err := validateNativeAutoShapeNonVisual(nonVisual, dialect, &gaps)
	if err != nil {
		return NativeElement{}, err
	}
	transform, preset, fill, stroke, err := validateNativeAutoShapeProperties(shapeProperties, dialect, extractor.theme, &gaps)
	if err != nil {
		return NativeElement{}, err
	}
	if style != nil {
		gaps.add("pptx.autoshape-theme-style-unavailable", "theme and style-matrix references are preserved but not resolved by native PPTX v1", true)
	}

	paragraphs := []NativeParagraph{}
	var textBodyLayout *NativeTextBodyLayout
	if textBody != nil {
		if err := extractor.reserveNativeTextOutput(textBody, dialect); err != nil {
			return NativeElement{}, err
		}
		layout, layoutErr := extractNativeTextBodyLayout(textBody, dialect)
		if layoutErr != nil {
			var duplicate nativeDuplicateSingletonError
			if isNativeDuplicateSingleton(layoutErr, &duplicate) || !isNativeTextLayoutUnsupported(layoutErr) {
				return NativeElement{}, layoutErr
			}
			gaps.add("pptx.autoshape-text-layout-unavailable", layoutErr.Error(), true)
		} else if boundsErr := validateNativeTextBodyBounds(layout, transform); boundsErr != nil {
			if !isNativeTextLayoutUnsupported(boundsErr) {
				return NativeElement{}, boundsErr
			}
			gaps.add("pptx.autoshape-text-layout-unavailable", boundsErr.Error(), true)
		} else {
			textBodyLayout = layout
		}
		parsed, parseErr := extractor.extractNativeParagraphs(textBody, dialect)
		if parseErr != nil {
			var duplicate nativeDuplicateSingletonError
			if isNativeDuplicateSingleton(parseErr, &duplicate) {
				return NativeElement{}, parseErr
			}
			gaps.add("pptx.autoshape-text-unavailable", "shape text or inherited text formatting is outside the native PPTX v1 subset", true)
		} else {
			paragraphs = parsed
		}
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
		Kind: NativeElementKindShape, ID: elementID, Provenance: NativeProvenanceParsed,
		Transform: transform, Preset: preset, Fill: fill, Stroke: stroke, Paragraphs: &paragraphs, TextBody: textBodyLayout,
		Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	if len(gaps.values) == 0 {
		nativePreserveTextCheckingMetadata(&element, node, dialect)
		return element, nil
	}

	status := NativeCompatibilityStatusPreserveOnly
	reason := "pptx.autoshape-preserve-only"
	if gaps.refused() {
		status = NativeCompatibilityStatusRefused
		reason = "pptx.autoshape-refused"
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

func rejectNativeAutoShapeDialectMix(node *nativeXMLNode, dialect nativeExtractDialect) error {
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
			return fmt.Errorf("pptxpatch: native extract: AutoShape mixes Strict and Transitional namespaces")
		}
		for _, attr := range current.Attrs {
			if attr.Name.Space != "xmlns" && attr.Name.Local != "xmlns" && opposite[attr.Name.Space] {
				return fmt.Errorf("pptxpatch: native extract: AutoShape mixes Strict and Transitional namespaces")
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

func isNativeDuplicateSingleton(err error, target *nativeDuplicateSingletonError) bool {
	if err == nil {
		return false
	}
	for current := err; current != nil; {
		if value, ok := current.(nativeDuplicateSingletonError); ok {
			*target = value
			return true
		}
		type unwrapper interface{ Unwrap() error }
		wrapped, ok := current.(unwrapper)
		if !ok {
			return false
		}
		current = wrapped.Unwrap()
	}
	return false
}

func validateNativeAutoShapeNonVisual(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeShapeGapSet) (string, string, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.autoshape-nonvisual-unavailable", "shape nonvisual metadata is not fully modeled in native PPTX v1", false)
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.presentation, Local: "cNvPr"},
		xml.Name{Space: dialect.presentation, Local: "cNvSpPr"},
		xml.Name{Space: dialect.presentation, Local: "nvPr"}); err != nil {
		gaps.add("pptx.autoshape-nonvisual-unavailable", "shape nonvisual metadata is not fully modeled in native PPTX v1", false)
	}
	cNvPr, err := nativeSingleton(node, dialect.presentation, "cNvPr", true)
	if err != nil {
		return "", "", err
	}
	cNvSpPr, err := nativeSingleton(node, dialect.presentation, "cNvSpPr", true)
	if err != nil {
		return "", "", err
	}
	nvPr, err := nativeSingleton(node, dialect.presentation, "nvPr", true)
	if err != nil {
		return "", "", err
	}
	if err := requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}); err != nil || len(cNvPr.Children) != 0 || !onlyNativeXMLSpace(cNvPr.Text) {
		gaps.add("pptx.autoshape-nonvisual-unavailable", "shape accessibility, hyperlink, or extension metadata is preserved but not editable", false)
	}
	nativeID, err := canonicalNativeUnsignedID(cNvPr, "", "id", 1)
	if err != nil {
		return "", "", err
	}
	name, _ := exactNativeAttr(cNvPr, "", "name")
	if utf16CodeUnitLengthBounded(name, 1025) > 1024 {
		return "", "", fmt.Errorf("pptxpatch: native extract: AutoShape name exceeds contract bound")
	}
	if err := requireOnlyNativeAttrs(cNvSpPr, xml.Name{Local: "txBox"}); err != nil || len(cNvSpPr.Children) != 0 || !onlyNativeXMLSpace(cNvSpPr.Text) {
		gaps.add("pptx.autoshape-locks-unavailable", "shape locks and nonvisual properties are preserved but not editable", false)
	}
	if value, ok := exactNativeAttr(cNvSpPr, "", "txBox"); ok {
		textBox, boolErr := nativeBool(value)
		if boolErr != nil {
			return "", "", boolErr
		}
		if textBox {
			return "", "", fmt.Errorf("pptxpatch: native extract: text box routed to AutoShape extractor")
		}
	}
	if err := requireEmptyNativeElement(nvPr); err != nil {
		gaps.add("pptx.autoshape-inheritance-unavailable", "shape placeholder or inherited nonvisual properties are not resolved by native PPTX v1", true)
	}
	return "cNvPr-" + nativeID, name, nil
}

func validateNativeAutoShapeProperties(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme, gaps *nativeShapeGapSet) (NativeTransform, *NativeShapePreset, *string, *NativeStroke, error) {
	if err := requireOnlyNativeAttrs(node); err != nil {
		gaps.add("pptx.autoshape-properties-unavailable", "shape property attributes are not modeled in native PPTX v1", true)
	}
	allowed := []xml.Name{
		{Space: dialect.drawing, Local: "xfrm"}, {Space: dialect.drawing, Local: "prstGeom"}, {Space: dialect.drawing, Local: "custGeom"},
		{Space: dialect.drawing, Local: "noFill"}, {Space: dialect.drawing, Local: "solidFill"}, {Space: dialect.drawing, Local: "gradFill"},
		{Space: dialect.drawing, Local: "pattFill"}, {Space: dialect.drawing, Local: "blipFill"}, {Space: dialect.drawing, Local: "grpFill"},
		{Space: dialect.drawing, Local: "ln"}, {Space: dialect.drawing, Local: "effectLst"}, {Space: dialect.drawing, Local: "effectDag"},
		{Space: dialect.drawing, Local: "scene3d"}, {Space: dialect.drawing, Local: "sp3d"}, {Space: dialect.drawing, Local: "extLst"},
	}
	if err := requireOnlyNativeChildren(node, allowed...); err != nil {
		gaps.add("pptx.autoshape-properties-unavailable", "shape properties contain unknown rendering markup", true)
	}
	for _, name := range allowed {
		if _, err := nativeSingleton(node, name.Space, name.Local, false); err != nil {
			return NativeTransform{}, nil, nil, nil, err
		}
	}
	xfrm, err := nativeSingleton(node, dialect.drawing, "xfrm", true)
	if err != nil {
		return NativeTransform{}, nil, nil, nil, err
	}
	transform, err := validateNativeAutoShapeTransform(xfrm, dialect, gaps)
	if err != nil {
		return NativeTransform{}, nil, nil, nil, err
	}
	preset := validateNativeAutoShapeGeometry(node, dialect, gaps)
	fill := validateNativeAutoShapeFill(node, dialect, theme, gaps)
	stroke, err := validateNativeAutoShapeLine(node, dialect, theme, false, gaps)
	if err != nil {
		return NativeTransform{}, nil, nil, nil, err
	}
	for _, name := range []string{"effectLst", "effectDag", "scene3d", "sp3d", "extLst"} {
		if child, _ := nativeSingleton(node, dialect.drawing, name, false); child != nil {
			gaps.add("pptx.autoshape-effects-unavailable", "shape effects, 3D, or extension markup is preserved but not approximated", true)
		}
	}
	return transform, preset, fill, stroke, nil
}

func validateNativeAutoShapeTransform(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeShapeGapSet) (NativeTransform, error) {
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "rot"}, xml.Name{Local: "flipH"}, xml.Name{Local: "flipV"}); err != nil {
		gaps.add("pptx.autoshape-transform-unavailable", "shape transform contains unmodeled attributes", true)
	}
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "off"},
		xml.Name{Space: dialect.drawing, Local: "ext"}); err != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid AutoShape transform children: %w", err)
	}
	off, err := nativeSingleton(node, dialect.drawing, "off", true)
	if err != nil {
		return NativeTransform{}, err
	}
	ext, err := nativeSingleton(node, dialect.drawing, "ext", true)
	if err != nil {
		return NativeTransform{}, err
	}
	if err := requireOnlyNativeAttrs(off, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil || requireOnlyNativeChildren(off) != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid AutoShape offset")
	}
	if err := requireOnlyNativeAttrs(ext, xml.Name{Local: "cx"}, xml.Name{Local: "cy"}); err != nil || requireOnlyNativeChildren(ext) != nil {
		return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid AutoShape extent")
	}
	x, err := requiredCanonicalNativeShapeInt(off, "x", -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	y, err := requiredCanonicalNativeShapeInt(off, "y", -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	cx, err := requiredCanonicalNativeShapeInt(ext, "cx", 1, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	cy, err := requiredCanonicalNativeShapeInt(ext, "cy", 1, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	if value, ok := exactNativeAttr(node, "", "rot"); ok {
		rotation, parseErr := parseCanonicalNativeInt(value, -nativeMaxSafeInteger, nativeMaxSafeInteger)
		if parseErr != nil {
			return NativeTransform{}, fmt.Errorf("pptxpatch: native extract: invalid AutoShape rotation")
		}
		if rotation != 0 {
			gaps.add("pptx.autoshape-transform-unavailable", "rotated shapes are preserved but not approximated by native PPTX v1", true)
		}
	}
	for _, name := range []string{"flipH", "flipV"} {
		if value, ok := exactNativeAttr(node, "", name); ok {
			flip, parseErr := nativeBool(value)
			if parseErr != nil {
				return NativeTransform{}, parseErr
			}
			if flip {
				gaps.add("pptx.autoshape-transform-unavailable", "flipped shapes are preserved but not approximated by native PPTX v1", true)
			}
		}
	}
	return NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}, nil
}

func requiredCanonicalNativeShapeInt(node *nativeXMLNode, local string, minimum, maximum int64) (int64, error) {
	value, ok := exactNativeAttr(node, "", local)
	if !ok {
		return 0, fmt.Errorf("pptxpatch: native extract: missing AutoShape %s", local)
	}
	parsed, err := parseCanonicalNativeInt(value, minimum, maximum)
	if err != nil {
		return 0, fmt.Errorf("pptxpatch: native extract: invalid AutoShape %s", local)
	}
	return parsed, nil
}

func validateNativeAutoShapeGeometry(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeShapeGapSet) *NativeShapePreset {
	presetGeometry, _ := nativeSingleton(node, dialect.drawing, "prstGeom", false)
	customGeometry, _ := nativeSingleton(node, dialect.drawing, "custGeom", false)
	if presetGeometry == nil || customGeometry != nil {
		gaps.add("pptx.autoshape-geometry-unavailable", "custom, inherited, or missing shape geometry is preserved but not approximated", true)
		return nil
	}
	if err := requireOnlyNativeAttrs(presetGeometry, xml.Name{Local: "prst"}); err != nil {
		gaps.add("pptx.autoshape-geometry-unavailable", "shape preset geometry attributes are outside the native v1 subset", true)
		return nil
	}
	adjustments, err := nativeSingleton(presetGeometry, dialect.drawing, "avLst", true)
	if err != nil || requireOnlyNativeChildren(presetGeometry, xml.Name{Space: dialect.drawing, Local: "avLst"}) != nil || requireEmptyNativeElement(adjustments) != nil {
		gaps.add("pptx.autoshape-adjustments-unavailable", "shape geometry adjustments are preserved but not approximated", true)
		return nil
	}
	value, ok := exactNativeAttr(presetGeometry, "", "prst")
	if !ok {
		gaps.add("pptx.autoshape-geometry-unavailable", "shape preset geometry is missing", true)
		return nil
	}
	var preset NativeShapePreset
	switch value {
	case "rect":
		preset = NativeShapePresetRect
	case "ellipse":
		preset = NativeShapePresetEllipse
	case "triangle":
		preset = NativeShapePresetTriangle
	case "diamond":
		preset = NativeShapePresetDiamond
	default:
		gaps.add("pptx.autoshape-preset-unavailable", "this PowerPoint preset is outside the exact native v1 AutoShape subset", true)
		return nil
	}
	return &preset
}

func validateNativeAutoShapeFill(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme, gaps *nativeShapeGapSet) *string {
	noFill, _ := nativeSingleton(node, dialect.drawing, "noFill", false)
	solidFill, _ := nativeSingleton(node, dialect.drawing, "solidFill", false)
	unsupported := false
	for _, name := range []string{"gradFill", "pattFill", "blipFill", "grpFill"} {
		child, _ := nativeSingleton(node, dialect.drawing, name, false)
		unsupported = unsupported || child != nil
	}
	if unsupported || (noFill == nil && solidFill == nil) || (noFill != nil && solidFill != nil) {
		gaps.add("pptx.autoshape-fill-unavailable", "gradient, pattern, picture, group, inherited, or conflicting fill is preserved but not approximated", true)
		return nil
	}
	if noFill != nil {
		if requireEmptyNativeElement(noFill) != nil {
			gaps.add("pptx.autoshape-fill-unavailable", "shape no-fill markup is not exact", true)
		}
		return nil
	}
	color, err := exactNativeSolidColor(solidFill, dialect, theme)
	if err != nil {
		gaps.add("pptx.autoshape-fill-unavailable", "only an exact sRGB or documented theme solid fill is representable; unmodeled color transforms are preserved", true)
		return nil
	}
	return &color
}

func exactNativeAutoShapeColor(node *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	return exactNativeSolidColor(node, dialect, nativeResolvedTheme{})
}

func validateNativeAutoShapeLine(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme, allowLineEnds bool, gaps *nativeShapeGapSet) (*NativeStroke, error) {
	line, err := nativeSingleton(node, dialect.drawing, "ln", false)
	if err != nil {
		return nil, err
	}
	if line == nil {
		gaps.add("pptx.autoshape-line-unavailable", "missing or inherited outline is preserved but not approximated", true)
		return nil, nil
	}
	if err := requireOnlyNativeAttrs(line,
		xml.Name{Local: "w"}, xml.Name{Local: "cap"}, xml.Name{Local: "cmpd"}, xml.Name{Local: "algn"}); err != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline attributes are outside the native v1 subset", true)
		return nil, nil
	}
	lineChildren := []xml.Name{
		{Space: dialect.drawing, Local: "noFill"},
		{Space: dialect.drawing, Local: "solidFill"},
		{Space: dialect.drawing, Local: "prstDash"},
		{Space: dialect.drawing, Local: "round"},
		{Space: dialect.drawing, Local: "bevel"},
		{Space: dialect.drawing, Local: "miter"},
	}
	if allowLineEnds {
		lineChildren = append(lineChildren,
			xml.Name{Space: dialect.drawing, Local: "headEnd"},
			xml.Name{Space: dialect.drawing, Local: "tailEnd"},
		)
	}
	if err := requireOnlyNativeChildren(line, lineChildren...); err != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline arrows, effects, or unknown markup are preserved but not approximated", true)
		return nil, nil
	}
	for _, name := range []string{"noFill", "solidFill", "prstDash", "round", "bevel", "miter", "headEnd", "tailEnd"} {
		if !allowLineEnds && (name == "headEnd" || name == "tailEnd") {
			continue
		}
		if _, err := nativeSingleton(line, dialect.drawing, name, false); err != nil {
			return nil, err
		}
	}
	widthValue, widthOK := exactNativeAttr(line, "", "w")
	width, parseErr := parseCanonicalNativeInt(widthValue, 0, nativeMaxLineWidthEmu)
	if !widthOK || parseErr != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline width is missing or non-canonical and is preserved without approximation", true)
		return nil, nil
	}
	capValue, capOK := exactNativeAttr(line, "", "cap")
	compound, compoundOK := exactNativeAttr(line, "", "cmpd")
	alignment, alignmentOK := exactNativeAttr(line, "", "algn")
	var cap NativeStrokeCap
	switch capValue {
	case "flat":
		cap = NativeStrokeCapFlat
	case "rnd":
		cap = NativeStrokeCapRound
	case "sq":
		cap = NativeStrokeCapSquare
	default:
		capOK = false
	}
	if !capOK || !compoundOK || compound != "sng" || !alignmentOK || alignment != "ctr" {
		gaps.add("pptx.autoshape-line-unavailable", "only explicit single centered outlines with a native cap are representable", true)
		return nil, nil
	}
	noFill, _ := nativeSingleton(line, dialect.drawing, "noFill", false)
	solidFill, _ := nativeSingleton(line, dialect.drawing, "solidFill", false)
	if (noFill == nil) == (solidFill == nil) {
		gaps.add("pptx.autoshape-line-unavailable", "outline requires exactly one explicit no-fill or solid sRGB fill", true)
		return nil, nil
	}
	dashNode, _ := nativeSingleton(line, dialect.drawing, "prstDash", false)
	if dashNode == nil || requireOnlyNativeAttrs(dashNode, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(dashNode) != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline dash semantics are not explicit", true)
		return nil, nil
	}
	dashValue, _ := exactNativeAttr(dashNode, "", "val")
	if dashValue != "solid" {
		gaps.add("pptx.autoshape-dash-unavailable", "dashed outlines are preserved but not approximated", true)
		return nil, nil
	}
	round, _ := nativeSingleton(line, dialect.drawing, "round", false)
	bevel, _ := nativeSingleton(line, dialect.drawing, "bevel", false)
	miter, _ := nativeSingleton(line, dialect.drawing, "miter", false)
	joinCount := 0
	for _, join := range []*nativeXMLNode{round, bevel, miter} {
		if join != nil {
			joinCount++
		}
	}
	if joinCount != 1 {
		gaps.add("pptx.autoshape-line-unavailable", "outline requires one explicit native join", true)
		return nil, nil
	}
	join := NativeStrokeJoinRound
	var miterLimit *int64
	if round != nil {
		if requireEmptyNativeElement(round) != nil {
			gaps.add("pptx.autoshape-line-unavailable", "round outline join is malformed", true)
			return nil, nil
		}
	} else if bevel != nil {
		join = NativeStrokeJoinBevel
		if requireEmptyNativeElement(bevel) != nil {
			gaps.add("pptx.autoshape-line-unavailable", "bevel outline join is malformed", true)
			return nil, nil
		}
	} else {
		join = NativeStrokeJoinMiter
		if requireOnlyNativeAttrs(miter, xml.Name{Local: "lim"}) != nil || requireOnlyNativeChildren(miter) != nil {
			gaps.add("pptx.autoshape-line-unavailable", "miter outline join is malformed", true)
			return nil, nil
		}
		limitValue, ok := exactNativeAttr(miter, "", "lim")
		limit, limitErr := parseCanonicalNativeInt(limitValue, 0, nativeMaxDrawingPercentage)
		if !ok || limitErr != nil {
			gaps.add("pptx.autoshape-line-unavailable", "outline miter limit is missing or non-canonical and is preserved without approximation", true)
			return nil, nil
		}
		miterLimit = int64Pointer(limit)
	}
	if noFill != nil {
		if requireEmptyNativeElement(noFill) != nil {
			gaps.add("pptx.autoshape-line-unavailable", "outline no-fill markup is malformed", true)
		}
		return nil, nil
	}
	color, colorErr := exactNativeSolidColor(solidFill, dialect, theme)
	if colorErr != nil {
		gaps.add("pptx.autoshape-line-unavailable", "only an exact sRGB or documented theme solid outline is representable; unmodeled color transforms are preserved", true)
		return nil, nil
	}
	dash := NativeStrokeDashSolid
	return &NativeStroke{Color: color, WidthEMU: int64Pointer(width), Cap: &cap, Join: &join, Dash: &dash, MiterLimit: miterLimit}, nil
}
