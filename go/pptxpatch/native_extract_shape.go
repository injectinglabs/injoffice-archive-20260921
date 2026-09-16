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
		placeholder, err := nativePlaceholderMetadata(node, dialect)
		if placeholder != nil && (err == nil || isNativePlaceholderUnsupported(err)) {
			return true, nil
		}
		return false, err
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
	paintProperties := shapeProperties
	var styleErr error
	if style != nil {
		var resolved *nativeXMLNode
		resolved, styleErr = resolveNativeShapeStyle(shapeProperties, style, extractor.slideDependencies.themeRoot, dialect, extractor.theme)
		if styleErr == nil {
			paintProperties = resolved
		}
	}
	transform, preset, geometry, fill, stroke, err := validateNativeAutoShapeProperties(paintProperties, dialect, extractor.theme, &gaps)
	if err != nil {
		return NativeElement{}, err
	}
	if style != nil {
		if styleErr != nil {
			gaps.add("pptx.autoshape-theme-style-unavailable", "theme and style-matrix references are preserved but not resolved by native PPTX v1", true)
		} else {
			gaps.add("pptx.autoshape-theme-style-preview", "solid theme fill/outline references are resolved from the source matrix; style-bound targets remain read-only", false)
		}
	}

	paragraphs := []NativeParagraph{}
	var textBodyLayout *NativeTextBodyLayout
	var authoredFit *nativeAuthoredAutoFit
	var inheritedOmissions *nativeInheritedTextOmissions
	paragraphSpacingApplied := false
	textOmitted := false
	if textBody != nil {
		if err := extractor.reserveNativeTextOutput(textBody, dialect); err != nil {
			return NativeElement{}, err
		}
		layout, fit, layoutErr := extractNativeTextBodyLayoutAuthored(textBody, dialect, extractor.options.AllowSourceFrameAutoFitPreview)
		if layoutErr != nil {
			var duplicate nativeDuplicateSingletonError
			if isNativeDuplicateSingleton(layoutErr, &duplicate) || !isNativeTextLayoutUnsupported(layoutErr) {
				return NativeElement{}, layoutErr
			}
			gaps.add("pptx.autoshape-text-layout-unavailable", "shape geometry retained; text omitted: "+layoutErr.Error(), false)
			textOmitted = true
		} else if boundsErr := validateNativeTextBodyBounds(layout, transform); boundsErr != nil {
			if !isNativeTextLayoutUnsupported(boundsErr) {
				return NativeElement{}, boundsErr
			}
			gaps.add("pptx.autoshape-text-layout-unavailable", "shape geometry retained; text omitted: "+boundsErr.Error(), false)
			textOmitted = true
		} else {
			textBodyLayout = layout
			authoredFit = fit
		}
		paintText := textBody
		fontReferenceUsed := false
		var parseErr error
		var inheritedSpacing []nativeParagraphSpacingSource
		if extractor.options.AllowInheritedTextPreview {
			placeholder, err := nativeTextPlaceholder(node, dialect)
			if err != nil {
				parseErr = err
			} else if placeholder != nil {
				parseErr = unsupportedNativeTextContent("placeholder inherited preview unsupported")
			} else if styleErr != nil {
				parseErr = styleErr
			} else {
				paintText, inheritedOmissions, inheritedSpacing, parseErr = extractor.inheritedTextPreview(textBody, style, true, dialect)
			}
		} else if style != nil && styleErr == nil {
			placeholder, placeholderErr := nativeTextPlaceholder(node, dialect)
			if placeholderErr != nil {
				parseErr = placeholderErr
			} else if placeholder == nil {
				paintText, fontReferenceUsed, parseErr = resolveNativeShapeTextFontReference(textBody, style, dialect, extractor.theme)
				if parseErr == nil && fontReferenceUsed && (extractor.fontReferenceBlocked || !nativeShapeReferenceEmptyLayer(extractor.slideDependencies.masterRoot, dialect.presentation, "txStyles", true, dialect)) {
					parseErr = unsupportedNativeTextContent("shape font reference competes with unqualified external text defaults")
				}
			}
		}
		var parsed []NativeParagraph
		presentationProjection := false
		if parseErr == nil {
			parsed, parseErr = extractor.extractNativeParagraphs(paintText, dialect)
			// Exact text stays exact. Text that is not self-contained may still
			// be completed by explicit presentation levels, as a read-only projection.
			if parseErr != nil && !extractor.options.AllowInheritedTextPreview && extractor.presentationTextPreviewStyle != nil {
				parsed, parseErr = extractor.extractNativeShapeParagraphs(paintText, dialect)
				presentationProjection = parseErr == nil
			}
		}
		if parseErr != nil {
			var duplicate nativeDuplicateSingletonError
			if isNativeDuplicateSingleton(parseErr, &duplicate) {
				return NativeElement{}, parseErr
			}
			gaps.add("pptx.autoshape-text-unavailable", "shape geometry retained; text omitted because its content or inherited formatting is unsupported: "+parseErr.Error(), false)
			textOmitted = true
		} else {
			paragraphs = parsed
			if fontReferenceUsed && !textOmitted {
				gaps.add("pptx.shape-font-reference-preview", "shape text font/color resolved from its authored theme reference; target remains read-only", false)
			}
			if presentationProjection {
				gaps.add(nativePresentationTextStylePreviewCode, nativePresentationTextStylePreviewMessage, false)
			}
		}
		if textOmitted {
			paragraphs = []NativeParagraph{}
			textBodyLayout = nil
			authoredFit = nil
			inheritedOmissions = nil
		} else {
			// Authored paragraph spacing resolves after the authored fontScale,
			// because a percentage gap measures the run sizes we will paint.
			nativeApplyAuthoredFontScale(paragraphs, authoredFit)
			paragraphSpacingApplied = nativeApplyParagraphSpacing(paragraphs, inheritedSpacing, inheritedOmissions)
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
		Transform: transform, Preset: preset, Geometry: geometry, Fill: fill, Stroke: stroke, Paragraphs: &paragraphs, TextBody: textBodyLayout,
		Passthrough: []NativePassthroughRef{}, Children: nil,
		Source:        &NativeSourceAnchor{PartName: slidePart, ObjectID: objectID, FingerprintSHA256: fingerprint},
		Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}},
	}
	if extractor.options.AllowInheritedTextPreview && textBody != nil {
		nativeMarkInheritedTextPreview(&element)
		if paragraphSpacingApplied {
			nativeMarkParagraphSpacing(&element)
		}
		nativeMarkInheritedTextOmissions(&element, inheritedOmissions)
	}
	if name != "" {
		element.Name = stringPointer(name)
	}
	if len(gaps.values) == 0 {
		nativeMarkVerticalTextPreview(&element)
		nativeMarkSourceFrameAutoFit(&element)
		nativeMarkAuthoredAutoFit(&element, authoredFit)
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
	nativeMarkSourceFrameAutoFit(&element)
	nativeMarkAuthoredAutoFit(&element, authoredFit)
	nativeMarkVerticalTextPreview(&element)
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

func validateNativeAutoShapeProperties(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme, gaps *nativeShapeGapSet) (NativeTransform, *NativeShapePreset, *NativeEvaluatedGeometry, *string, *NativeStroke, error) {
	// p:spPr/@bwMode (ECMA-376 Part 1 §19.3.1.44, ST_BlackWhiteMode) selects how
	// the shape is rendered when the application is displaying black and white.
	// "auto" and "clr" both keep the shape's own colors, which is what PowerPoint
	// paints in normal view, so neither changes a pixel; the background extractor
	// already accepts the same hint. Any restating mode keeps refusing.
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "bwMode"}); err != nil || !nativeNeutralBlackWhiteMode(node) {
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
			return NativeTransform{}, nil, nil, nil, nil, err
		}
	}
	xfrm, err := nativeSingleton(node, dialect.drawing, "xfrm", true)
	if err != nil {
		return NativeTransform{}, nil, nil, nil, nil, err
	}
	transform, err := validateNativeAutoShapeTransform(xfrm, dialect, gaps)
	if err != nil {
		return NativeTransform{}, nil, nil, nil, nil, err
	}
	var preset *NativeShapePreset
	var geometry *NativeEvaluatedGeometry
	custom := nativeChild(node, dialect.drawing, "custGeom")
	if custom != nil && nativeChild(node, dialect.drawing, "prstGeom") == nil {
		geometry, err = evaluateNativeCustomGeometry(custom, dialect.drawing, *transform.Cx, *transform.Cy)
		if err != nil {
			gaps.add("pptx.autoshape-geometry-unavailable", "custom geometry is outside the evaluated profile: "+err.Error(), true)
		} else {
			gaps.add("pptx.custom-geometry-preview", "DrawingML custom paths and text rectangle evaluated from source; geometry remains read-only", false)
		}
	} else {
		legacyGaps := nativeShapeGapSet{}
		preset = validateNativeAutoShapeGeometry(node, dialect, &legacyGaps)
		presetNode := nativeChild(node, dialect.drawing, "prstGeom")
		if preset != nil || presetNode == nil || custom != nil {
			for _, gap := range legacyGaps.values {
				gaps.add(gap.code, gap.message, gap.refusal)
			}
		} else {
			geometry, err = evaluateNativePresetSource(presetNode, dialect.drawing, *transform.Cx, *transform.Cy)
			if err != nil {
				gaps.add("pptx.autoshape-geometry-unavailable", "preset geometry is outside the evaluated profile: "+err.Error(), true)
			} else {
				gaps.add("pptx.preset-catalog-preview", "DrawingML preset catalog paths and adjustments evaluated from source; geometry remains read-only", false)
			}
		}
	}
	fill := validateNativeAutoShapeFill(node, dialect, theme, gaps)
	if geometry != nil && fill != nil {
		for _, path := range geometry.Paths {
			if path.FillMode != "norm" && path.FillMode != "none" {
				gaps.add("pptx.deterministic-path-tone-preview", "DrawingML shaded paths use linear-srgb-path-tone-20-40-v1; these relative-tone preview strengths are not qualified PowerPoint colors", false)
				break
			}
		}
	}
	stroke, err := validateNativeAutoShapeLine(node, dialect, theme, false, gaps)
	if err != nil {
		return NativeTransform{}, nil, nil, nil, nil, err
	}
	for _, name := range []string{"effectLst", "effectDag", "scene3d", "sp3d", "extLst"} {
		if child, _ := nativeSingleton(node, dialect.drawing, name, false); child != nil {
			gaps.add("pptx.autoshape-effects-unavailable", "shape effects, 3D, or extension markup is preserved but not approximated", true)
		}
	}
	return transform, preset, geometry, fill, stroke, nil
}

// nativeNeutralBlackWhiteMode reports whether a ST_BlackWhiteMode display hint
// leaves normal color rendering untouched. Only the two modes that mean "paint
// the object's own colors" qualify; gray, black, white, and hidden restate the
// paint and stay outside the exact subset.
func nativeNeutralBlackWhiteMode(node *nativeXMLNode) bool {
	value, ok := exactNativeAttr(node, "", "bwMode")
	if !ok {
		return true
	}
	return value == "auto" || value == "clr"
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
	cx, err := requiredCanonicalNativeShapeInt(ext, "cx", 0, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	cy, err := requiredCanonicalNativeShapeInt(ext, "cy", 0, nativeMaxSafeInteger)
	if err != nil {
		return NativeTransform{}, err
	}
	orientation, err := parseNativeSourceAffine(node)
	if err != nil {
		return NativeTransform{}, err
	}
	result := NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}
	// Keep the original exact-cardinal representation where it is sufficient.
	if !orientation.FlipH && !orientation.FlipV && orientation.Rotation%5400000 == 0 && (orientation.Rotation%10800000 == 0 || cx%2 == cy%2) {
		if q := orientation.Rotation / 5400000; q != 0 {
			result.QuarterTurns = int64Pointer(q)
			gaps.add("pptx.quarter-turn-preview", "source quarter-turn rotation is rendered with an exact integer affine; rotated targets remain read-only", false)
		}
	} else {
		result.RotationAngle = int64Pointer(orientation.Rotation)
		if orientation.FlipH {
			result.FlipH = &orientation.FlipH
		}
		if orientation.FlipV {
			result.FlipV = &orientation.FlipV
		}
		gaps.add("pptx.source-affine-preview", "DrawingML orientation uses bounded rational affine preview; transformed targets remain read-only", false)
	}
	return result, nil
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
	case "roundRect", "rightArrow", "hexagon":
		preset = NativeShapePreset(value)
		gaps.add("pptx.autoshape-preset-preview", "default "+value+" outline and text rectangle use DrawingML preset equations; preset target remains read-only", false)
	case "pentagon":
		preset = NativeShapePresetPentagon
		gaps.add("pptx.autoshape-preset-preview", "default pentagon outline is rendered from DrawingML preset equations; preset target remains read-only", false)
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
	lineChildren = append(lineChildren,
		xml.Name{Space: dialect.drawing, Local: "headEnd"},
		xml.Name{Space: dialect.drawing, Local: "tailEnd"},
	)
	if err := requireOnlyNativeChildren(line, lineChildren...); err != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline arrows, effects, or unknown markup are preserved but not approximated", true)
		return nil, nil
	}
	for _, name := range []string{"noFill", "solidFill", "prstDash", "round", "bevel", "miter", "headEnd", "tailEnd"} {
		if _, err := nativeSingleton(line, dialect.drawing, name, false); err != nil {
			return nil, err
		}
	}
	widthValue, widthOK := exactNativeAttr(line, "", "w")
	width, parseErr := parseCanonicalNativeInt(widthValue, 0, nativeMaxLineWidthEmu)
	if unpainted, err := nativeUnpaintedAutoShapeLine(line, dialect, widthOK, parseErr, gaps); unpainted || err != nil {
		return nil, err
	}
	if !allowLineEnds && !nativeUnarrowedAutoShapeLine(line, dialect) {
		gaps.add("pptx.autoshape-line-unavailable", "outline arrows, effects, or unknown markup are preserved but not approximated", true)
		return nil, nil
	}
	if !widthOK || parseErr != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline width is missing or non-canonical and is preserved without approximation", true)
		return nil, nil
	}
	// ECMA-376 Part 1 §20.1.2.1.24 leaves cap, cmpd, and algn optional on
	// a:ln. cmpd and algn carry the schema defaults "sng" and "ctr". cap has no
	// schema default, but PowerPoint writes cap only for "rnd" and "sq" and omits
	// it for flat ends, so the omitted cap is the flat cap PowerPoint paints.
	capValue, capOK := exactNativeAttr(line, "", "cap")
	if !capOK {
		capValue = "flat"
	}
	compound, compoundOK := exactNativeAttr(line, "", "cmpd")
	if !compoundOK {
		compound = "sng"
	}
	alignment, alignmentOK := exactNativeAttr(line, "", "algn")
	if !alignmentOK {
		alignment = "ctr"
	}
	var cap NativeStrokeCap
	capNative := true
	switch capValue {
	case "flat":
		cap = NativeStrokeCapFlat
	case "rnd":
		cap = NativeStrokeCapRound
	case "sq":
		cap = NativeStrokeCapSquare
	default:
		capNative = false
	}
	if !capNative || compound != "sng" || alignment != "ctr" {
		gaps.add("pptx.autoshape-line-unavailable", "only explicit single centered outlines with a native cap are representable", true)
		return nil, nil
	}
	noFill, _ := nativeSingleton(line, dialect.drawing, "noFill", false)
	solidFill, _ := nativeSingleton(line, dialect.drawing, "solidFill", false)
	if (noFill == nil) == (solidFill == nil) {
		gaps.add("pptx.autoshape-line-unavailable", "outline requires exactly one explicit no-fill or solid sRGB fill", true)
		return nil, nil
	}
	// The a:prstDash / a:custDash choice is optional on a:ln. An outline with
	// neither member is an unbroken line, the same stroke a "solid" preset dash
	// describes, so an absent dash paints identically to the explicit one.
	dashNode, _ := nativeSingleton(line, dialect.drawing, "prstDash", false)
	dashValue := "solid"
	if dashNode != nil {
		if requireOnlyNativeAttrs(dashNode, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(dashNode) != nil {
			gaps.add("pptx.autoshape-line-unavailable", "outline dash semantics are not explicit", true)
			return nil, nil
		}
		dashValue, _ = exactNativeAttr(dashNode, "", "val")
	}
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

// nativeUnarrowedAutoShapeLine reports whether an outline draws no line ends.
// a:headEnd / a:tailEnd default to type="none" (ECMA-376 Part 1 §20.1.2.1.4 and
// §20.1.2.1.5), so an absent, empty, or type="none" end paints nothing and the
// arrowless native AutoShape stroke is an exact description of the source. A
// named arrow type still refuses, because the AutoShape stroke cannot carry it.
func nativeUnarrowedAutoShapeLine(line *nativeXMLNode, dialect nativeExtractDialect) bool {
	for _, name := range []string{"headEnd", "tailEnd"} {
		end, _ := nativeSingleton(line, dialect.drawing, name, false)
		if end == nil {
			continue
		}
		arrow, err := exactNativeConnectorArrow(end)
		if err != nil || arrow {
			return false
		}
	}
	return true
}

// nativeUnpaintedAutoShapeLine recognizes an outline whose fill is explicitly
// a:noFill. Such an outline paints nothing regardless of width, cap, compound,
// alignment, dash or join, so those attributes are only required to be
// canonical when present. Any other structure keeps the strict validation.
func nativeUnpaintedAutoShapeLine(line *nativeXMLNode, dialect nativeExtractDialect, widthOK bool, widthErr error, gaps *nativeShapeGapSet) (bool, error) {
	noFill, _ := nativeSingleton(line, dialect.drawing, "noFill", false)
	solidFill, _ := nativeSingleton(line, dialect.drawing, "solidFill", false)
	if noFill == nil || solidFill != nil {
		return false, nil
	}
	if requireEmptyNativeElement(noFill) != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline no-fill markup is malformed", true)
		return true, nil
	}
	if widthOK && widthErr != nil {
		gaps.add("pptx.autoshape-line-unavailable", "outline width is non-canonical and is preserved without approximation", true)
		return true, nil
	}
	for _, check := range []struct {
		name   string
		values []string
	}{{"cap", []string{"flat", "rnd", "sq"}}, {"cmpd", []string{"sng", "dbl", "thickThin", "thinThick", "tri"}}, {"algn", []string{"ctr", "in"}}} {
		value, ok := exactNativeAttr(line, "", check.name)
		if !ok {
			continue
		}
		valid := false
		for _, candidate := range check.values {
			valid = valid || candidate == value
		}
		if !valid {
			gaps.add("pptx.autoshape-line-unavailable", "outline "+check.name+" is outside the DrawingML enumeration", true)
			return true, nil
		}
	}
	joins := 0
	for _, name := range []string{"round", "bevel", "miter"} {
		if child, _ := nativeSingleton(line, dialect.drawing, name, false); child != nil {
			joins++
		}
	}
	if joins > 1 {
		gaps.add("pptx.autoshape-line-unavailable", "outline declares conflicting joins", true)
		return true, nil
	}
	if dash, _ := nativeSingleton(line, dialect.drawing, "prstDash", false); dash != nil && (requireOnlyNativeAttrs(dash, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(dash) != nil) {
		gaps.add("pptx.autoshape-line-unavailable", "outline dash markup is malformed", true)
		return true, nil
	}
	return true, nil
}
