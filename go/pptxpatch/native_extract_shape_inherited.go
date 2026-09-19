package pptxpatch

import "encoding/xml"

// Placeholder-inherited AutoShape frames.
//
// A slide placeholder declares only what it overrides: PowerPoint paints the
// rest — the frame rectangle, its geometry, its fill and its outline — from the
// matching layout placeholder and then the master one (ECMA-376 Part 1
// §19.3.1.36 p:ph and the placeholder inheritance it defines). The exact
// AutoShape projection refuses every one of those omissions, so a slide whose
// shapes are ordinary placeholders paints nothing at all: p:nvPr/p:ph alone
// raises pptx.autoshape-inheritance-unavailable, an absent a:prstGeom raises
// pptx.autoshape-geometry-unavailable, and an absent fill or a:ln raise the
// fill and line refusals. The element is then marked refused and rendered as an
// empty placeholder box.
//
// The read-only approximate tier (NativePPTXExtractOptions.AllowInheritedTextPreview)
// already owns validators for this chain — the ones resolveNativePlaceholderPreview
// uses for the text path. This file reuses them for the FRAME alone: master then
// layout, nearest wins, ancestors only, with the slide shape's own markup left to
// the AutoShape extractor so a local declaration always wins. The text body is
// not consulted, because a frame does not depend on one.
//
// Nothing here is guessed. When the chain is outside the validated subset, no
// frame is produced and every previous refusal stands.
const nativeInheritedShapeFrameCode = "pptx.autoshape-placeholder-inheritance-approximate"

const nativeInheritedShapeFrameMessage = "placeholder frame geometry, fill and outline are resolved from the layout/master chain as a read-only preview; the placeholder binding and anything the chain does not model stay preserved but not editable"

// nativeInheritedShapeFrame is what a placeholder's ancestry contributes. A nil
// frame means the chain did not qualify and nothing may be inherited.
type nativeInheritedShapeFrame struct {
	// transform is the nearest inherited a:xfrm, used only when the slide
	// placeholder declares none of its own.
	transform *nativeXMLNode
	fill      *string
	stroke    *NativeStroke
}

// inheritedAutoShapeFrame resolves what a slide placeholder inherits, or nil
// when the shape is not a placeholder or its ancestry is outside the preview.
func (extractor *nativeExtractor) inheritedAutoShapeFrame(node *nativeXMLNode, dialect nativeExtractDialect) *nativeInheritedShapeFrame {
	if node == nil {
		return nil
	}
	identity, err := nativePlaceholderMetadata(node, dialect)
	if err != nil || identity == nil {
		return nil
	}
	// Only the title/body families are resolved here. A frame family (pic,
	// clipArt) inherits its ancestor's real geometry, which must be evaluated
	// rather than flattened to a rectangle; that path already exists in
	// resolveNativePlaceholderPreview and keeps its own refusals.
	kind, family, ok := nativePlaceholderFamily(identity.kind)
	if !ok {
		return nil
	}
	layout, layoutIdentity, err := nativeMatchingPlaceholder(extractor.slideDependencies.layoutRoot, *identity, true, dialect)
	if err != nil && identity.kind != "" {
		layout, layoutIdentity, err = nativeMatchingPlaceholder(extractor.slideDependencies.layoutRoot, *identity, false, dialect)
	}
	if err != nil {
		return nil
	}
	if _, layoutFamily, ok := nativePlaceholderFamily(layoutIdentity.kind); !ok || layoutFamily != family {
		return nil
	}
	var master *nativeXMLNode
	if extractor.slideDependencies.masterRoot != nil {
		if matched, _, err := nativeMatchingPlaceholder(extractor.slideDependencies.masterRoot, nativePlaceholderIdentity{kind: family}, false, dialect); err == nil {
			master = matched
		}
	}
	preview := &nativePlaceholderPreview{kind: kind, sourceKind: identity.kind}
	var transform *nativeXMLNode
	geometryDeclared := false
	for _, shape := range []*nativeXMLNode{master, layout} {
		if shape == nil {
			continue
		}
		if requireOnlyNativeAttrs(shape) != nil || requireOnlyNativeChildren(shape,
			xml.Name{Space: dialect.presentation, Local: "nvSpPr"},
			xml.Name{Space: dialect.presentation, Local: "spPr"},
			xml.Name{Space: dialect.presentation, Local: "style"},
			xml.Name{Space: dialect.presentation, Local: "txBody"}) != nil || !onlyNativeXMLSpace(shape.Text) {
			return nil
		}
		if err := extractor.validateNativePlaceholderPreviewNonVisual(shape, true, dialect, preview); err != nil {
			return nil
		}
		properties, err := nativeSingleton(shape, dialect.presentation, "spPr", true)
		if err != nil {
			return nil
		}
		localTransform, err := extractor.validateNativePlaceholderPreviewShapeProperties(properties, true, dialect, preview)
		if err != nil {
			return nil
		}
		if localTransform != nil {
			transform = localTransform
		}
		if nativeChild(properties, dialect.drawing, "prstGeom") != nil || nativeChild(properties, dialect.drawing, "custGeom") != nil {
			geometryDeclared = true
		}
	}
	// An ancestor a:prstGeom is what validated as the unadjusted rectangle the
	// AutoShape path then paints. A chain that declares no geometry at all
	// inherits none, so the geometry refusal must stand.
	if !geometryDeclared {
		return nil
	}
	return &nativeInheritedShapeFrame{transform: transform, fill: preview.fill, stroke: preview.stroke}
}

// nativeShapeDeclaresFill reports whether a p:spPr states its own fill. An
// a:blipFill or a:grpFill counts: the shape overrides its ancestry there even
// though this tier cannot paint the result.
func nativeShapeDeclaresFill(node *nativeXMLNode, dialect nativeExtractDialect) bool {
	for _, name := range []string{"noFill", "solidFill", "gradFill", "pattFill", "blipFill", "grpFill"} {
		if nativeChild(node, dialect.drawing, name) != nil {
			return true
		}
	}
	return false
}
