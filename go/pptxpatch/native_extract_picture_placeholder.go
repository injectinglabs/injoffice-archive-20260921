package pptxpatch

import (
	"encoding/xml"
)

// nativePicturePlaceholderFrameCode discloses that a picture's box came from
// the placeholder chain rather than from the picture's own markup.
const nativePicturePlaceholderFrameCode = nativePlaceholderPreviewCode

// nativePicturePlaceholderIdentity reads the p:ph binding of a p:pic. The
// non-visual container of a picture is p:nvPicPr, so nativePlaceholderMetadata
// (which resolves p:nvSpPr) does not apply; both share the p:nvPr/p:ph shape.
func nativePicturePlaceholderIdentity(nonVisual *nativeXMLNode, dialect nativeExtractDialect) (*nativePlaceholderIdentity, error) {
	properties, err := nativeSingleton(nonVisual, dialect.presentation, "nvPr", true)
	if err != nil {
		return nil, err
	}
	return nativePlaceholderIdentityFromProperties(properties, dialect)
}

// projectNativePicturePlaceholderFrame proposes the box a picture placeholder
// inherits, as a private read-only projection of its own p:spPr.
//
// A picture that fills a layout placeholder authors an empty p:spPr and takes
// its box from the placeholder chain (ECMA-376 §19.3.1.36: a p:ph resolves
// slide -> layout -> master). Native PPTX v1 models no inherited geometry, so
// the strict projection refuses such a picture and the slide paints nothing.
//
// This follows the DOCX approximate image-extent pattern: the projection only
// proposes the inherited a:xfrm (and the inherited a:prstGeom when the picture
// declares no geometry of its own) and hands the result back to the UNMODIFIED
// strict qualifier, which keeps every rule it had. A picture whose inherited
// frame is outside that subset still refuses exactly as it did before, and the
// strict tier never calls this at all.
//
// It returns nil when the picture declares no placeholder, when no ancestor in
// the chain declares a box, or when the ancestor placeholder markup is outside
// the exact subset — each of which leaves the existing refusal in place.
func (extractor *nativeExtractor) projectNativePicturePlaceholderFrame(nonVisual, shapeProperties *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	identity, err := nativePicturePlaceholderIdentity(nonVisual, dialect)
	if identity == nil || err != nil {
		return nil
	}
	for _, root := range []*nativeXMLNode{extractor.slideDependencies.layoutRoot, extractor.slideDependencies.masterRoot} {
		if root == nil {
			continue
		}
		matched, _, matchErr := nativeMatchingPlaceholder(root, *identity, true, dialect)
		if matchErr != nil && identity.kind != "" {
			matched, _, matchErr = nativeMatchingPlaceholder(root, *identity, false, dialect)
		}
		if matchErr != nil {
			continue
		}
		ancestor, ancestorErr := nativeSingleton(matched, dialect.presentation, "spPr", false)
		if ancestorErr != nil || ancestor == nil {
			continue
		}
		transform := nativeChild(ancestor, dialect.drawing, "xfrm")
		if transform == nil {
			// This ancestor states no box; PowerPoint keeps walking the chain.
			continue
		}
		projected := *shapeProperties
		projected.Children = append([]*nativeXMLNode{transform}, shapeProperties.Children...)
		if nativeChild(shapeProperties, dialect.drawing, "prstGeom") == nil && nativeChild(shapeProperties, dialect.drawing, "custGeom") == nil {
			if geometry := nativeChild(ancestor, dialect.drawing, "prstGeom"); geometry != nil {
				projected.Children = append(projected.Children, geometry)
			}
		}
		return &projected
	}
	return nil
}

// nativePlaceholderIdentityFromProperties reads p:ph out of an already
// resolved p:nvPr, so p:sp and p:pic share one placeholder-identity contract.
func nativePlaceholderIdentityFromProperties(properties *nativeXMLNode, dialect nativeExtractDialect) (*nativePlaceholderIdentity, error) {
	ph, err := nativeSingleton(properties, dialect.presentation, "ph", false)
	if err != nil || ph == nil {
		return nil, err
	}
	identity := &nativePlaceholderIdentity{}
	identity.kind, _ = exactNativeAttr(ph, "", "type")
	if value, ok := exactNativeAttr(ph, "", "idx"); ok {
		identity.index, err = parseCanonicalNativeInt(value, 0, 4294967295)
		if err != nil {
			return identity, err
		}
	}
	// ST_PlaceholderSize/orient/hasCustomPrompt identify the slot; they do not
	// widen title/body inheritance. Unknown leftover markup stays unqualified.
	if requireOnlyNativeAttrs(properties) != nil || requireOnlyNativeChildren(properties, xml.Name{Space: dialect.presentation, Local: "ph"}) != nil || requireOnlyNativeAttrs(ph, xml.Name{Local: "type"}, xml.Name{Local: "idx"}, xml.Name{Local: "sz"}, xml.Name{Local: "orient"}, xml.Name{Local: "hasCustomPrompt"}) != nil || requireOnlyNativeChildren(ph) != nil {
		return identity, unsupportedNativePlaceholder("placeholder metadata is outside the exact subset")
	}
	return identity, nil
}
