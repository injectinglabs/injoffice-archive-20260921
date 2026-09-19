package pptxpatch

import (
	"encoding/xml"
	"errors"
	"fmt"
)

// nsNativeMarkupCompatibility is the Markup Compatibility and Extensibility
// namespace (ECMA-376 Part 3). It is version-independent, so the strict and
// transitional dialects share it.
const nsNativeMarkupCompatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006"

// A slide's embedded controls live in p:cSld/p:controls, outside p:spTree
// (ECMA-376 §19.3.1.15). PowerPoint writes each one as an mc:AlternateContent
// whose mc:Choice requires the VML namespace and whose mc:Fallback carries an
// ordinary p:pic — a snapshot of the control with its own a:blipFill and
// a:xfrm. A consumer that does not understand the required namespace is
// REQUIRED by MCE to process the fallback, which is exactly this reader's
// position: native PPTX v1 models no ActiveX control.
//
// The control markup itself stays preserve-only passthrough; only the picture
// the fallback already states is projected, at the box the fallback states.
// Controls paint above p:spTree, so these elements are appended after it.
func (extractor *nativeExtractor) extractNativeControlPictures(controls *nativeXMLNode, part, slideID string, used map[string]bool, relationships []nativeExtractRelationship, dialect nativeExtractDialect) ([]NativeElement, error) {
	pictures := nativeControlFallbackPictures(controls, dialect)
	if len(pictures) == 0 {
		return nil, nil
	}
	if len(pictures) > nativeMaxTotalElements-extractor.elementsEmitted || len(pictures) > nativeMaxNodes-extractor.outputNodesEmitted {
		return nil, fmt.Errorf("pptxpatch: native extract: element/output node budget exceeded")
	}
	elements := make([]NativeElement, 0, len(pictures))
	for _, picture := range pictures {
		element, err := extractor.extractPicture(picture, part, slideID, relationships, dialect)
		if err != nil {
			// A fallback this reader cannot project leaves the control exactly
			// as it was before: preserved bytes and no element. The refusal is
			// this picture's limit, not the slide's, so the rest still paint.
			var refusal nativePictureProjectionRefusal
			if !errors.As(err, &refusal) {
				return nil, err
			}
			continue
		}
		if element.Source == nil || element.Source.RelationshipID == nil {
			continue
		}
		if used[*element.Source.RelationshipID] {
			// Two elements may not claim one image relationship; the p:spTree
			// picture that already claimed it keeps it.
			continue
		}
		used[*element.Source.RelationshipID] = true
		elements = append(elements, element)
	}
	extractor.elementsEmitted += len(elements)
	extractor.outputNodesEmitted += len(elements)
	return elements, nil
}

// nativeControlFallbackPictures returns the p:pic of every control that states
// one, in document order. A control with no fallback picture — a p:control
// that names only its own relationship — contributes nothing, and neither does
// mc:Choice content, which is the branch this reader declined.
func nativeControlFallbackPictures(controls *nativeXMLNode, dialect nativeExtractDialect) []*nativeXMLNode {
	var pictures []*nativeXMLNode
	appendControl := func(control *nativeXMLNode) {
		if picture := nativeChild(control, dialect.presentation, "pic"); picture != nil {
			pictures = append(pictures, picture)
		}
	}
	for _, child := range controls.Children {
		switch child.Name {
		case xml.Name{Space: dialect.presentation, Local: "control"}:
			appendControl(child)
		case xml.Name{Space: nsNativeMarkupCompatibility, Local: "AlternateContent"}:
			fallback := nativeChild(child, nsNativeMarkupCompatibility, "Fallback")
			if fallback == nil {
				continue
			}
			for _, control := range nativeChildren(fallback, dialect.presentation, "control") {
				appendControl(control)
			}
		}
	}
	return pictures
}
