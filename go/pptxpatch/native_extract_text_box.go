package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

// A transparent rectangular text box has the same paint geometry as the
// transform-only text boxes emitted by our authoring path. Keep all source
// markup intact; fills, outlines and non-rectangular shapes need their own model.
func validateNativeTextBoxShapeProperties(node *nativeXMLNode, dialect nativeExtractDialect) error {
	if err := requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "xfrm"},
		xml.Name{Space: dialect.drawing, Local: "prstGeom"},
		xml.Name{Space: dialect.drawing, Local: "noFill"},
		xml.Name{Space: dialect.drawing, Local: "ln"}); err != nil {
		return err
	}
	if !onlyNativeXMLSpace(node.Text) {
		return fmt.Errorf("pptxpatch: text box shape properties contain direct text")
	}
	geometry, err := nativeSingleton(node, dialect.drawing, "prstGeom", false)
	if err != nil {
		return err
	}
	fill, err := nativeSingleton(node, dialect.drawing, "noFill", false)
	if err != nil {
		return err
	}
	line, err := nativeSingleton(node, dialect.drawing, "ln", false)
	if err != nil {
		return err
	}
	if fill != nil && requireEmptyNativeElement(fill) != nil {
		return fmt.Errorf("pptxpatch: text box noFill has unsupported metadata")
	}
	if geometry != nil {
		preset, ok := exactNativeAttr(geometry, "", "prst")
		if !ok || preset != "rect" || fill == nil || !onlyNativeXMLSpace(geometry.Text) || requireOnlyNativeAttrs(geometry, xml.Name{Local: "prst"}) != nil || requireOnlyNativeChildren(geometry, xml.Name{Space: dialect.drawing, Local: "avLst"}) != nil {
			return fmt.Errorf("pptxpatch: text box requires a transparent rectangle without adjustments")
		}
		adjustments, err := nativeSingleton(geometry, dialect.drawing, "avLst", false)
		if err != nil {
			return err
		}
		if adjustments != nil && requireEmptyNativeElement(adjustments) != nil {
			return fmt.Errorf("pptxpatch: text box geometry adjustments are unsupported")
		}
	}
	if line != nil {
		if requireOnlyNativeAttrs(line) != nil || !onlyNativeXMLSpace(line.Text) || requireOnlyNativeChildren(line, xml.Name{Space: dialect.drawing, Local: "noFill"}) != nil {
			return fmt.Errorf("pptxpatch: text box outline is unsupported")
		}
		none, err := nativeSingleton(line, dialect.drawing, "noFill", true)
		if err != nil {
			return err
		}
		if requireEmptyNativeElement(none) != nil {
			return fmt.Errorf("pptxpatch: text box outline noFill has unsupported metadata")
		}
	}
	return nil
}
