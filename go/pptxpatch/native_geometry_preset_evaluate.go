package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

func evaluateNativePresetGeometry(name string, adjustments map[string]int64, width, height int64) (*NativeEvaluatedGeometry, error) {
	node, err := prepareNativePresetGeometry(name, adjustments)
	if err != nil {
		return nil, err
	}
	return evaluateNativePreparedPresetGeometry(node, width, height)
}

func evaluateNativePreparedPresetGeometry(node *nativeXMLNode, width, height int64) (*NativeEvaluatedGeometry, error) {
	g, err := newNativeGeometryGuides(float64(width), float64(height))
	if err != nil {
		return nil, err
	}
	for _, listName := range []string{"avLst", "gdLst"} {
		list := nativeChild(node, nativePresetDrawingNS, listName)
		if list == nil {
			continue
		}
		guides := []nativeGeometryGuide{}
		for _, guide := range list.Children {
			name, _ := exactNativeAttr(guide, "", "name")
			formula, _ := exactNativeAttr(guide, "", "fmla")
			guides = append(guides, nativeGeometryGuide{name, formula})
		}
		if err := g.evaluateWithIntermediateLimit(guides, nativeGeometryMaxIntermediate); err != nil {
			return nil, err
		}
	}
	// Validated handles and connection sites do not paint a shape. Retain them in
	// the immutable original resource while evaluating only the 2D paint clauses.
	for _, name := range []string{"ahLst", "cxnLst"} {
		list := nativeChild(node, nativePresetDrawingNS, name)
		if list == nil {
			continue
		}
		if err := validateNativePresetMetadata(list, g); err != nil {
			return nil, err
		}
		list.Children = nil
	}
	return evaluateNativeGeometryWithIntermediateLimit(node, nativePresetDrawingNS, width, height, nativeGeometryMaxIntermediate)
}

func validateNativePresetMetadata(list *nativeXMLNode, g nativeGeometryGuides) error {
	ns := nativePresetDrawingNS
	if err := requireOnlyNativeAttrs(list); err != nil {
		return err
	}
	allowed := []xml.Name{{Space: ns, Local: "cxn"}}
	if list.Name.Local == "ahLst" {
		allowed = []xml.Name{{Space: ns, Local: "ahXY"}, {Space: ns, Local: "ahPolar"}}
	}
	if err := requireOnlyNativeChildren(list, allowed...); err != nil {
		return err
	}
	for _, item := range list.Children {
		keys := []string{"ang"}
		if item.Name.Local == "ahXY" {
			keys = []string{"gdRefX", "gdRefY", "minX", "maxX", "minY", "maxY"}
		}
		if item.Name.Local == "ahPolar" {
			keys = []string{"gdRefR", "gdRefAng", "minR", "maxR", "minAng", "maxAng"}
		}
		attrs := []xml.Name{}
		for _, key := range keys {
			attrs = append(attrs, xml.Name{Local: key})
		}
		if err := requireOnlyNativeAttrs(item, attrs...); err != nil {
			return err
		}
		if item.Name.Local == "cxn" {
			if _, ok := exactNativeAttr(item, "", "ang"); !ok {
				return fmt.Errorf("missing connection angle")
			}
		}
		for _, attr := range item.Attrs {
			if strings.HasPrefix(attr.Name.Local, "gdRef") {
				if _, ok := g.values[attr.Value]; !ok {
					return fmt.Errorf("unknown handle guide")
				}
			} else if attr.Name.Space == "" {
				if _, err := g.resolve(attr.Value); err != nil {
					return err
				}
			}
		}
		if err := requireOnlyNativeChildren(item, xml.Name{Space: ns, Local: "pos"}); err != nil {
			return err
		}
		if len(item.Children) != 1 {
			return fmt.Errorf("missing or repeated metadata position")
		}
		pos := item.Children[0]
		if err := requireOnlyNativeAttrs(pos, xml.Name{Local: "x"}, xml.Name{Local: "y"}); err != nil {
			return err
		}
		if err := requireOnlyNativeChildren(pos); err != nil {
			return err
		}
		for _, key := range []string{"x", "y"} {
			if _, err := nativeGeometryAttribute(g, pos, key); err != nil {
				return err
			}
		}
	}
	return nil
}
