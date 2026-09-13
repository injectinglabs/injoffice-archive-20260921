package pptxpatch

import (
	"encoding/xml"
	"strings"
)

// nativeLiteralPieColors resolves only local, complete source paint: an optional
// series solid RGB/no-line plus explicitly indexed complete point overrides.
// No chart/theme palette or partial shape-property inheritance is consulted.
func nativeLiteralPieColors(nodes []*nativeXMLNode, count int, d nativeExtractDialect) ([]string, bool) {
	if count < 1 || count > 64 || len(nodes) > 65 {
		return nil, false
	}
	colors := make([]string, count)
	if len(nodes) > 0 && nodes[0].Name == (xml.Name{Space: d.chart, Local: "spPr"}) {
		color, ok := nativeLiteralPieSolidPaint(nodes[0], d)
		if !ok {
			return nil, false
		}
		for i := range colors {
			colors[i] = color
		}
		nodes = nodes[1:]
	}
	if len(nodes) > count {
		return nil, false
	}
	seen := make([]bool, count)
	for _, point := range nodes {
		if point.Name != (xml.Name{Space: d.chart, Local: "dPt"}) || !nativeLiteralPiePaintSequence(point, d.chart, "idx", "spPr") {
			return nil, false
		}
		index := point.Children[0]
		if requireOnlyNativeAttrs(index, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(index) != nil {
			return nil, false
		}
		raw, ok := exactNativeAttr(index, "", "val")
		if !ok {
			return nil, false
		}
		value, err := parseCanonicalNativeInt(raw, 0, int64(count)-1)
		if err != nil {
			return nil, false
		}
		// Index with int64 directly: no narrowing conversion of an XML scalar.
		if seen[value] {
			return nil, false
		}
		seen[value] = true
		color, ok := nativeLiteralPieSolidPaint(point.Children[1], d)
		if !ok {
			return nil, false
		}
		colors[value] = color
	}
	for _, color := range colors {
		if color == "" {
			return nil, false
		}
	}
	return colors, true
}

func nativeLiteralPiePaintSequence(node *nativeXMLNode, ns string, names ...string) bool {
	if node == nil || requireOnlyNativeAttrs(node) != nil || !onlyNativeXMLSpace(node.Text) || len(node.Children) != len(names) {
		return false
	}
	for i, name := range names {
		if node.Children[i].Name != (xml.Name{Space: ns, Local: name}) {
			return false
		}
	}
	return true
}

func nativeLiteralPieSolidPaint(node *nativeXMLNode, d nativeExtractDialect) (string, bool) {
	if node == nil || node.Name != (xml.Name{Space: d.chart, Local: "spPr"}) || !nativeLiteralPiePaintSequence(node, d.drawing, "solidFill", "ln") || !nativeLiteralPiePaintSequence(node.Children[0], d.drawing, "srgbClr") || !nativeLiteralPiePaintSequence(node.Children[1], d.drawing, "noFill") || requireEmptyNativeElement(node.Children[1].Children[0]) != nil {
		return "", false
	}
	rgb := node.Children[0].Children[0]
	if requireOnlyNativeAttrs(rgb, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(rgb) != nil {
		return "", false
	}
	color, ok := exactNativeAttr(rgb, "", "val")
	if !ok || !inspectionRGB.MatchString(color) {
		return "", false
	}
	return "#" + strings.ToUpper(color), true
}
