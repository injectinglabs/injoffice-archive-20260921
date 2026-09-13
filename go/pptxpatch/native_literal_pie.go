package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// All-or-nothing source-literal profile. Frame fitting is host policy, not Office layout.
func extractNativeLiteralPie(payload []byte, part string, d nativeExtractDialect) *NativeLiteralPie {
	root, err := parseNativeXML(payload, part)
	if err != nil {
		return nil
	}
	seq := func(n *nativeXMLNode, ns string, names ...string) bool {
		if n == nil || requireOnlyNativeAttrs(n) != nil || !onlyNativeXMLSpace(n.Text) || len(n.Children) != len(names) {
			return false
		}
		for i, name := range names {
			if n.Children[i].Name != (xml.Name{Space: ns, Local: name}) {
				return false
			}
		}
		return true
	}
	val := func(n *nativeXMLNode, name string, min, max int64) (int64, error) {
		if n == nil || n.Name != (xml.Name{Space: d.chart, Local: name}) || requireOnlyNativeAttrs(n, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(n) != nil {
			return 0, fmt.Errorf("invalid pie value")
		}
		raw, ok := exactNativeAttr(n, "", "val")
		if !ok {
			return 0, fmt.Errorf("missing pie value")
		}
		return parseCanonicalNativeInt(raw, min, max)
	}
	transparent := func(n *nativeXMLNode) bool {
		return seq(n, d.drawing, "noFill", "ln") && requireEmptyNativeElement(n.Children[0]) == nil && seq(n.Children[1], d.drawing, "noFill") && requireEmptyNativeElement(n.Children[1].Children[0]) == nil
	}
	if root.Name != (xml.Name{Space: d.chart, Local: "chartSpace"}) || !seq(root, d.chart, "chart", "spPr") || !transparent(root.Children[1]) {
		return nil
	}
	chart := root.Children[0]
	if !seq(chart, d.chart, "autoTitleDeleted", "plotArea") {
		return nil
	}
	if _, err := val(chart.Children[0], "autoTitleDeleted", 1, 1); err != nil {
		return nil
	}
	plot := chart.Children[1]
	if !seq(plot, d.chart, "layout", "pieChart", "spPr") || requireEmptyNativeElement(plot.Children[0]) != nil || !transparent(plot.Children[2]) {
		return nil
	}
	pie := plot.Children[1]
	if !seq(pie, d.chart, "varyColors", "ser", "firstSliceAng") {
		return nil
	}
	if _, err := val(pie.Children[0], "varyColors", 0, 0); err != nil {
		return nil
	}
	angle, err := val(pie.Children[2], "firstSliceAng", 0, 360)
	if err != nil {
		return nil
	}
	ser := pie.Children[1]
	if requireOnlyNativeAttrs(ser) != nil || !onlyNativeXMLSpace(ser.Text) || len(ser.Children) < 3 || len(ser.Children) > 68 {
		return nil
	}
	if _, err := val(ser.Children[0], "idx", 0, 0); err != nil {
		return nil
	}
	if _, err := val(ser.Children[1], "order", 0, 0); err != nil {
		return nil
	}
	value := ser.Children[len(ser.Children)-1]
	if value.Name != (xml.Name{Space: d.chart, Local: "val"}) || !seq(value, d.chart, "numLit") {
		return nil
	}
	lit := value.Children[0]
	if requireOnlyNativeAttrs(lit) != nil || !onlyNativeXMLSpace(lit.Text) || len(lit.Children) < 3 || len(lit.Children) > 66 {
		return nil
	}
	format := lit.Children[0]
	if format.Name != (xml.Name{Space: d.chart, Local: "formatCode"}) || requireOnlyNativeAttrs(format) != nil || len(format.Children) != 0 || format.Text != "General" {
		return nil
	}
	count, err := val(lit.Children[1], "ptCount", 1, 64)
	if err != nil || count != int64(len(lit.Children)-2) {
		return nil
	}
	colors, ok := nativeLiteralPieColors(ser.Children[2:len(ser.Children)-1], len(lit.Children)-2, d)
	if !ok {
		return nil
	}
	result := &NativeLiteralPie{Profile: "literal-pie-v1", FirstSliceAngle: angle, Values: []int64{}, Colors: colors}
	for i := 0; i < len(lit.Children)-2; i++ {
		pt := lit.Children[i+2]
		if pt.Name != (xml.Name{Space: d.chart, Local: "pt"}) || requireOnlyNativeAttrs(pt, xml.Name{Local: "idx"}) != nil || requireOnlyNativeChildren(pt, xml.Name{Space: d.chart, Local: "v"}) != nil || len(pt.Children) != 1 {
			return nil
		}
		idx, ok := exactNativeAttr(pt, "", "idx")
		if !ok || idx != strconv.Itoa(i) {
			return nil
		}
		v := pt.Children[0]
		if requireOnlyNativeAttrs(v) != nil || len(v.Children) != 0 {
			return nil
		}
		number, err := parseCanonicalNativeInt(v.Text, 1, 1_000_000_000)
		if err != nil {
			return nil
		}
		result.Values = append(result.Values, number)
	}
	return result
}
