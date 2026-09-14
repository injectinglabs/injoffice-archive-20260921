package pptxpatch

import (
	"encoding/xml"
	"strings"
)

type nativeChartAxis struct {
	Labels           *NativeChartAxisLabels
	ID               int64
	CrossAxisID      int64
	Orientation      string
	Position         string
	Deleted          bool
	Color            string
	Width            int64
	Min              string
	Max              string
	CrossesAt        string
	CategoryCrossing string
}

func nativeChartToken(node *nativeXMLNode, allowed ...string) (string, bool) {
	raw, ok := nativeChartAttribute(node)
	if !ok {
		return "", false
	}
	for _, value := range allowed {
		if raw == value {
			return raw, true
		}
	}
	return "", false
}
func nativeChartAxisPaint(node *nativeXMLNode, d nativeExtractDialect) (string, int64, bool) {
	return nativeChartLinePaint(node, d, false)
}

// Shared complete local line paint; series additionally require a round join.
func nativeChartLinePaint(node *nativeXMLNode, d nativeExtractDialect, roundJoin bool) (string, int64, bool) {
	c := nativeChartChildren(node, d.drawing)
	if requireEmptyNativeElement(c.take("noFill")) != nil {
		return "", 0, false
	}
	line := c.take("ln")
	if line == nil || !c.done() || requireOnlyNativeAttrs(line, xml.Name{Local: "w"}, xml.Name{Local: "cap"}, xml.Name{Local: "cmpd"}, xml.Name{Local: "algn"}) != nil || !onlyNativeXMLSpace(line.Text) {
		return "", 0, false
	}
	for name, value := range map[string]string{"cap": "flat", "cmpd": "sng", "algn": "ctr"} {
		raw, ok := exactNativeAttr(line, "", name)
		if !ok || raw != value {
			return "", 0, false
		}
	}
	raw, ok := exactNativeAttr(line, "", "w")
	if !ok {
		return "", 0, false
	}
	width, err := parseCanonicalNativeInt(raw, 1, 20116800)
	if err != nil {
		return "", 0, false
	}
	count := 4
	if roundJoin {
		count++
	}
	if len(line.Children) != count {
		return "", 0, false
	}
	fill := line.Children[0]
	if fill.Name != (xml.Name{Space: d.drawing, Local: "solidFill"}) || !nativeLiteralPiePaintSequence(fill, d.drawing, "srgbClr") {
		return "", 0, false
	}
	rgb := fill.Children[0]
	color, ok := nativeChartAttribute(rgb)
	if !ok || !inspectionRGB.MatchString(color) {
		return "", 0, false
	}
	dash := line.Children[1]
	if dash.Name != (xml.Name{Space: d.drawing, Local: "prstDash"}) {
		return "", 0, false
	}
	if _, ok := nativeChartToken(dash, "solid"); !ok {
		return "", 0, false
	}
	offset := 2
	if roundJoin {
		join := line.Children[2]
		if join.Name != (xml.Name{Space: d.drawing, Local: "round"}) || requireEmptyNativeElement(join) != nil {
			return "", 0, false
		}
		offset++
	}
	for i, name := range []string{"headEnd", "tailEnd"} {
		end := line.Children[i+offset]
		if end.Name != (xml.Name{Space: d.drawing, Local: name}) || requireOnlyNativeAttrs(end, xml.Name{Local: "type"}) != nil || requireOnlyNativeChildren(end) != nil {
			return "", 0, false
		}
		raw, ok := exactNativeAttr(end, "", "type")
		if !ok || raw != "none" {
			return "", 0, false
		}
	}
	return "#" + strings.ToUpper(color), width, true
}
func extractNativeChartAxis(node *nativeXMLNode, d nativeExtractDialect, valueAxis bool) (*nativeChartAxis, bool) {
	return extractNativeChartAxisWithCrossBetween(node, d, valueAxis, true)
}

// Scatter axes have no category midpoint/boundary setting. This private policy
// leaves the existing bar profile grammar unchanged.
func extractNativeChartAxisWithCrossBetween(node *nativeXMLNode, d nativeExtractDialect, valueAxis, requireBetween bool) (*nativeChartAxis, bool) {
	c := nativeChartChildren(node, d.chart)
	axis := &nativeChartAxis{}
	var ok bool
	axis.ID, ok = nativeChartInteger(c.take("axId"), 0, 4294967295)
	if !ok {
		return nil, false
	}
	scale := nativeChartChildren(c.take("scaling"), d.chart)
	axis.Orientation, ok = nativeChartToken(scale.take("orientation"), "minMax", "maxMin")
	if !ok {
		return nil, false
	}
	if valueAxis {
		axis.Max, ok = nativeChartAttribute(scale.take("max"))
		if !ok {
			return nil, false
		}
		axis.Min, ok = nativeChartAttribute(scale.take("min"))
		if !ok {
			return nil, false
		}
		maximum, err := parseNativeChartDecimal(axis.Max)
		if err != nil {
			return nil, false
		}
		minimum, err := parseNativeChartDecimal(axis.Min)
		if err != nil {
			return nil, false
		}
		zero, _ := parseNativeChartDecimal("0")
		if minimum.compare(maximum) >= 0 || minimum.compare(zero) > 0 || maximum.compare(zero) < 0 {
			return nil, false
		}
	}
	if !scale.done() {
		return nil, false
	}
	deleted, ok := nativeChartInteger(c.take("delete"), 0, 1)
	if !ok {
		return nil, false
	}
	axis.Deleted = deleted == 1
	axis.Position, ok = nativeChartToken(c.take("axPos"), "b", "l", "r", "t")
	if !ok {
		return nil, false
	}
	if !axis.Deleted {
		var format *string
		if c.has("numFmt") {
			n := c.take("numFmt")
			if !valueAxis || requireOnlyNativeAttrs(n, xml.Name{Local: "formatCode"}, xml.Name{Local: "sourceLinked"}) != nil || requireOnlyNativeChildren(n) != nil {
				return nil, false
			}
			raw, found := exactNativeAttr(n, "", "formatCode")
			linked, explicit := exactNativeAttr(n, "", "sourceLinked")
			if !found || !nativeChartFixedFormat(raw) || !explicit || linked != "0" {
				return nil, false
			}
			format = &raw
		}
		major, valid := nativeChartToken(c.take("majorTickMark"), "none", "out")
		if !valid {
			return nil, false
		}
		if _, valid := nativeChartToken(c.take("minorTickMark"), "none"); !valid {
			return nil, false
		}
		position, valid := nativeChartToken(c.take("tickLblPos"), "none", "low", "high")
		if !valid {
			return nil, false
		}
		axis.Color, axis.Width, ok = nativeChartAxisPaint(c.take("spPr"), d)
		if !ok {
			return nil, false
		}
		if position == "none" {
			if format != nil || major != "none" {
				return nil, false
			}
		} else {
			if valueAxis != (format != nil) {
				return nil, false
			}
			style, valid := extractNativeChartAxisLabelStyle(c.take("txPr"), d)
			if !valid {
				return nil, false
			}
			axis.Labels = &NativeChartAxisLabels{Profile: "explicit-axis-labels-v1", Position: position, MajorTickMark: major, Style: *style, NumberFormat: format}
		}
	}

	axis.CrossAxisID, ok = nativeChartInteger(c.take("crossAx"), 0, 4294967295)
	if !ok || axis.ID == axis.CrossAxisID {
		return nil, false
	}
	if valueAxis {
		// crossesAt belongs to the value axis and locates the perpendicular category
		// axis. Restrict this first profile to exact zero, including lexical aliases.
		axis.CrossesAt, ok = nativeChartAttribute(c.take("crossesAt"))
		if !ok {
			return nil, false
		}
		cross, err := parseNativeChartDecimal(axis.CrossesAt)
		if err != nil || !cross.isZero() {
			return nil, false
		}
		if requireBetween {
			if _, ok := nativeChartToken(c.take("crossBetween"), "between"); !ok {
				return nil, false
			}
		}
	} else {
		if axis.CategoryCrossing, ok = nativeChartToken(c.take("crosses"), "min"); !ok {
			return nil, false
		}
	}
	if axis.Labels != nil {
		if valueAxis {
			unit, valid := nativeChartAttribute(c.take("majorUnit"))
			if !valid {
				return nil, false
			}
			if _, valid := nativeChartAxisTickCount(axis.Min, axis.Max, unit); !valid {
				return nil, false
			}
			axis.Labels.MajorUnit = &unit
		} else {
			// Explicit single-level labels: no automatic skips or offset-base guessing.
			for _, item := range []struct{ name, value string }{{"auto", "0"}, {"lblAlgn", "ctr"}, {"lblOffset", "0"}, {"tickLblSkip", "1"}, {"tickMarkSkip", "1"}, {"noMultiLvlLbl", "1"}} {
				expected := item.value
				if item.name == "lblOffset" && d.chart == nsChartStrict {
					expected = "0%"
				}
				if _, valid := nativeChartToken(c.take(item.name), expected); !valid {
					return nil, false
				}
			}
		}
	}

	return axis, c.done()
}
