package pptxpatch

import (
	"encoding/xml"
	"slices"
	"strings"
)

// Private descriptors retain source sequence and raw literals. References never
// carry resolved cache values; the outer workbook adapter owns cell authority.
type nativeChartRadarSeries struct {
	Index, Order                                      int64
	Title                                             *string
	TitleReference, CategoryReference, ValueReference *nativeChartReference
	Categories, Values                                []string
	Color                                             string
	Width                                             int64
	Fill                                              *string
}
type nativeChartRadarSource struct {
	Workbook                bool
	Style                   string
	Series                  []nativeChartRadarSeries
	CategoryAxis, ValueAxis nativeChartAxis
	ExternalData            *nativeXMLNode
	DispBlanksAs            *string
}

func nativeRadarPaint(node *nativeXMLNode, d nativeExtractDialect, filled bool) (string, int64, *string, bool) {
	if !filled {
		color, width, ok := nativeChartConnectedLinePaint(node, d)
		return color, width, nil, ok
	}
	if node == nil || len(node.Children) != 2 {
		return "", 0, nil, false
	}
	fill := node.Children[0]
	if fill.Name != (xml.Name{Space: d.drawing, Local: "solidFill"}) || !nativeLiteralPiePaintSequence(fill, d.drawing, "srgbClr") {
		return "", 0, nil, false
	}
	raw, ok := nativeChartAttribute(fill.Children[0])
	if !ok || !inspectionRGB.MatchString(raw) {
		return "", 0, nil, false
	}
	// Reuse the complete explicit line predicate on a private structural copy.
	// Only its no-fill precondition differs; the original source tree is untouched.
	projected := *node
	projected.Children = append([]*nativeXMLNode(nil), node.Children...)
	projected.Children[0] = &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "noFill"}}
	color, width, ok := nativeChartConnectedLinePaint(&projected, d)
	resolved := "#" + strings.ToUpper(raw)
	return color, width, &resolved, ok
}
func extractNativeRadarSeries(node *nativeXMLNode, d nativeExtractDialect, workbook, filled bool) (*nativeChartRadarSeries, bool) {
	c := nativeChartChildren(node, d.chart)
	header, ok := nativeChartWorkbookSeriesHeader(c, d)
	if !ok || !workbook && header.TitleReference != nil {
		return nil, false
	}
	out := &nativeChartRadarSeries{Index: header.Index, Order: header.Order, Title: header.Title, TitleReference: header.TitleReference}
	out.Color, out.Width, out.Fill, ok = nativeRadarPaint(c.take("spPr"), d, filled)
	if !ok {
		return nil, false
	}
	marker := nativeChartChildren(c.take("marker"), d.chart)
	if _, ok = nativeChartToken(marker.take("symbol"), "none"); !ok || !marker.done() {
		return nil, false
	}
	if workbook {
		out.CategoryReference, ok = nativeChartWorkbookSeriesReference(c, d, "cat", false)
		if !ok {
			return nil, false
		}
		out.ValueReference, ok = nativeChartWorkbookSeriesReference(c, d, "val", true)
		if !ok || out.CategoryReference.Range.Count != out.ValueReference.Range.Count || out.ValueReference.Range.Count < 3 || out.ValueReference.Range.Count > 256 {
			return nil, false
		}
	} else {
		categories := nativeChartChildren(c.take("cat"), d.chart)
		out.Categories, ok = nativeChartLiteralPoints(categories.take("strLit"), d, false)
		if !ok || !categories.done() {
			return nil, false
		}
		values := nativeChartChildren(c.take("val"), d.chart)
		out.Values, ok = nativeChartLiteralPoints(values.take("numLit"), d, true)
		if !ok || !values.done() || len(out.Values) != len(out.Categories) || len(out.Values) < 3 {
			return nil, false
		}
	}
	return out, c.done()
}

// Source grammar qualification only. Radial projection, axis display and filled
// painter order remain separate admission obligations; this is not a deck hook.
func extractNativeChartRadarSource(payload []byte, part string, d nativeExtractDialect, workbook bool) *nativeChartRadarSource {
	node, err := parseNativeXML(payload, part)
	if err != nil || node.Name != (xml.Name{Space: d.chart, Local: "chartSpace"}) {
		return nil
	}
	root := nativeChartChildren(node, d.chart)
	chart := nativeChartChildren(root.take("chart"), d.chart)
	if _, ok := nativeChartInteger(chart.take("autoTitleDeleted"), 1, 1); !ok {
		return nil
	}
	plot := nativeChartChildren(chart.take("plotArea"), d.chart)
	if requireEmptyNativeElement(plot.take("layout")) != nil {
		return nil
	}
	body := nativeChartChildren(plot.take("radarChart"), d.chart)
	style, ok := nativeChartToken(body.take("radarStyle"), "standard", "filled")
	if !ok {
		return nil
	}
	if _, ok = nativeChartInteger(body.take("varyColors"), 0, 0); !ok {
		return nil
	}
	out := &nativeChartRadarSource{Workbook: workbook, Style: style, Series: []nativeChartRadarSeries{}}
	indices, orders := map[int64]bool{}, map[int64]bool{}
	for body.has("ser") {
		s, valid := extractNativeRadarSeries(body.take("ser"), d, workbook, style == "filled")
		if !valid || len(out.Series) >= 16 || indices[s.Index] || orders[s.Order] {
			return nil
		}
		if len(out.Series) > 0 {
			first := out.Series[0]
			if workbook {
				if first.CategoryReference.Formula != s.CategoryReference.Formula || first.ValueReference.Range.Count != s.ValueReference.Range.Count {
					return nil
				}
			} else if !slices.Equal(first.Categories, s.Categories) {
				return nil
			}
		}
		indices[s.Index] = true
		orders[s.Order] = true
		out.Series = append(out.Series, *s)
	}
	if len(out.Series) == 0 {
		return nil
	}
	for i := range out.Series {
		if !orders[int64(i)] {
			return nil
		}
	}
	firstID, ok := nativeChartInteger(body.take("axId"), 0, 4294967295)
	if !ok {
		return nil
	}
	secondID, ok := nativeChartInteger(body.take("axId"), 0, 4294967295)
	if !ok || firstID == secondID || !body.done() {
		return nil
	}
	var category, value *nativeChartAxis
	for plot.has("catAx") || plot.has("valAx") {
		if plot.has("catAx") {
			if category != nil {
				return nil
			}
			category, ok = extractNativeChartAxisWithCrossBetween(plot.take("catAx"), d, false, false)
		} else {
			if value != nil {
				return nil
			}
			value, ok = extractNativeChartAxisWithCrossBetween(plot.take("valAx"), d, true, false)
		}
		if !ok {
			return nil
		}
	}
	if category == nil || value == nil || category.CrossAxisID != value.ID || value.CrossAxisID != category.ID || !((firstID == category.ID && secondID == value.ID) || (secondID == category.ID && firstID == value.ID)) {
		return nil
	}
	out.CategoryAxis, out.ValueAxis = *category, *value
	if !nativeChartTransparent(plot.take("spPr"), d) || !plot.done() {
		return nil
	}
	if _, ok = nativeChartInteger(chart.take("plotVisOnly"), 0, 0); !ok {
		return nil
	}
	if chart.has("dispBlanksAs") {
		v, valid := nativeChartToken(chart.take("dispBlanksAs"), "gap", "zero", "span")
		if !valid {
			return nil
		}
		out.DispBlanksAs = &v
	}
	if !chart.done() || !nativeChartTransparent(root.take("spPr"), d) {
		return nil
	}
	if workbook {
		out.ExternalData = root.take("externalData")
		if out.ExternalData == nil {
			return nil
		}
	}
	if !root.done() {
		return nil
	}
	return out
}
