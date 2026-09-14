package pptxpatch

import (
	"encoding/xml"
)

// Private preparation descriptor. Reference cells remain unresolved until the
// existing actual-XLSX bridge admits them; caches never populate literal arrays.
type nativeChartBubbleSeries struct {
	Index, Order                              int64
	Title                                     *string
	TitleReference                            *nativeChartReference
	XValues, Values, Sizes                    []string
	XReference, ValueReference, SizeReference *nativeChartReference
	Colors                                    []string
}
type nativeChartBubbleSource struct {
	Workbook       bool
	BubbleScale    int64
	SizeRepresents string
	Series         []nativeChartBubbleSeries
	XAxis, YAxis   nativeChartAxis
	ExternalData   *nativeXMLNode
	DispBlanksAs   *string
}

func nativeBubbleNumbers(c *nativeChartCursor, d nativeExtractDialect, container string, workbook bool) ([]string, *nativeChartReference, bool) {
	if workbook {
		ref, ok := nativeChartWorkbookSeriesReference(c, d, container, true)
		return nil, ref, ok
	}
	child := nativeChartChildren(c.take(container), d.chart)
	values, ok := nativeChartLiteralPoints(child.take("numLit"), d, true)
	return values, nil, ok && child.done()
}
func extractNativeBubbleSeries(node *nativeXMLNode, d nativeExtractDialect, workbook bool) (*nativeChartBubbleSeries, bool) {
	c := nativeChartChildren(node, d.chart)
	header, ok := nativeChartWorkbookSeriesHeader(c, d)
	if !ok || !workbook && header.TitleReference != nil {
		return nil, false
	}
	out := &nativeChartBubbleSeries{Index: header.Index, Order: header.Order, Title: header.Title, TitleReference: header.TitleReference}
	var paint *nativeXMLNode
	if c.has("spPr") {
		paint = c.take("spPr")
	}
	if _, ok = nativeChartInteger(c.take("invertIfNegative"), 0, 0); !ok {
		return nil, false
	}
	points := []*nativeXMLNode{}
	for c.has("dPt") {
		points = append(points, c.take("dPt"))
		if len(points) > 256 {
			return nil, false
		}
	}
	out.XValues, out.XReference, ok = nativeBubbleNumbers(c, d, "xVal", workbook)
	if !ok {
		return nil, false
	}
	out.Values, out.ValueReference, ok = nativeBubbleNumbers(c, d, "yVal", workbook)
	if !ok {
		return nil, false
	}
	out.Sizes, out.SizeReference, ok = nativeBubbleNumbers(c, d, "bubbleSize", workbook)
	if !ok {
		return nil, false
	}
	if _, ok = nativeChartInteger(c.take("bubble3D"), 0, 0); !ok || !c.done() {
		return nil, false
	}
	count := len(out.Values)
	if workbook {
		if out.XReference.Range.Count != out.ValueReference.Range.Count || out.SizeReference.Range.Count != out.ValueReference.Range.Count {
			return nil, false
		}
		if out.ValueReference.Range.Count < 1 || out.ValueReference.Range.Count > 256 {
			return nil, false
		}
		count = int(out.ValueReference.Range.Count)
	} else {
		if len(out.XValues) != count || len(out.Sizes) != count {
			return nil, false
		}
		for _, raw := range out.Sizes {
			value, err := parseNativeChartDecimal(raw)
			if err != nil || value.coefficient.Sign() < 0 {
				return nil, false
			}
		}
	}
	out.Colors, ok = nativeChartSeriesColors(paint, points, count, d)
	return out, ok
}

// Initial 2D nonnegative profile requires explicit visual clauses. The chart's
// default radius is host policy, not inferred from cached values or Office.
func extractNativeChartBubbleSource(payload []byte, part string, d nativeExtractDialect, workbook bool) *nativeChartBubbleSource {
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
	body := nativeChartChildren(plot.take("bubbleChart"), d.chart)
	if _, ok := nativeChartInteger(body.take("varyColors"), 0, 0); !ok {
		return nil
	}
	out := &nativeChartBubbleSource{Workbook: workbook, Series: []nativeChartBubbleSeries{}}
	indices, orders := map[int64]bool{}, map[int64]bool{}
	for body.has("ser") {
		s, ok := extractNativeBubbleSeries(body.take("ser"), d, workbook)
		if !ok || len(out.Series) >= 16 || indices[s.Index] || orders[s.Order] {
			return nil
		}
		indices[s.Index] = true
		orders[s.Order] = true
		out.Series = append(out.Series, *s)
	}
	if len(out.Series) == 0 {
		return nil
	}
	// Preserve XML series sequence; order is a complete metadata permutation.
	for i := range out.Series {
		if !orders[int64(i)] {
			return nil
		}
	}
	if _, ok := nativeChartInteger(body.take("bubble3D"), 0, 0); !ok {
		return nil
	}
	var ok bool
	out.BubbleScale, ok = nativeChartPercentage(body.take("bubbleScale"), d, 0, 300)
	if !ok {
		return nil
	}
	if _, ok = nativeChartInteger(body.take("showNegBubbles"), 0, 0); !ok {
		return nil
	}
	out.SizeRepresents, ok = nativeChartToken(body.take("sizeRepresents"), "area", "w")
	if !ok {
		return nil
	}
	x, ok := nativeChartInteger(body.take("axId"), 0, 4294967295)
	if !ok {
		return nil
	}
	y, ok := nativeChartInteger(body.take("axId"), 0, 4294967295)
	if !ok || x == y || !body.done() {
		return nil
	}
	axes, ok := extractNativeChartConnectedAxes(plot, d, true, x, y)
	if !ok {
		return nil
	}
	out.XAxis, out.YAxis = axes[0], axes[1]
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
