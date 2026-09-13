package pptxpatch

import (
	"encoding/xml"
	"slices"
	"sort"
	"unicode/utf16"
)

type nativeChartLineSeries struct {
	Index, Order    int64
	Title           *string
	Values, XValues []string
	Color           string
	Width           int64
}
type nativeChartLine struct {
	Family       string
	Categories   []string
	Series       []nativeChartLineSeries
	XAxis, YAxis nativeChartAxis
}

// Straight connected series require a complete source round-join/flat-cap line.
// The line style is not reconstructed from a theme, marker, or first point.
func nativeChartConnectedLinePaint(node *nativeXMLNode, d nativeExtractDialect) (string, int64, bool) {
	return nativeChartLinePaint(node, d, true)
}
func extractNativeChartConnectedSeries(node *nativeXMLNode, d nativeExtractDialect, scatter bool) (*nativeChartLineSeries, []string, bool) {
	c := nativeChartChildren(node, d.chart)
	series := &nativeChartLineSeries{}
	var ok bool
	series.Index, ok = nativeChartInteger(c.take("idx"), 0, 4294967295)
	if !ok {
		return nil, nil, false
	}
	series.Order, ok = nativeChartInteger(c.take("order"), 0, nativeChartMaxSeries-1)
	if !ok {
		return nil, nil, false
	}
	if c.has("tx") {
		tx := nativeChartChildren(c.take("tx"), d.chart)
		title, ok := nativeChartText(tx.take("v"))
		if !ok || !tx.done() || nativeChartEncodedString.MatchString(title) || len(utf16.Encode([]rune(title))) > 1024 {
			return nil, nil, false
		}
		series.Title = &title
	}
	series.Color, series.Width, ok = nativeChartConnectedLinePaint(c.take("spPr"), d)
	if !ok {
		return nil, nil, false
	}
	marker := nativeChartChildren(c.take("marker"), d.chart)
	if _, ok := nativeChartToken(marker.take("symbol"), "none"); !ok || !marker.done() {
		return nil, nil, false
	}
	categories := []string{}
	xKind, yKind, literal := "cat", "val", "strLit"
	if scatter {
		xKind, yKind, literal = "xVal", "yVal", "numLit"
	}
	x := nativeChartChildren(c.take(xKind), d.chart)
	xValues, ok := nativeChartLiteralPoints(x.take(literal), d, scatter)
	if !ok || !x.done() {
		return nil, nil, false
	}
	y := nativeChartChildren(c.take(yKind), d.chart)
	series.Values, ok = nativeChartLiteralPoints(y.take("numLit"), d, true)
	if !ok || !y.done() || len(series.Values) != len(xValues) {
		return nil, nil, false
	}
	if _, ok := nativeChartInteger(c.take("smooth"), 0, 0); !ok || !c.done() {
		return nil, nil, false
	}
	if scatter {
		series.XValues = xValues
	} else {
		categories = xValues
	}
	return series, categories, true
}
func extractNativeChartConnected(payload []byte, part string, d nativeExtractDialect, scatter bool) *nativeChartLine {
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
	family := "lineChart"
	if scatter {
		family = "scatterChart"
	}
	lines := nativeChartChildren(plot.take(family), d.chart)
	if scatter {
		if _, ok := nativeChartToken(lines.take("scatterStyle"), "line"); !ok {
			return nil
		}
	} else {
		if _, ok := nativeChartToken(lines.take("grouping"), "standard"); !ok {
			return nil
		}
	}
	if _, ok := nativeChartInteger(lines.take("varyColors"), 0, 0); !ok {
		return nil
	}
	result := &nativeChartLine{Family: family, Categories: []string{}, Series: []nativeChartLineSeries{}}
	ids, orders := map[int64]bool{}, map[int64]bool{}
	for lines.has("ser") {
		series, categories, ok := extractNativeChartConnectedSeries(lines.take("ser"), d, scatter)
		if !ok || len(result.Series) >= nativeChartMaxSeries || ids[series.Index] || orders[series.Order] {
			return nil
		}
		ids[series.Index] = true
		orders[series.Order] = true
		if len(result.Series) == 0 {
			result.Categories = categories
		} else if !scatter && !slices.Equal(result.Categories, categories) {
			return nil
		}
		result.Series = append(result.Series, *series)
	}
	if len(result.Series) == 0 {
		return nil
	}
	sort.Slice(result.Series, func(i, j int) bool { return result.Series[i].Order < result.Series[j].Order })
	for i, series := range result.Series {
		if series.Order != int64(i) {
			return nil
		}
	}
	if !scatter {
		if _, ok := nativeChartInteger(lines.take("marker"), 0, 0); !ok {
			return nil
		}
		if _, ok := nativeChartInteger(lines.take("smooth"), 0, 0); !ok {
			return nil
		}
	}
	xID, ok := nativeChartInteger(lines.take("axId"), 0, 4294967295)
	if !ok {
		return nil
	}
	yID, ok := nativeChartInteger(lines.take("axId"), 0, 4294967295)
	if !ok || xID == yID || !lines.done() {
		return nil
	}
	axes, ok := extractNativeChartConnectedAxes(plot, d, scatter, xID, yID)
	if !ok {
		return nil
	}
	result.XAxis, result.YAxis = axes[0], axes[1]
	if !nativeChartTransparent(plot.take("spPr"), d) || !plot.done() || !chart.done() || !nativeChartTransparent(root.take("spPr"), d) || !root.done() {
		return nil
	}
	return result
}
func extractNativeChartLine(payload []byte, part string, d nativeExtractDialect) *nativeChartLine {
	return extractNativeChartConnected(payload, part, d, false)
}
