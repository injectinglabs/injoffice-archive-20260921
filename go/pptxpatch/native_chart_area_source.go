package pptxpatch

import (
	"encoding/xml"
	"slices"
	"unicode/utf16"
)

type nativeChartAreaSeries struct {
	Index, Order int64
	Title        *string
	Values       []string
	Color        string
}
type nativeChartArea struct {
	Grouping     string
	Categories   []string
	Series       []nativeChartAreaSeries
	XAxis, YAxis nativeChartAxis
}

func extractNativeChartAreaSeries(node *nativeXMLNode, d nativeExtractDialect) (*nativeChartAreaSeries, []string, bool) {
	c := nativeChartChildren(node, d.chart)
	series := &nativeChartAreaSeries{}
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
	// A uniform explicit fill and no outline avoid inventing interpolation of
	// per-point styles or stroking artificial interval/clipping boundaries.
	series.Color, ok = nativeLiteralPieSolidPaint(c.take("spPr"), d)
	if !ok {
		return nil, nil, false
	}
	cat := nativeChartChildren(c.take("cat"), d.chart)
	categories, ok := nativeChartLiteralPoints(cat.take("strLit"), d, false)
	if !ok || !cat.done() {
		return nil, nil, false
	}
	val := nativeChartChildren(c.take("val"), d.chart)
	series.Values, ok = nativeChartLiteralPoints(val.take("numLit"), d, true)
	if !ok || !val.done() || len(series.Values) != len(categories) || !c.done() {
		return nil, nil, false
	}
	return series, categories, true
}

// Strict literal source admission; opaque chart ownership remains unchanged.
func extractNativeChartArea(payload []byte, part string, d nativeExtractDialect) *nativeChartArea {
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
	area := nativeChartChildren(plot.take("areaChart"), d.chart)
	grouping, ok := nativeChartToken(area.take("grouping"), "standard", "stacked", "percentStacked")
	if !ok {
		return nil
	}
	if _, ok := nativeChartInteger(area.take("varyColors"), 0, 0); !ok {
		return nil
	}
	result := &nativeChartArea{Grouping: grouping, Categories: []string{}, Series: []nativeChartAreaSeries{}}
	indices, orders := map[int64]bool{}, map[int64]bool{}
	for area.has("ser") {
		series, categories, ok := extractNativeChartAreaSeries(area.take("ser"), d)
		if !ok || len(result.Series) >= nativeChartMaxSeries || indices[series.Index] || orders[series.Order] {
			return nil
		}
		indices[series.Index], orders[series.Order] = true, true
		if len(result.Series) == 0 {
			result.Categories = categories
		} else if !slices.Equal(result.Categories, categories) {
			return nil
		}
		result.Series = append(result.Series, *series)
	}
	if len(result.Series) == 0 {
		return nil
	}
	// Preserve XML series sequence; order is a complete metadata permutation.
	// The source record retains lexemes; exact cumulative geometry belongs to
	// the renderer. Admission needs only already-bounded decimal signs/order,
	// not allocating every rational band during read-only extraction.
	for i, series := range result.Series {
		if !orders[int64(i)] {
			return nil
		}
		if grouping != "standard" {
			for _, raw := range series.Values {
				value, err := parseNativeChartDecimal(raw)
				if err != nil || value.coefficient.Sign() < 0 {
					return nil
				}
			}
		}
	}
	xID, ok := nativeChartInteger(area.take("axId"), 0, 4294967295)
	if !ok {
		return nil
	}
	yID, ok := nativeChartInteger(area.take("axId"), 0, 4294967295)
	if !ok || xID == yID || !area.done() {
		return nil
	}
	axes, ok := extractNativeChartConnectedAxes(plot, d, false, xID, yID)
	if !ok {
		return nil
	}
	result.XAxis, result.YAxis = axes[0], axes[1]
	if !nativeChartTransparent(plot.take("spPr"), d) || !plot.done() || !chart.done() || !nativeChartTransparent(root.take("spPr"), d) || !root.done() {
		return nil
	}
	return result
}
