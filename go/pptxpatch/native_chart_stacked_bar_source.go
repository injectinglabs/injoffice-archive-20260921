package pptxpatch

import (
	"encoding/xml"
	"slices"
)

// Dedicated admission preserves the existing clustered profile unchanged.
type nativeChartStackedBar struct {
	nativeChartBar
	Grouping string
}

func extractNativeChartStackedBar(payload []byte, part string, d nativeExtractDialect) *nativeChartStackedBar {
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
	bars := nativeChartChildren(plot.take("barChart"), d.chart)
	direction, ok := nativeChartToken(bars.take("barDir"), "bar", "col")
	if !ok {
		return nil
	}
	grouping, ok := nativeChartToken(bars.take("grouping"), "stacked", "percentStacked")
	if !ok {
		return nil
	}
	if _, ok := nativeChartInteger(bars.take("varyColors"), 0, 0); !ok {
		return nil
	}
	result := &nativeChartStackedBar{Grouping: grouping, nativeChartBar: nativeChartBar{Direction: direction, Series: []nativeChartLiteralSeries{}}}
	indices := map[int64]bool{}
	orders := map[int64]bool{}
	for bars.has("ser") {
		series, ok := extractNativeChartLiteralSeries(bars.take("ser"), d)
		if !ok || len(result.Series) >= nativeChartMaxSeries || indices[series.Index] || orders[series.Order] {
			return nil
		}
		indices[series.Index] = true
		orders[series.Order] = true
		if len(result.Series) == 0 {
			result.Categories = series.Categories
		} else if !slices.Equal(result.Categories, series.Categories) {
			return nil
		}
		result.Series = append(result.Series, *series)
	}
	if len(result.Series) == 0 {
		return nil
	}
	// Office stacks in XML sequence even when c:order is a different
	// contiguous permutation. Preserve both, without sorting authoritative data.
	for i := range result.Series {
		if !orders[int64(i)] {
			return nil
		}
	}
	result.GapWidth, ok = nativeChartPercentage(bars.take("gapWidth"), d, 0, 500)
	if !ok {
		return nil
	}
	if _, ok := nativeChartPercentage(bars.take("overlap"), d, 100, 100); !ok {
		return nil
	}
	firstID, ok := nativeChartInteger(bars.take("axId"), 0, 4294967295)
	if !ok {
		return nil
	}
	secondID, ok := nativeChartInteger(bars.take("axId"), 0, 4294967295)
	if !ok || firstID == secondID || !bars.done() {
		return nil
	}
	// Axis list ordering need not mirror the barChart's ID list: bind by ID.
	var category, value *nativeChartAxis
	for plot.has("catAx") || plot.has("valAx") {
		if plot.has("catAx") {
			if category != nil {
				return nil
			}
			category, ok = extractNativeChartAxis(plot.take("catAx"), d, false)
		} else {
			if value != nil {
				return nil
			}
			value, ok = extractNativeChartAxis(plot.take("valAx"), d, true)
		}
		if !ok {
			return nil
		}
	}
	if category == nil || value == nil || category.CrossAxisID != value.ID || value.CrossAxisID != category.ID || !((firstID == category.ID && secondID == value.ID) || (secondID == category.ID && firstID == value.ID)) {
		return nil
	}
	// Position must agree with a perpendicular bar/column pair. Crossing and
	// orientation determine the actual line coordinates inside the host frame.
	if direction == "col" {
		if category.Position != "b" || value.Position != "l" {
			return nil
		}
	} else {
		if category.Position != "l" || value.Position != "b" {
			return nil
		}
	}
	if !nativeChartTransparent(plot.take("spPr"), d) || !plot.done() || !chart.done() || !nativeChartTransparent(root.take("spPr"), d) || !root.done() {
		return nil
	}
	if !validNativeChartAxisLabels(nativeLiteralChartAxis(*category, false), nativeLiteralChartAxis(*value, true), false) || !validNativeChartAxisLabels(nativeLiteralChartAxis(*value, true), nativeLiteralChartAxis(*category, false), true) {
		return nil
	}
	result.CategoryAxis = *category
	result.ValueAxis = *value
	return result
}

func extractNativeLiteralStackedBar(payload []byte, part string, d nativeExtractDialect) *NativeLiteralStackedBar {
	source := extractNativeChartStackedBar(payload, part, d)
	if source == nil {
		return nil
	}
	direction := "bar"
	if source.Direction == "col" {
		direction = "column"
	}
	out := &NativeLiteralStackedBar{Profile: "literal-stacked-bar-v1", DataOrigin: "literal", Grouping: source.Grouping, BarDirection: direction, GapWidth: source.GapWidth, Overlap: 100, Categories: source.Categories, Series: []NativeLiteralBarSeries{}, CategoryAxis: nativeLiteralChartAxis(source.CategoryAxis, false), ValueAxis: nativeLiteralChartAxis(source.ValueAxis, true)}
	for _, s := range source.Series {
		out.Series = append(out.Series, NativeLiteralBarSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, Colors: s.Colors})
	}
	return out
}
