package pptxpatch

import (
	"encoding/xml"
	"slices"
)

// Stacked lines retain the existing straight-series paint/marker/source guards.
type nativeChartStackedLine struct {
	nativeChartLine
	Grouping string
}

func extractNativeChartStackedLine(payload []byte, part string, d nativeExtractDialect) *nativeChartStackedLine {
	scatter := false
	grouping := ""
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
		var ok bool
		grouping, ok = nativeChartToken(lines.take("grouping"), "stacked", "percentStacked")
		if !ok {
			return nil
		}
	}
	if _, ok := nativeChartInteger(lines.take("varyColors"), 0, 0); !ok {
		return nil
	}
	result := &nativeChartStackedLine{Grouping: grouping, nativeChartLine: nativeChartLine{Family: family, Categories: []string{}, Series: []nativeChartLineSeries{}}}
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
	// Office stacks in XML sequence even when c:order is a different
	// contiguous permutation. Preserve both, without sorting authoritative data.
	for i := range result.Series {
		if !orders[int64(i)] {
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

func extractNativeLiteralStackedLine(payload []byte, part string, d nativeExtractDialect) *NativeLiteralStackedLine {
	source := extractNativeChartStackedLine(payload, part, d)
	if source == nil {
		return nil
	}
	out := &NativeLiteralStackedLine{Grouping: source.Grouping, NativeLiteralConnected: NativeLiteralConnected{Profile: "literal-stacked-line-v1", DataOrigin: "literal", Categories: source.Categories, Series: []NativeLiteralConnectedSeries{}, XAxis: nativeLiteralChartAxis(source.XAxis, false), YAxis: nativeLiteralChartAxis(source.YAxis, true)}}
	for _, s := range source.Series {
		out.Series = append(out.Series, NativeLiteralConnectedSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, Color: s.Color, WidthEMU: s.Width})
	}
	return out
}
