package pptxpatch

import (
	"encoding/xml"
)

type nativeChartWorkbookSource struct {
	Family         string
	Grouping       string
	Overlap        *int64
	Direction      string
	GapWidth       int64
	BubbleScale    *int64
	SizeRepresents string
	XAxis, YAxis   nativeChartAxis
	Series         []nativeChartWorkbookSeries
	ExternalData   *nativeXMLNode
	DispBlanksAs   *string
}

// This parallel source descriptor path never edits XML or substitutes literal
// values into the existing literal-only parser. Rendering awaits an explicit
// embedded-workbook resolver; all existing literal profiles stay unchanged.
func extractNativeChartWorkbookSource(payload []byte, part string, d nativeExtractDialect) *nativeChartWorkbookSource {
	if bubble := extractNativeChartBubbleSource(payload, part, d, true); bubble != nil {
		out := &nativeChartWorkbookSource{Family: "bubbleChart", BubbleScale: &bubble.BubbleScale, SizeRepresents: bubble.SizeRepresents, XAxis: bubble.XAxis, YAxis: bubble.YAxis, ExternalData: bubble.ExternalData, DispBlanksAs: bubble.DispBlanksAs, Series: []nativeChartWorkbookSeries{}}
		for _, s := range bubble.Series {
			out.Series = append(out.Series, nativeChartWorkbookSeries{Index: s.Index, Order: s.Order, Title: s.Title, TitleReference: s.TitleReference, XReference: s.XReference, ValueReference: s.ValueReference, SizeReference: s.SizeReference, Colors: s.Colors})
		}
		return out
	}
	node, e := parseNativeXML(payload, part)
	if e != nil || node.Name != (xml.Name{Space: d.chart, Local: "chartSpace"}) {
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
	family := ""
	for _, name := range []string{"barChart", "lineChart", "scatterChart"} {
		if plot.has(name) {
			family = name
			break
		}
	}
	if family == "" {
		return nil
	}
	body := nativeChartChildren(plot.take(family), d.chart)
	source := &nativeChartWorkbookSource{Family: family, Series: []nativeChartWorkbookSeries{}}
	var ok bool
	if family == "barChart" {
		source.Direction, ok = nativeChartToken(body.take("barDir"), "col", "bar")
		if !ok {
			return nil
		}
		if source.Grouping, ok = nativeChartToken(body.take("grouping"), "clustered", "stacked", "percentStacked"); !ok {
			return nil
		}
	} else if family == "lineChart" {
		if source.Grouping, ok = nativeChartToken(body.take("grouping"), "standard", "stacked", "percentStacked"); !ok {
			return nil
		}
	} else {
		if _, ok = nativeChartToken(body.take("scatterStyle"), "line"); !ok {
			return nil
		}
	}
	if _, ok = nativeChartInteger(body.take("varyColors"), 0, 0); !ok {
		return nil
	}
	indices, orders := map[int64]bool{}, map[int64]bool{}
	for body.has("ser") {
		var series *nativeChartWorkbookSeries
		if family == "barChart" {
			series, ok = extractNativeChartWorkbookBarSeries(body.take("ser"), d)
		} else {
			series, ok = extractNativeChartWorkbookConnectedSeries(body.take("ser"), d, family == "scatterChart")
		}
		if !ok || len(source.Series) >= nativeChartMaxSeries || indices[series.Index] || orders[series.Order] {
			return nil
		}
		if family != "scatterChart" && len(source.Series) > 0 && source.Series[0].ValueReference.Range.Count != series.ValueReference.Range.Count {
			return nil
		}
		indices[series.Index] = true
		orders[series.Order] = true
		source.Series = append(source.Series, *series)
	}
	if len(source.Series) == 0 {
		return nil
	}
	stacked := source.Grouping == "stacked" || source.Grouping == "percentStacked"
	// All families retain XML sequence and original order metadata.
	for i := range source.Series {
		if !orders[int64(i)] {
			return nil
		}
	}
	if family == "barChart" {
		source.GapWidth, ok = nativeChartPercentage(body.take("gapWidth"), d, 0, 500)
		if !ok {
			return nil
		}
		overlap := int64(0)
		if stacked {
			overlap = 100
		}
		if _, ok = nativeChartPercentage(body.take("overlap"), d, overlap, overlap); !ok {
			return nil
		}
		if stacked {
			source.Overlap = &overlap
		}
	} else if family == "lineChart" {
		if _, ok = nativeChartInteger(body.take("marker"), 0, 0); !ok {
			return nil
		}
		if _, ok = nativeChartInteger(body.take("smooth"), 0, 0); !ok {
			return nil
		}
	}
	first, ok := nativeChartInteger(body.take("axId"), 0, 4294967295)
	if !ok {
		return nil
	}
	second, ok := nativeChartInteger(body.take("axId"), 0, 4294967295)
	if !ok || first == second || !body.done() {
		return nil
	}
	axes, ok := extractNativeWorkbookChartAxes(plot, d, source.Family, source.Direction, first, second)
	if !ok {
		return nil
	}
	source.XAxis, source.YAxis = axes[0], axes[1]
	if !nativeChartTransparent(plot.take("spPr"), d) || !plot.done() {
		return nil
	}
	// All referenced cells are plotted. No hidden-cell filtering is guessed.
	if _, ok = nativeChartInteger(chart.take("plotVisOnly"), 0, 0); !ok {
		return nil
	}
	if chart.has("dispBlanksAs") {
		blank, ok := nativeChartToken(chart.take("dispBlanksAs"), "gap", "zero", "span")
		if !ok {
			return nil
		}
		source.DispBlanksAs = &blank
	}
	if !chart.done() || !nativeChartTransparent(root.take("spPr"), d) {
		return nil
	}
	source.ExternalData = root.take("externalData")
	if source.ExternalData == nil || !root.done() {
		return nil
	}
	return source
}
func extractNativeWorkbookChartAxes(plot *nativeChartCursor, d nativeExtractDialect, family, direction string, first, second int64) ([2]nativeChartAxis, bool) {
	if family != "barChart" {
		return extractNativeChartConnectedAxes(plot, d, family == "scatterChart", first, second)
	}
	var result [2]nativeChartAxis
	var category, value *nativeChartAxis
	for plot.has("catAx") || plot.has("valAx") {
		var ok bool
		if plot.has("catAx") {
			if category != nil {
				return result, false
			}
			category, ok = extractNativeChartAxis(plot.take("catAx"), d, false)
		} else {
			if value != nil {
				return result, false
			}
			value, ok = extractNativeChartAxis(plot.take("valAx"), d, true)
		}
		if !ok {
			return result, false
		}
	}
	if category == nil || value == nil || category.ID == value.ID || category.CrossAxisID != value.ID || value.CrossAxisID != category.ID || !((category.ID == first && value.ID == second) || (category.ID == second && value.ID == first)) {
		return result, false
	}
	if direction == "col" {
		if category.Position != "b" || value.Position != "l" {
			return result, false
		}
		result = [2]nativeChartAxis{*category, *value}
	} else {
		if category.Position != "l" || value.Position != "b" {
			return result, false
		}
		result = [2]nativeChartAxis{*value, *category}
	}
	if !validNativeChartAxisLabels(nativeLiteralChartAxis(*category, false), nativeLiteralChartAxis(*value, true), false) || !validNativeChartAxisLabels(nativeLiteralChartAxis(*value, true), nativeLiteralChartAxis(*category, false), true) {
		return result, false
	}
	return result, true
}
