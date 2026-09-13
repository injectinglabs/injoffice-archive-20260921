package pptxpatch

import (
	"encoding/xml"
	"slices"
	"sort"
	"strings"
)

type nativeChartBar struct {
	Direction    string
	Categories   []string
	Series       []nativeChartLiteralSeries
	CategoryAxis nativeChartAxis
	ValueAxis    nativeChartAxis
	GapWidth     int64
}

func nativeChartTransparent(node *nativeXMLNode, d nativeExtractDialect) bool {
	return nativeLiteralPiePaintSequence(node, d.drawing, "noFill", "ln") && requireEmptyNativeElement(node.Children[0]) == nil && nativeLiteralPiePaintSequence(node.Children[1], d.drawing, "noFill") && requireEmptyNativeElement(node.Children[1].Children[0]) == nil
}

// Strict uses an explicit percent suffix; Transitional additionally permits
// the legacy integer spelling. Both encode a whole percentage, not EMU.
func nativeChartPercentage(node *nativeXMLNode, d nativeExtractDialect, low, high int64) (int64, bool) {
	raw, ok := nativeChartAttribute(node)
	if !ok {
		return 0, false
	}
	if strings.HasSuffix(raw, "%") {
		raw = strings.TrimSuffix(raw, "%")
	} else if d.chart == nsChartStrict {
		return 0, false
	}
	value, err := parseCanonicalNativeInt(raw, low, high)
	return value, err == nil
}
func extractNativeChartBar(payload []byte, part string, d nativeExtractDialect) *nativeChartBar {
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
	if _, ok := nativeChartToken(bars.take("grouping"), "clustered"); !ok {
		return nil
	}
	if _, ok := nativeChartInteger(bars.take("varyColors"), 0, 0); !ok {
		return nil
	}
	result := &nativeChartBar{Direction: direction, Series: []nativeChartLiteralSeries{}}
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
	sort.Slice(result.Series, func(i, j int) bool { return result.Series[i].Order < result.Series[j].Order })
	for i, series := range result.Series {
		if series.Order != int64(i) {
			return nil
		}
	}
	result.GapWidth, ok = nativeChartPercentage(bars.take("gapWidth"), d, 0, 500)
	if !ok {
		return nil
	}
	if _, ok := nativeChartPercentage(bars.take("overlap"), d, 0, 0); !ok {
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
	result.CategoryAxis = *category
	result.ValueAxis = *value
	return result
}

func extractNativeLiteralBar(payload []byte, part string, d nativeExtractDialect) *NativeLiteralBar {
	source := extractNativeChartBar(payload, part, d)
	if source == nil {
		return nil
	}

	direction := "bar"
	if source.Direction == "col" {
		direction = "column"
	}
	result := &NativeLiteralBar{Profile: "literal-bar-v1", BarDirection: direction, Grouping: "clustered", DataOrigin: "literal", GapWidth: source.GapWidth, Overlap: 0, Categories: source.Categories, Series: []NativeLiteralBarSeries{}, CategoryAxis: nativeLiteralChartAxis(source.CategoryAxis, false), ValueAxis: nativeLiteralChartAxis(source.ValueAxis, true)}
	for _, series := range source.Series {
		result.Series = append(result.Series, NativeLiteralBarSeries{Index: series.Index, Order: series.Order, Title: series.Title, Values: series.Values, Colors: series.Colors})
	}
	return result
}

func nativeLiteralChartAxis(a nativeChartAxis, value bool) NativeLiteralBarAxis {
	result := NativeLiteralBarAxis{ID: a.ID, CrossAxisID: a.CrossAxisID, Orientation: a.Orientation, Position: a.Position, Deleted: a.Deleted}
	if !a.Deleted {
		result.Color = &a.Color
		result.WidthEMU = &a.Width
	}
	if value {
		result.Min = &a.Min
		result.Max = &a.Max
		result.CrossesAt = &a.CrossesAt
	}
	return result
}
