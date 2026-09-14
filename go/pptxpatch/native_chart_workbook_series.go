package pptxpatch

import "unicode/utf16"

// Dedicated source descriptors do not contain resolved/cached cell values.
// A referenced title is one string cell; source literal titles remain literal.
type nativeChartWorkbookSeries struct {
	Index, Order      int64
	Title             *string
	TitleReference    *nativeChartReference
	CategoryReference *nativeChartReference
	XReference        *nativeChartReference
	ValueReference    *nativeChartReference
	SizeReference     *nativeChartReference
	Colors            []string
	Color             string
	Width             int64
	Fill              *string
}

func nativeChartWorkbookSeriesHeader(c *nativeChartCursor, d nativeExtractDialect) (*nativeChartWorkbookSeries, bool) {
	result := &nativeChartWorkbookSeries{}
	var ok bool
	result.Index, ok = nativeChartInteger(c.take("idx"), 0, 4294967295)
	if !ok {
		return nil, false
	}
	result.Order, ok = nativeChartInteger(c.take("order"), 0, nativeChartMaxSeries-1)
	if !ok {
		return nil, false
	}
	if c.has("tx") {
		tx := nativeChartChildren(c.take("tx"), d.chart)
		if tx.has("v") {
			value, ok := nativeChartText(tx.take("v"))
			if !ok || nativeChartEncodedString.MatchString(value) || len(utf16.Encode([]rune(value))) > 1024 {
				return nil, false
			}
			result.Title = &value
		} else {
			ref, ok := extractNativeChartReference(tx.take("strRef"), d, false)
			if !ok || ref.Range.Count != 1 {
				return nil, false
			}
			result.TitleReference = ref
		}
		if !tx.done() {
			return nil, false
		}
	}
	return result, true
}
func nativeChartWorkbookSeriesReference(c *nativeChartCursor, d nativeExtractDialect, container string, numeric bool) (*nativeChartReference, bool) {
	source := nativeChartChildren(c.take(container), d.chart)
	name := "strRef"
	if numeric {
		name = "numRef"
	}
	ref, ok := extractNativeChartReference(source.take(name), d, numeric)
	return ref, ok && source.done()
}
func extractNativeChartWorkbookBarSeries(node *nativeXMLNode, d nativeExtractDialect) (*nativeChartWorkbookSeries, bool) {
	c := nativeChartChildren(node, d.chart)
	result, ok := nativeChartWorkbookSeriesHeader(c, d)
	if !ok {
		return nil, false
	}
	var base *nativeXMLNode
	if c.has("spPr") {
		base = c.take("spPr")
	}
	if _, ok = nativeChartInteger(c.take("invertIfNegative"), 0, 0); !ok {
		return nil, false
	}
	points := []*nativeXMLNode{}
	for c.has("dPt") {
		points = append(points, c.take("dPt"))
		if len(points) > nativeChartMaxCategories {
			return nil, false
		}
	}
	result.CategoryReference, ok = nativeChartWorkbookSeriesReference(c, d, "cat", false)
	if !ok {
		return nil, false
	}
	result.ValueReference, ok = nativeChartWorkbookSeriesReference(c, d, "val", true)
	if !ok || !c.done() || result.CategoryReference.Range.Count != result.ValueReference.Range.Count {
		return nil, false
	}
	count := result.ValueReference.Range.Count
	if count < 1 || count > 256 {
		return nil, false
	}
	result.Colors, ok = nativeChartSeriesColors(base, points, int(count), d)
	return result, ok
}
func extractNativeChartWorkbookConnectedSeries(node *nativeXMLNode, d nativeExtractDialect, scatter bool) (*nativeChartWorkbookSeries, bool) {
	c := nativeChartChildren(node, d.chart)
	result, ok := nativeChartWorkbookSeriesHeader(c, d)
	if !ok {
		return nil, false
	}
	result.Color, result.Width, ok = nativeChartConnectedLinePaint(c.take("spPr"), d)
	if !ok {
		return nil, false
	}
	marker := nativeChartChildren(c.take("marker"), d.chart)
	if _, ok = nativeChartToken(marker.take("symbol"), "none"); !ok || !marker.done() {
		return nil, false
	}
	x, y := "cat", "val"
	if scatter {
		x, y = "xVal", "yVal"
	}
	first, ok := nativeChartWorkbookSeriesReference(c, d, x, scatter)
	if !ok {
		return nil, false
	}
	result.ValueReference, ok = nativeChartWorkbookSeriesReference(c, d, y, true)
	if !ok || first.Range.Count != result.ValueReference.Range.Count {
		return nil, false
	}
	if _, ok = nativeChartInteger(c.take("smooth"), 0, 0); !ok || !c.done() {
		return nil, false
	}
	if scatter {
		result.XReference = first
	} else {
		result.CategoryReference = first
	}
	return result, true
}
