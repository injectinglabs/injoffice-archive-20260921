package pptxpatch

import (
	"encoding/xml"
	"regexp"
	"unicode/utf16"
)

var nativeChartEncodedString = regexp.MustCompile(`(?i)_x[0-9a-f]{4}_`)

const nativeChartMaxSeries = 16
const nativeChartMaxCategories = 256
const nativeChartMaxCategoryUnits = 32768

type nativeChartLiteralSeries struct {
	Index      int64
	Order      int64
	Title      *string
	Categories []string
	Values     []string
	Colors     []string
}

// Every consumed child is namespace-qualified and source-ordered. Unexpected
// children are never searched past or silently discarded.
type nativeChartCursor struct {
	node  *nativeXMLNode
	ns    string
	index int
	valid bool
}

func nativeChartChildren(node *nativeXMLNode, ns string) *nativeChartCursor {
	return &nativeChartCursor{node: node, ns: ns, valid: node != nil && requireOnlyNativeAttrs(node) == nil && onlyNativeXMLSpace(node.Text)}
}
func (c *nativeChartCursor) take(name string) *nativeXMLNode {
	if !c.valid || c.index >= len(c.node.Children) || c.node.Children[c.index].Name != (xml.Name{Space: c.ns, Local: name}) {
		c.valid = false
		return nil
	}
	n := c.node.Children[c.index]
	c.index++
	return n
}
func (c *nativeChartCursor) has(name string) bool {
	return c.valid && c.index < len(c.node.Children) && c.node.Children[c.index].Name == (xml.Name{Space: c.ns, Local: name})
}
func (c *nativeChartCursor) done() bool { return c.valid && c.index == len(c.node.Children) }
func nativeChartAttribute(node *nativeXMLNode) (string, bool) {
	if node == nil || requireOnlyNativeAttrs(node, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(node) != nil {
		return "", false
	}
	return exactNativeAttr(node, "", "val")
}
func nativeChartInteger(node *nativeXMLNode, low, high int64) (int64, bool) {
	raw, ok := nativeChartAttribute(node)
	if !ok {
		return 0, false
	}
	value, err := parseCanonicalNativeInt(raw, low, high)
	return value, err == nil
}
func nativeChartText(node *nativeXMLNode) (string, bool) {
	if node == nil || requireOnlyNativeAttrs(node) != nil || len(node.Children) != 0 {
		return "", false
	}
	return node.Text, true
}
func nativeChartLiteralPoints(node *nativeXMLNode, d nativeExtractDialect, numeric bool) ([]string, bool) {
	c := nativeChartChildren(node, d.chart)
	if numeric {
		text, ok := nativeChartText(c.take("formatCode"))
		if !ok || text != "General" {
			return nil, false
		}
	}
	count, ok := nativeChartInteger(c.take("ptCount"), 1, nativeChartMaxCategories)
	if !ok || !c.valid || count != int64(len(node.Children)-c.index) {
		return nil, false
	}
	values := make([]string, len(node.Children)-c.index)
	seen := make([]bool, len(values))
	units := 0
	for c.has("pt") {
		pt := c.take("pt")
		if requireOnlyNativeAttrs(pt, xml.Name{Local: "idx"}) != nil || !onlyNativeXMLSpace(pt.Text) || len(pt.Children) != 1 || pt.Children[0].Name != (xml.Name{Space: d.chart, Local: "v"}) {
			return nil, false
		}
		raw, ok := exactNativeAttr(pt, "", "idx")
		if !ok {
			return nil, false
		}
		index, err := parseCanonicalNativeInt(raw, 0, count-1)
		if err != nil || seen[index] {
			return nil, false
		}
		seen[index] = true
		text, ok := nativeChartText(pt.Children[0])
		if !ok {
			return nil, false
		}
		if numeric {
			if _, err := parseNativeChartDecimal(text); err != nil {
				return nil, false
			}
		} else {
			if nativeChartEncodedString.MatchString(text) {
				return nil, false
			}
			units += len(utf16.Encode([]rune(text)))
			if units > nativeChartMaxCategoryUnits {
				return nil, false
			}
		}
		values[index] = text
	}
	return values, c.done()
}
func nativeChartSeriesColors(base *nativeXMLNode, points []*nativeXMLNode, count int, d nativeExtractDialect) ([]string, bool) {
	if count < 1 || count > nativeChartMaxCategories || len(points) > count {
		return nil, false
	}
	colors := make([]string, count)
	if base != nil {
		color, ok := nativeLiteralPieSolidPaint(base, d)
		if !ok {
			return nil, false
		}
		for i := range colors {
			colors[i] = color
		}
	}
	seen := make([]bool, count)
	for _, point := range points {
		p := nativeChartChildren(point, d.chart)
		index, ok := nativeChartInteger(p.take("idx"), 0, int64(count)-1)
		if !ok || seen[index] {
			return nil, false
		}
		seen[index] = true
		color, ok := nativeLiteralPieSolidPaint(p.take("spPr"), d)
		if !ok || !p.done() {
			return nil, false
		}
		colors[index] = color
	}
	for _, color := range colors {
		if color == "" {
			return nil, false
		}
	}
	return colors, true
}
func extractNativeChartLiteralSeries(node *nativeXMLNode, d nativeExtractDialect) (*nativeChartLiteralSeries, bool) {
	c := nativeChartChildren(node, d.chart)
	index, ok := nativeChartInteger(c.take("idx"), 0, 4294967295)
	if !ok {
		return nil, false
	}
	order, ok := nativeChartInteger(c.take("order"), 0, nativeChartMaxSeries-1)
	if !ok {
		return nil, false
	}
	result := &nativeChartLiteralSeries{Index: index, Order: order}
	if c.has("tx") {
		tx := nativeChartChildren(c.take("tx"), d.chart)
		text, ok := nativeChartText(tx.take("v"))
		if !ok || !tx.done() || nativeChartEncodedString.MatchString(text) || len(utf16.Encode([]rune(text))) > 1024 {
			return nil, false
		}
		result.Title = &text
	}
	var base *nativeXMLNode
	if c.has("spPr") {
		base = c.take("spPr")
	}
	// An explicit false is required even for currently nonnegative values, so a
	// future signed-data extension cannot silently inherit an inverted paint.
	invert, ok := nativeChartInteger(c.take("invertIfNegative"), 0, 0)
	if !ok || invert != 0 {
		return nil, false
	}
	points := []*nativeXMLNode{}
	for c.has("dPt") {
		points = append(points, c.take("dPt"))
		if len(points) > nativeChartMaxCategories {
			return nil, false
		}
	}
	cat := nativeChartChildren(c.take("cat"), d.chart)
	categories, ok := nativeChartLiteralPoints(cat.take("strLit"), d, false)
	if !ok || !cat.done() {
		return nil, false
	}
	val := nativeChartChildren(c.take("val"), d.chart)
	values, ok := nativeChartLiteralPoints(val.take("numLit"), d, true)
	if !ok || !val.done() || len(categories) != len(values) || !c.done() {
		return nil, false
	}
	colors, ok := nativeChartSeriesColors(base, points, len(values), d)
	if !ok {
		return nil, false
	}
	result.Categories = categories
	result.Values = values
	result.Colors = colors
	return result, true
}
