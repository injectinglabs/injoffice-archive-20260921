package pptxpatch

import (
	"encoding/xml"
	"regexp"
	"strings"
	"unicode"
	"unicode/utf16"
	"unicode/utf8"
)

// A reference remains a source expression. It is never rewritten into numLit
// or populated from the chart cache. Workbook resolution is a separate step.
type nativeChartReference struct {
	Kind         string
	Formula      string
	Range        nativeChartCellRange
	CachePresent bool
}
type nativeChartCellRange struct {
	Sheet                                    string
	StartRow, StartColumn, EndRow, EndColumn int64 // zero-based, inclusive
	Count                                    int64
}

var nativeChartDirectRange = regexp.MustCompile(`^(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))!\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Za-z]{1,3})\$?([1-9][0-9]{0,6}))?$`)

func parseNativeChartCellRange(formula string) (nativeChartCellRange, bool) {
	var result nativeChartCellRange
	if len(formula) > 1024 || !utf8.ValidString(formula) {
		return result, false
	}
	match := nativeChartDirectRange.FindStringSubmatch(formula)
	if match == nil {
		return result, false
	}
	result.Sheet = match[2]
	if match[1] != "" {
		result.Sheet = strings.ReplaceAll(match[1], "''", "'")
	}
	if result.Sheet == "" || len(result.Sheet) > 256 || strings.ContainsAny(result.Sheet, "[]:*?/\\") || strings.IndexFunc(result.Sheet, unicode.IsControl) >= 0 || strings.HasPrefix(result.Sheet, "'") || strings.HasSuffix(result.Sheet, "'") {
		return result, false
	}
	column := func(raw string) (int64, bool) {
		var value int64
		for _, c := range strings.ToUpper(raw) {
			value = value*26 + int64(c-'A'+1)
		}
		return value - 1, value >= 1 && value <= 16384
	}
	var ok bool
	result.StartColumn, ok = column(match[3])
	if !ok {
		return result, false
	}
	row, e := parseCanonicalNativeInt(match[4], 1, 1048576)
	if e != nil {
		return result, false
	}
	result.StartRow = row - 1
	result.EndColumn, result.EndRow = result.StartColumn, result.StartRow
	if match[5] != "" {
		result.EndColumn, ok = column(match[5])
		if !ok {
			return result, false
		}
		row, e = parseCanonicalNativeInt(match[6], 1, 1048576)
		if e != nil {
			return result, false
		}
		result.EndRow = row - 1
	}
	if result.EndRow < result.StartRow || result.EndColumn < result.StartColumn || result.EndRow != result.StartRow && result.EndColumn != result.StartColumn {
		return result, false
	}
	result.Count = (result.EndRow - result.StartRow + 1) * (result.EndColumn - result.StartColumn + 1)
	return result, result.Count <= nativeChartMaxCategories
}

func extractNativeChartReference(node *nativeXMLNode, d nativeExtractDialect, numeric bool) (*nativeChartReference, bool) {
	kind, cache := "strRef", "strCache"
	if numeric {
		kind, cache = "numRef", "numCache"
	}
	if node == nil || node.Name != (xml.Name{Space: d.chart, Local: kind}) {
		return nil, false
	}
	cursor := nativeChartChildren(node, d.chart)
	formula, ok := nativeChartText(cursor.take("f"))
	if !ok {
		return nil, false
	}
	area, ok := parseNativeChartCellRange(formula)
	if !ok {
		return nil, false
	}
	result := &nativeChartReference{Kind: kind, Formula: formula, Range: area}
	if cursor.has(cache) {
		if !nativeChartIgnoredCache(cursor.take(cache), d, numeric) {
			return nil, false
		}
		result.CachePresent = true
	}
	return result, cursor.done()
}

// Cache structure is bounded and qualified, but neither values, ordering,
// point count nor number format are used to resolve the reference. A stale
// cache may deliberately disagree with the authoritative workbook cells.
func nativeChartIgnoredCache(node *nativeXMLNode, d nativeExtractDialect, numeric bool) bool {
	c := nativeChartChildren(node, d.chart)
	units := 0
	if numeric && c.has("formatCode") {
		format, ok := nativeChartText(c.take("formatCode"))
		if !ok || len(format) > 1024 {
			return false
		}
	}
	if c.has("ptCount") {
		if _, ok := nativeChartInteger(c.take("ptCount"), 0, 4294967295); !ok {
			return false
		}
	}
	points := 0
	for c.has("pt") {
		points++
		if points > nativeChartMaxCategories {
			return false
		}
		point := c.take("pt")
		attrs := []xml.Name{{Local: "idx"}}
		if numeric {
			attrs = append(attrs, xml.Name{Local: "formatCode"})
		}
		if requireOnlyNativeAttrs(point, attrs...) != nil || !onlyNativeXMLSpace(point.Text) || len(point.Children) != 1 || point.Children[0].Name != (xml.Name{Space: d.chart, Local: "v"}) {
			return false
		}
		index, ok := exactNativeAttr(point, "", "idx")
		if !ok {
			return false
		}
		if _, e := parseCanonicalNativeInt(index, 0, 4294967295); e != nil {
			return false
		}
		if format, ok := exactNativeAttr(point, "", "formatCode"); ok && len(format) > 1024 {
			return false
		}
		text, ok := nativeChartText(point.Children[0])
		if !ok {
			return false
		}
		units += len(utf16.Encode([]rune(text)))
		if units > nativeChartMaxCategoryUnits {
			return false
		}
	}
	return c.done()
}
