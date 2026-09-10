package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"
)

func chartSpecWithCaches(read func(string) (string, bool), spec ChartWriteSpec) ChartWriteSpec {
	// Clone the slice: caches are source-derived, never caller-supplied evidence.
	spec.Series = append([]WriteSeries(nil), spec.Series...)
	type key struct {
		ref             string
		numeric, single bool
	}
	memo := map[key]string{}
	cache := func(ref string, numeric, single bool) string {
		k := key{ref, numeric, single}
		if value, ok := memo[k]; ok {
			return value
		}
		value := chartReferenceCache(read, ref, numeric, single)
		memo[k] = value
		return value
	}
	for i := range spec.Series {
		s := &spec.Series[i]
		s.nameCache = cache(s.NameRef, false, true)
		s.categoryCache = cache(s.CategoriesRef, spec.Type == "scatter", false)
		s.valueCache = cache(s.ValuesRef, true, false)
	}
	return spec
}

// Cache only bounded, one-dimensional literal ranges. Formula caches may be
// stale, so even a formula with a cached value is deliberately not copied.
// Unsupported refs retain their formula and omit the cache (Excel can recalc).
func chartReferenceCache(read func(string) (string, bool), ref string, numeric, single bool) string {
	split := strings.LastIndex(ref, "!")
	if split < 1 {
		return ""
	}
	sheet := ref[:split]
	if strings.HasPrefix(sheet, "'") && strings.HasSuffix(sheet, "'") {
		sheet = strings.ReplaceAll(sheet[1:len(sheet)-1], "''", "'")
	}
	if strings.ContainsAny(sheet, "[]:") {
		return ""
	}
	r1, c1, r2, c2, err := parseDimensionReference(strings.ReplaceAll(ref[split+1:], "$", ""))
	if err != nil || (r1 != r2 && c1 != c2) {
		return ""
	}
	count := (r2 - r1 + 1) * (c2 - c1 + 1)
	if count < 1 || count > 10000 || (single && count != 1) {
		return ""
	}
	part, err := worksheetPartFor(read, sheet)
	if err != nil {
		return ""
	}
	source, ok := read(part)
	if !ok || len(source) > 32*1024*1024 {
		return ""
	}
	type text struct {
		Text string `xml:"t"`
		Runs []struct {
			Text string `xml:"t"`
		} `xml:"r"`
	}
	type cell struct {
		Ref     string `xml:"r,attr"`
		Type    string `xml:"t,attr"`
		Formula *struct {
			Ref string `xml:"ref,attr"`
		} `xml:"f"`
		Value  *string `xml:"v"`
		Inline text    `xml:"is"`
	}
	var worksheet struct {
		Rows []struct {
			Cells []cell `xml:"c"`
		} `xml:"sheetData>row"`
	}
	if xml.Unmarshal([]byte(source), &worksheet) != nil {
		return ""
	}
	cells := map[string]cell{}
	for _, row := range worksheet.Rows {
		for _, item := range row.Cells {
			// Shared/array formula dependents may carry cached values without
			// their own formula element. Never mistake these for literal cells.
			if item.Formula != nil && item.Formula.Ref != "" {
				fr1, fc1, fr2, fc2, err := parseDimensionReference(item.Formula.Ref)
				if err != nil || (fr1 <= r2 && fr2 >= r1 && fc1 <= c2 && fc2 >= c1) {
					return ""
				}
			}
			if _, duplicate := cells[item.Ref]; duplicate {
				return ""
			}
			cells[item.Ref] = item
		}
	}
	var shared struct {
		Items []text `xml:"si"`
	}
	sharedLoaded := false
	textValue := func(value text) string {
		s := value.Text
		for _, run := range value.Runs {
			s += run.Text
		}
		return s
	}
	var points strings.Builder
	for i := 0; i < count; i++ {
		row, col := r1, c1
		if r1 == r2 {
			col += i
		} else {
			row += i
		}
		item, exists := cells[cellReference(row, col)]
		if !exists {
			if single {
				return ""
			}
			continue
		}
		if item.Formula != nil {
			return ""
		}
		value := ""
		switch item.Type {
		case "", "n":
			if item.Value == nil {
				continue
			}
			value = *item.Value
			n, err := strconv.ParseFloat(value, 64)
			if err != nil || math.IsNaN(n) || math.IsInf(n, 0) {
				return ""
			}
			// Even omitted style or style 0 can have a custom number format or
			// inherit a row/column style. Without full effective-style formatting,
			// raw numbers are not proven display strings for category/name caches.
			if !numeric {
				return ""
			}
		case "inlineStr":
			if numeric {
				return ""
			}
			value = textValue(item.Inline)
		case "s":
			if numeric || item.Value == nil {
				return ""
			}
			if !sharedLoaded {
				raw, ok := chartSharedStringsPart(read)
				if !ok || len(raw) > 32*1024*1024 || xml.Unmarshal([]byte(raw), &shared) != nil {
					return ""
				}
				sharedLoaded = true
			}
			index, err := strconv.Atoi(*item.Value)
			if err != nil || index < 0 || index >= len(shared.Items) {
				return ""
			}
			value = textValue(shared.Items[index])
		default:
			return ""
		}
		fmt.Fprintf(&points, `<c:pt idx="%d"><c:v>%s</c:v></c:pt>`, i, esc(value))
	}
	tag, format := "strCache", ""
	if numeric {
		tag, format = "numCache", "<c:formatCode>General</c:formatCode>"
	}
	return fmt.Sprintf(`<c:%s>%s<c:ptCount val="%d"/>%s</c:%s>`, tag, format, count, points.String(), tag)
}

func chartSharedStringsPart(read func(string) (string, bool)) (string, bool) {
	rels, ok := read("xl/_rels/workbook.xml.rels")
	if !ok || len(rels) > 32*1024*1024 {
		return "", false
	}
	routes, err := parseRoutingRelationships([]byte(rels))
	if err != nil {
		return "", false
	}
	id := ""
	for _, route := range routes {
		if route.relType != relTypeSharedStringsTransitional && route.relType != relTypeSharedStringsStrict {
			continue
		}
		if id != "" {
			return "", false
		}
		id = route.id
	}
	if id == "" {
		return "", false
	}
	part, err := relTargetOfType(rels, id, "xl", relTypeSharedStringsTransitional, relTypeSharedStringsStrict)
	if err != nil {
		return "", false
	}
	return read(part)
}
