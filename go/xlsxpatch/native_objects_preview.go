package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"math"
	"path"
	"sort"
	"strconv"
	"strings"
)

// NativeWorkbookObjectsV1 is a read-only supplement, never mutation authority.
// Chart values come from saved chart caches, not evaluated worksheet formulas.
type NativeWorkbookObjectsV1 struct {
	Protocol      string                 `json:"protocol"`
	Version       int                    `json:"version"`
	PackageSHA256 string                 `json:"package_sha256"`
	Tables        []NativeTablePreviewV1 `json:"tables"`
	Charts        []NativeChartPreviewV1 `json:"charts"`
}
type NativeTablePreviewV1 struct {
	Part          string                    `json:"part"`
	SheetPart     string                    `json:"sheet_part"`
	Name          string                    `json:"name"`
	Ref           string                    `json:"ref"`
	Style         string                    `json:"style"`
	HeaderRows    int                       `json:"header_rows"`
	TotalRows     int                       `json:"total_rows"`
	RowStripes    bool                      `json:"row_stripes"`
	ColumnStripes bool                      `json:"column_stripes"`
	Warnings      []string                  `json:"warnings"`
	FillPreview   *NativeTableFillPreviewV1 `json:"fill_preview,omitempty"`
}

// NativeTableFillPreviewV1 qualifies fills and default-font header text only.
type NativeTableFillPreviewV1 struct {
	Header             string `json:"header"`
	Stripe             string `json:"stripe"`
	Body               string `json:"body"`
	HeaderFontStyleIDs []int  `json:"header_font_style_ids"`
	FillStyleIDs       []int  `json:"fill_style_ids"`
}
type NativeChartPreviewV1 struct {
	Part     string                       `json:"part"`
	Type     string                       `json:"type"`
	Series   []NativeChartSeriesPreviewV1 `json:"series"`
	Warnings []string                     `json:"warnings"`
}
type NativeChartSeriesPreviewV1 struct {
	Name   string     `json:"name"`
	Values []*float64 `json:"values"`
	Labels []string   `json:"labels"`
}

type previewXML struct {
	name     xml.Name
	attrs    []xml.Attr
	children []*previewXML
	text     string
}

func (n *previewXML) attr(name string) string {
	for _, a := range n.attrs {
		if a.Name.Space == "" && a.Name.Local == name {
			return a.Value
		}
	}
	return ""
}
func (n *previewXML) child(name string) *previewXML {
	if n == nil {
		return nil
	}
	var found *previewXML
	for _, c := range n.children {
		if c.name.Space == n.name.Space && c.name.Local == name {
			if found != nil {
				return nil
			}
			found = c
		}
	}
	return found
}
func (n *previewXML) textAt(names ...string) string {
	for _, name := range names {
		n = n.child(name)
		if n == nil {
			return ""
		}
	}
	return n.text
}
func parsePreviewXML(data []byte) (*previewXML, error) {
	if len(data) > 2*1024*1024 {
		return nil, fmt.Errorf("object XML exceeds 2 MiB")
	}
	if _, _, err := preflightNativeCoreXML(data); err != nil {
		return nil, err
	}
	dec := xml.NewDecoder(bytes.NewReader(data))
	var root *previewXML
	var stack []*previewXML
	count := 0
	for {
		t, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch v := t.(type) {
		case xml.StartElement:
			count++
			if count > 20000 || len(stack) > 64 {
				return nil, fmt.Errorf("object XML structure exceeds bound")
			}
			n := &previewXML{name: v.Name, attrs: v.Attr}
			if len(stack) > 0 {
				p := stack[len(stack)-1]
				p.children = append(p.children, n)
			} else {
				root = n
			}
			stack = append(stack, n)
		case xml.EndElement:
			stack = stack[:len(stack)-1]
		case xml.CharData:
			if len(stack) > 0 {
				n := stack[len(stack)-1]
				n.text += string(v)
				if len(n.text) > 32767 {
					return nil, fmt.Errorf("object text exceeds bound")
				}
			}
		}
	}
	if root == nil {
		return nil, fmt.Errorf("empty object XML")
	}
	return root, nil
}

// InspectNativeWorkbookObjectsV1 reuses native ZIP/OPC/XML validation and never
// follows external references. All object parts are retained in the original.
func InspectNativeWorkbookObjectsV1(data []byte) (*NativeWorkbookObjectsV1, error) {
	workbook, err := ExtractNativeWorkbookV2(data)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return nil, err
	}
	result := &NativeWorkbookObjectsV1{Protocol: "injoffice.xlsx.preview-objects", Version: 1, PackageSHA256: workbook.Source.PackageSHA256, Tables: []NativeTablePreviewV1{}, Charts: []NativeChartPreviewV1{}}
	owners := map[string]string{}
	for _, sheet := range workbook.Sheets {
		relPart := path.Join(path.Dir(sheet.PartName), "_rels", path.Base(sheet.PartName)+".rels")
		if raw, ok := pkg.files[relPart]; ok {
			rels, e := parseRoutingRelationships(raw)
			if e != nil {
				return nil, e
			}
			for _, rel := range rels {
				if rel.targetMode == "External" || (rel.relType != "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" && rel.relType != "http://purl.oclc.org/ooxml/officeDocument/relationships/table") {
					continue
				}
				worksheet, e := parsePreviewXML(pkg.files[sheet.PartName])
				if e != nil {
					return nil, e
				}
				active := false
				tableParts := worksheet.child("tableParts")
				if tableParts != nil {
					for _, tablePart := range tableParts.children {
						if tablePart.name.Space != worksheet.name.Space || tablePart.name.Local != "tablePart" {
							continue
						}
						for _, a := range tablePart.attrs {
							if (a.Name.Space == officeRelNamespaceTransitional || a.Name.Space == officeRelNamespaceStrict) && a.Name.Local == "id" && a.Value == rel.id {
								active = true
							}
						}
					}
				}
				if !active {
					continue
				}
				target, e := resolveNativeRelationshipTarget(sheet.PartName, path.Dir(sheet.PartName), rel.target)
				if e != nil {
					return nil, e
				}
				actual, _, ok := pkg.index.lookupResolved(target)
				if ok {
					if previous := owners[actual]; previous != "" && previous != sheet.PartName {
						return nil, fmt.Errorf("table has ambiguous sheet ownership")
					}
					owners[actual] = sheet.PartName
				}
			}
		}
	}
	parts := make([]string, 0)
	for part, kind := range pkg.contentTypes.overrides {
		if kind == "application/vnd.openxmlformats-officedocument.drawingml.chart+xml" || kind == "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" {
			parts = append(parts, part)
		}
	}
	sort.Strings(parts)
	if len(parts) > 64 {
		return nil, fmt.Errorf("preview supports at most 64 chart/table parts")
	}
	for _, part := range parts {
		actual, _, ok := pkg.index.lookupResolved(part)
		if !ok {
			return nil, fmt.Errorf("missing object part")
		}
		root, e := parsePreviewXML(pkg.files[actual])
		if e != nil {
			return nil, fmt.Errorf("preview %s: %w", actual, e)
		}
		switch root.name {
		case xml.Name{Space: "http://schemas.openxmlformats.org/spreadsheetml/2006/main", Local: "table"}, xml.Name{Space: "http://purl.oclc.org/ooxml/spreadsheetml/main", Local: "table"}:
			table := NativeTablePreviewV1{Part: actual, SheetPart: owners[actual], Name: root.attr("displayName"), Ref: root.attr("ref"), HeaderRows: 1, Warnings: []string{"Table style metadata is source-derived; Office built-in style rendering is not yet qualified."}}
			if v := root.attr("headerRowCount"); v != "" {
				table.HeaderRows, e = strconv.Atoi(v)
				if e != nil || table.HeaderRows < 0 || table.HeaderRows > 1 {
					return nil, fmt.Errorf("invalid table header count")
				}
			}
			if v := root.attr("totalsRowCount"); v != "" {
				table.TotalRows, e = strconv.Atoi(v)
				if e != nil || table.TotalRows < 0 || table.TotalRows > 1 {
					return nil, fmt.Errorf("invalid table totals count")
				}
			}
			if style := root.child("tableStyleInfo"); style != nil {
				table.Style = style.attr("name")
				table.RowStripes = style.attr("showRowStripes") == "1" || style.attr("showRowStripes") == "true"
				table.ColumnStripes = style.attr("showColumnStripes") == "1" || style.attr("showColumnStripes") == "true"
			}
			if table.SheetPart == "" {
				table.Warnings = append(table.Warnings, "Table has no unambiguous worksheet relationship; no grid styling is applied.")
			}
			qualifyNativeTableFillPreview(pkg, root, &table)
			result.Tables = append(result.Tables, table)
		case xml.Name{Space: "http://schemas.openxmlformats.org/drawingml/2006/chart", Local: "chartSpace"}, xml.Name{Space: "http://purl.oclc.org/ooxml/drawingml/chart", Local: "chartSpace"}:
			result.Charts = append(result.Charts, previewChart(root, actual))
		default:
			return nil, fmt.Errorf("unexpected chart/table root in %s", actual)
		}
	}
	if !nativeTableStyleIDsWithinBudget(result.Tables) {
		return nil, fmt.Errorf("table preview style IDs exceed cumulative limit 16384")
	}
	points := 0
	for _, chart := range result.Charts {
		for _, series := range chart.Series {
			points += len(series.Values)
			if points > 65536 {
				return nil, fmt.Errorf("object cache points exceed cumulative limit 65536")
			}
		}
	}
	encoded, err := json.Marshal(result)
	if err != nil {
		return nil, err
	}
	if len(encoded) > 8*1024*1024 {
		return nil, fmt.Errorf("object preview exceeds 8 MiB JSON bound")
	}
	return result, nil
}

func previewChart(root *previewXML, part string) NativeChartPreviewV1 {
	result := NativeChartPreviewV1{Part: part, Type: "unsupported", Series: []NativeChartSeriesPreviewV1{}, Warnings: []string{"Saved chart caches may be stale. No formulas or external links are evaluated.", "Data preview only: Office chart styling, axes, titles, hyperlinks and drawing placement are not reproduced.", "Category labels are not interpreted; point numbers are saved cache indices, not category names."}}
	plot := root.child("chart").child("plotArea")
	if plot == nil {
		return result
	}
	var chart *previewXML
	plots := 0
	for _, n := range plot.children {
		if n.name.Space == plot.name.Space && strings.HasSuffix(n.name.Local, "Chart") {
			plots++
			chart = n
		}
	}
	if plots != 1 || chart == nil || chart.name.Local != "barChart" {
		result.Warnings = append(result.Warnings, "Only a single 2D clustered column/bar chart is supported.")
		return result
	}
	if grouping := chart.child("grouping"); grouping == nil || grouping.attr("val") != "clustered" {
		result.Warnings = append(result.Warnings, "Grouping is not explicitly clustered.")
		return result
	}
	direction := chart.child("barDir")
	if direction == nil || (direction.attr("val") != "col" && direction.attr("val") != "bar") {
		return result
	}
	result.Type = direction.attr("val")
	for _, series := range chart.children {
		if series.name.Space != chart.name.Space || series.name.Local != "ser" {
			continue
		}
		if len(result.Series) >= 32 {
			result.Type = "unsupported"
			result.Series = []NativeChartSeriesPreviewV1{}
			result.Warnings = append(result.Warnings, "Series count exceeds 32.")
			return result
		}
		cache := series.child("val").child("numRef").child("numCache")
		values, ok := previewNumericCache(cache)
		if !ok {
			result.Type = "unsupported"
			result.Series = []NativeChartSeriesPreviewV1{}
			result.Warnings = append(result.Warnings, "A complete bounded saved numeric cache is required for every series.")
			return result
		}
		name := series.textAt("tx", "v")
		if name == "" {
			name = series.textAt("tx", "strRef", "strCache", "pt", "v")
		}
		result.Series = append(result.Series, NativeChartSeriesPreviewV1{Name: name, Values: values, Labels: []string{}})
	}
	if len(result.Series) == 0 {
		result.Type = "unsupported"
	}
	return result
}

func previewNumericCache(cache *previewXML) ([]*float64, bool) {
	if cache == nil {
		return nil, false
	}
	countNode := cache.child("ptCount")
	if countNode == nil {
		return nil, false
	}
	count, err := strconv.Atoi(countNode.attr("val"))
	if err != nil || count < 1 || count > 1024 {
		return nil, false
	}
	values := make([]*float64, count)
	seen := map[int]bool{}
	for _, pt := range cache.children {
		if pt.name.Space != cache.name.Space || pt.name.Local != "pt" {
			continue
		}
		idx, e := strconv.Atoi(pt.attr("idx"))
		if e != nil || idx < 0 || idx >= count || seen[idx] {
			return nil, false
		}
		seen[idx] = true
		v, e := finiteNativeFloat(strings.TrimSpace(pt.textAt("v")), -math.MaxFloat64, math.MaxFloat64)
		if e != nil || math.IsNaN(v) || math.IsInf(v, 0) {
			return nil, false
		}
		values[idx] = &v
	}
	return values, len(seen) == count
}
