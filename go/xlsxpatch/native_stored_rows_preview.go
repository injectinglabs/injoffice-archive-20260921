package xlsxpatch

import (
	"encoding/xml"
	"math"
	"regexp"
	"strconv"
	"strings"
)

const nativeStoredRowPreviewLimit = 32
const nativeRowDescentNamespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"

var nativeStoredRowDecimal = regexp.MustCompile(`^[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?$`)

// Only the first 32 rows are projected. This is stored row geometry, not font
// metrics, column sizing, baseline positioning, or automatic row fitting.
type NativeStoredRowGeometryV1 struct {
	SheetPart string              `json:"sheet_part"`
	Rows      []NativeStoredRowV1 `json:"rows"`
	Warnings  []string            `json:"warnings"`
}
type NativeStoredRowV1 struct {
	Row          int     `json:"row"`
	HeightPoints float64 `json:"height_points"`
	Hidden       bool    `json:"hidden"`
}

func previewNativeStoredRows(raw []byte, part string) NativeStoredRowGeometryV1 {
	result := NativeStoredRowGeometryV1{SheetPart: part, Rows: []NativeStoredRowV1{}, Warnings: []string{}}
	unavailable := func() NativeStoredRowGeometryV1 {
		result.Rows = []NativeStoredRowV1{}
		result.Warnings = []string{"Stored row heights unavailable: source dimensions are missing, automatic, ambiguous or outside this bounded preview."}
		return result
	}
	root, err := parsePreviewXML(raw)
	if err != nil || (root.name != (xml.Name{Space: spreadsheetMLTransitional, Local: "worksheet"}) && root.name != (xml.Name{Space: spreadsheetMLStrict, Local: "worksheet"})) {
		return unavailable()
	}
	format := root.child("sheetFormatPr")
	data := root.child("sheetData")
	if format == nil || data == nil || len(data.attrs) != 0 || strings.TrimSpace(data.text) != "" || len(format.children) != 0 || strings.TrimSpace(format.text) != "" {
		return unavailable()
	}
	descent := false
	for _, attr := range format.attrs {
		if isPreviewNamespaceDeclaration(attr) {
			continue
		}
		if attr.Name.Space == nativeRowDescentNamespace && attr.Name.Local == "dyDescent" {
			if _, ok := boundedPreviewRowNumber(attr.Value, 409); !ok {
				return unavailable()
			}
			descent = true
			continue
		}
		if attr.Name.Space != "" {
			return unavailable()
		}
		switch attr.Name.Local {
		case "defaultRowHeight":
		case "customHeight", "zeroHeight":
			if _, ok := previewRowBool(attr.Value); !ok {
				return unavailable()
			}
		case "baseColWidth", "defaultColWidth": // Column metrics remain independently unqualified.
		default:
			return unavailable()
		}
	}
	height, ok := boundedPreviewRowNumber(format.attr("defaultRowHeight"), 409)
	custom, _ := previewRowBool(format.attr("customHeight"))
	zero, _ := previewRowBool(format.attr("zeroHeight"))
	if !ok || zero || (!custom && !descent) {
		return unavailable()
	}
	for row := 0; row < nativeStoredRowPreviewLimit; row++ {
		result.Rows = append(result.Rows, NativeStoredRowV1{Row: row, HeightPoints: height, Hidden: height == 0})
	}
	seen := map[int]bool{}
	for _, row := range data.children {
		if row.name.Space != root.name.Space || row.name.Local != "row" {
			return unavailable()
		}
		index, e := strconv.Atoi(row.attr("r"))
		if e != nil || index < 1 || index > 1048576 || strconv.Itoa(index) != row.attr("r") || seen[index] {
			return unavailable()
		}
		seen[index] = true
		if index > nativeStoredRowPreviewLimit {
			continue
		}
		if strings.TrimSpace(row.text) != "" {
			return unavailable()
		}
		rowDescent := false
		for _, attr := range row.attrs {
			if isPreviewNamespaceDeclaration(attr) {
				continue
			}
			if attr.Name.Space == nativeRowDescentNamespace && attr.Name.Local == "dyDescent" {
				if _, ok := boundedPreviewRowNumber(attr.Value, 409); !ok {
					return unavailable()
				}
				rowDescent = true
				descent = true
				continue
			}
			if attr.Name.Space != "" {
				return unavailable()
			}
			switch attr.Name.Local {
			case "r", "ht":
			case "hidden", "customHeight":
				if _, ok := previewRowBool(attr.Value); !ok {
					return unavailable()
				}
			case "spans":
				if !validPreviewRowSpans(attr.Value) {
					return unavailable()
				}
			default:
				return unavailable()
			}
		}
		rowHeight := height
		if rawCustom := row.attr("customHeight"); rawCustom != "" {
			fixed, _ := previewRowBool(rawCustom)
			if !fixed && !rowDescent {
				return unavailable()
			}
		}
		if rawHeight := row.attr("ht"); rawHeight != "" {
			var valid bool
			rowHeight, valid = boundedPreviewRowNumber(rawHeight, 409)
			fixed, _ := previewRowBool(row.attr("customHeight"))
			if !valid || (!fixed && !rowDescent) {
				return unavailable()
			}
		}
		hidden, _ := previewRowBool(row.attr("hidden"))
		result.Rows[index-1] = NativeStoredRowV1{Row: index - 1, HeightPoints: rowHeight, Hidden: hidden || rowHeight == 0}
	}
	result.Warnings = []string{"Only stored heights and visibility for the first 32 rows are applied. Column widths, font metrics and automatic row fitting are not reproduced."}
	if descent {
		result.Warnings = append(result.Warnings, "Source dyDescent fixes stored row height; its text-baseline positioning remains unmodeled.")
	}
	return result
}
func isPreviewNamespaceDeclaration(attr xml.Attr) bool {
	return attr.Name.Space == "xmlns" || (attr.Name.Space == "" && attr.Name.Local == "xmlns")
}
func boundedPreviewRowNumber(raw string, max float64) (float64, bool) {
	if len(raw) > 64 || !nativeStoredRowDecimal.MatchString(raw) {
		return 0, false
	}
	v, e := strconv.ParseFloat(raw, 64)
	return v, e == nil && !math.IsNaN(v) && !math.IsInf(v, 0) && v >= 0 && v <= max
}
func previewRowBool(raw string) (bool, bool) {
	switch raw {
	case "1", "true":
		return true, true
	case "0", "false":
		return false, true
	}
	return false, false
}
func validPreviewRowSpans(raw string) bool {
	if len(raw) > 256 {
		return false
	}
	last := 0
	for _, span := range strings.Fields(raw) {
		parts := strings.Split(span, ":")
		if len(parts) != 2 {
			return false
		}
		first, e := strconv.Atoi(parts[0])
		end, f := strconv.Atoi(parts[1])
		if e != nil || f != nil || first <= last || end < first || end > 16384 || strconv.Itoa(first) != parts[0] || strconv.Itoa(end) != parts[1] {
			return false
		}
		last = end
	}
	return last > 0
}
