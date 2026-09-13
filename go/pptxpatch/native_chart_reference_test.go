package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeChartDirectCellRange(t *testing.T) {
	for _, test := range []struct {
		formula, sheet                  string
		row, col, endRow, endCol, count int64
	}{
		{"Sheet1!$B$2:$B$4", "Sheet1", 1, 1, 3, 1, 3},
		{"'Quarter 1'!A1:C1", "Quarter 1", 0, 0, 0, 2, 3},
		{"'O''Brien ! Q'!$a1:$c1", "O'Brien ! Q", 0, 0, 0, 2, 3},
		{"'Δεδομένα'!XFD1048576", "Δεδομένα", 1048575, 16383, 1048575, 16383, 1},
		{"1!A1:A256", "1", 0, 0, 255, 0, 256},
	} {
		area, ok := parseNativeChartCellRange(test.formula)
		if !ok || area.Sheet != test.sheet || area.StartRow != test.row || area.StartColumn != test.col || area.EndRow != test.endRow || area.EndColumn != test.endCol || area.Count != test.count {
			t.Fatalf("%s: %+v %v", test.formula, area, ok)
		}
	}
	for _, formula := range []string{"A1:A2", "Sheet!A0", "Sheet!A01", "Sheet!XFE1", "Sheet!A1048577", "Sheet!A1:A257", "Sheet!A2:A1", "Sheet!B1:A1", "Sheet!A1:B2", "[1]Sheet!A1", "'[Book.xlsx]Sheet'!A1", "SUM(Sheet!A1:A2)", "Sheet!Name", "Sheet!A:A", "Sheet!1:2", "Sheet1:Sheet2!A1", "Sheet!A1,Sheet!B1", "Sheet!A1\n", "'Bad/Name'!A1", "'\n'!A1", "'''Bad'!A1", "Sheet!$A$1#", "Sheet!R1C1", "=Sheet!A1"} {
		if _, ok := parseNativeChartCellRange(formula); ok {
			t.Fatalf("unsupported reference accepted %q", formula)
		}
	}
}
func TestNativeChartReferenceNeverResolvesCaches(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for _, numeric := range []bool{false, true} {
			kind, cache := "strRef", "strCache"
			if numeric {
				kind, cache = "numRef", "numCache"
			}
			source := `<c:` + kind + ` xmlns:c="` + d.chart + `"><c:f>'Source Sheet'!$B$2:$B$3</c:f><c:` + cache + `><c:ptCount val="999"/><c:pt idx="900"><c:v>STALE</c:v></c:pt><c:pt idx="900"><c:v>NaN</c:v></c:pt></c:` + cache + `></c:` + kind + `>`
			node, e := parseNativeXML([]byte(source), "chart.xml")
			if e != nil {
				t.Fatal(e)
			}
			ref, ok := extractNativeChartReference(node, d, numeric)
			if !ok || !ref.CachePresent || ref.Range.Count != 2 || ref.Formula != "'Source Sheet'!$B$2:$B$3" {
				t.Fatalf("reference metadata missing: %+v", ref)
			}
			absent := strings.Replace(source, `<c:`+cache+`><c:ptCount val="999"/><c:pt idx="900"><c:v>STALE</c:v></c:pt><c:pt idx="900"><c:v>NaN</c:v></c:pt></c:`+cache+`>`, "", 1)
			node, _ = parseNativeXML([]byte(absent), "chart.xml")
			ref, ok = extractNativeChartReference(node, d, numeric)
			if !ok || ref.CachePresent {
				t.Fatal("cache absence should not invent values")
			}
			for _, pair := range [][2]string{{`</c:f>`, `</c:f><c:extLst/>`}, {`<c:f>`, `<c:f extra="1">`}, {`</c:` + cache + `>`, `<c:extLst/></c:` + cache + `>`}, {`idx="900"`, `idx="-1"`}, {`<c:v>STALE</c:v>`, `<c:unknown/>`}, {`$B$2:$B$3`, `Name`}} {
				node, e = parseNativeXML([]byte(strings.ReplaceAll(source, pair[0], pair[1])), "bad.xml")
				if e != nil {
					continue
				}
				if _, ok = extractNativeChartReference(node, d, numeric); ok {
					t.Fatalf("unqualified reference/cache accepted %v", pair)
				}
			}
		}
	}
}
