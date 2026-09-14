package pptxpatch

import (
	"encoding/json"
	"fmt"
	"strings"
	"testing"
)

// Repeat the original series in a deliberately nonnumeric XML sequence. Point
// XML remains untouched: indexed points and dPt overrides retain their authority.
func nativeSeriesOrderXML(source string) string {
	start := strings.Index(source, "<c:ser")
	end := strings.Index(source, "</c:ser>") + len("</c:ser>")
	original := source[start:end]
	var rows strings.Builder
	for _, order := range []int{2, 0, 1} {
		row := original
		idxStart := strings.Index(row, `<c:idx val="`) + len(`<c:idx val="`)
		idxEnd := strings.Index(row[idxStart:], `"`) + idxStart
		row = row[:idxStart] + fmt.Sprint(10+order) + row[idxEnd:]
		row = strings.Replace(row, `order val="0"`, fmt.Sprintf(`order val="%d"`, order), 1)
		rows.WriteString(row)
	}
	return source[:start] + rows.String() + source[end:]
}

func TestNativeChartSeriesOrderFamilies(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for _, family := range []string{"bar", "line", "scatter", "area", "bubble", "workbook-bar", "workbook-line", "workbook-scatter", "workbook-bubble"} {
			t.Run(fmt.Sprintf("%v/%s", strict, family), func(t *testing.T) {
				source := ""
				var extract func([]byte) any
				switch family {
				case "bar":
					source = nativeBarXML(strict, false)
					extract = func(b []byte) any { return extractNativeChartBar(b, "chart.xml", d) }
				case "line", "scatter":
					source = nativeConnectedXML(strict, family == "scatter")
					extract = func(b []byte) any { return extractNativeChartConnected(b, "chart.xml", d, family == "scatter") }
				case "area":
					source = nativeAreaXML(strict, "standard")
					extract = func(b []byte) any { return extractNativeChartArea(b, "chart.xml", d) }
				case "bubble", "workbook-bubble":
					source = nativeBubbleXML(strict, family == "workbook-bubble")
					extract = func(b []byte) any {
						return extractNativeChartBubbleSource(b, "chart.xml", d, family == "workbook-bubble")
					}
				default:
					source = nativeWorkbookChartXML(strict, strings.TrimPrefix(family, "workbook-"))
					extract = func(b []byte) any { return extractNativeChartWorkbookSource(b, "chart.xml", d) }
				}
				source = nativeSeriesOrderXML(source)
				payload := []byte(source)
				encoded, err := json.Marshal(extract(payload))
				if err != nil || string(encoded) == "null" {
					t.Fatalf("source refused: %s %v", encoded, err)
				}
				var decoded struct {
					Series []struct{ Index, Order int64 }
				}
				if err = json.Unmarshal(encoded, &decoded); err != nil {
					t.Fatal(err)
				}
				if len(decoded.Series) != 3 {
					t.Fatal("missing series")
				}
				for i, order := range []int64{2, 0, 1} {
					if decoded.Series[i].Index != 10+order || decoded.Series[i].Order != order {
						t.Fatalf("XML sequence/metadata changed: %s", encoded)
					}
				}
				if string(payload) != source {
					t.Fatal("source mutated")
				}
				for _, bad := range []string{"0", "3", "-0", "1.5"} {
					changed := strings.Replace(source, `order val="2"`, `order val="`+bad+`"`, 1)
					result, _ := json.Marshal(extract([]byte(changed)))
					if string(result) != "null" {
						t.Fatalf("invalid order %s admitted", bad)
					}
				}
			})
		}
	}
}
