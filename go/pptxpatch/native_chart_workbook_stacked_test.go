package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func nativeWorkbookStackedXML(strict bool, family, grouping string) string {
	source := nativeWorkbookChartXML(strict, family)
	source = strings.NewReplacer(`grouping val="clustered"`, `grouping val="`+grouping+`"`, `grouping val="standard"`, `grouping val="`+grouping+`"`, `overlap val="0"`, `overlap val="100"`, `overlap val="0%"`, `overlap val="100%"`).Replace(source)
	return source
}
func TestNativeChartWorkbookStackedSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, family := range []string{"column", "bar", "line"} {
			for _, grouping := range []string{"stacked", "percentStacked"} {
				d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
				source := nativeWorkbookStackedXML(strict, family, grouping)
				first := nativeWorkbookSeriesXML(strict, family == "line", false)
				second := strings.NewReplacer(`idx val="7"`, `idx val="9"`, `idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, `$B$`, `$C$`).Replace(first)
				source = strings.Replace(source, first, second+first, 1)
				payload := []byte(source)
				before := append([]byte(nil), payload...)
				result := extractNativeChartWorkbookSource(payload, "chart.xml", d)
				if result == nil || result.Grouping != grouping || len(result.Series) != 2 || result.Series[0].Order != 1 || result.Series[1].Order != 0 || result.Series[0].Index != 9 || !bytes.Equal(before, payload) {
					t.Fatal("lost grouping/XML source order")
				}
				public := nativeWorkbookPublicSource(result)
				if public.Grouping != grouping || public.Series[0].ValueReference == nil || !strings.Contains(public.Series[0].ValueReference.Formula, "$C$") || !public.Series[0].ValueReference.CachePresent {
					t.Fatal("lost exact workbook reference/cache evidence")
				}
				if family == "line" {
					if public.Overlap != nil {
						t.Fatal("line gained bar metadata")
					}
				} else if public.Overlap == nil || *public.Overlap != 100 {
					t.Fatal("lost full overlap")
				}
				if extractNativeLiteralStackedBar(payload, "chart.xml", d) != nil || extractNativeLiteralStackedLine(payload, "chart.xml", d) != nil {
					t.Fatal("reference source became literal")
				}
				for _, pair := range [][2]string{{`grouping val="` + grouping + `"`, `grouping val="guess"`}, {`order val="1"`, `order val="0"`}, {`overlap val="100"`, `overlap val="99"`}, {`overlap val="100%"`, `overlap val="99%"`}, {`plotVisOnly val="0"`, `plotVisOnly val="1"`}} {
					changed := strings.ReplaceAll(source, pair[0], pair[1])
					if changed != source && extractNativeChartWorkbookSource([]byte(changed), "bad.xml", d) != nil {
						t.Fatal("unqualified stacked source admitted", pair)
					}
				}
			}
		}
	}
}
