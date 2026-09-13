package pptxpatch

import (
	"regexp"
	"strings"
	"testing"
)

func nativeWorkbookSeriesXML(strict, connected, scatter bool) string {
	source := nativeBarSeriesXML(strict)
	last := "3"
	if connected {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		source = strings.Replace(nativeConnectedSeriesXML(scatter), "<c:ser>", `<c:ser xmlns:c="`+d.chart+`" xmlns:a="`+d.drawing+`">`, 1)
		last = "4"
	}
	strRef := `<c:strRef><c:f>Sheet1!$A$2:$A$` + last + `</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>STALE CATEGORY</c:v></c:pt></c:strCache></c:strRef>`
	numRef := `<c:numRef><c:f>Sheet1!$B$2:$B$` + last + `</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>999</c:v></c:pt></c:numCache></c:numRef>`
	source = regexp.MustCompile(`<c:strLit>.*?</c:strLit>`).ReplaceAllStringFunc(source, func(string) string { return strRef })
	source = regexp.MustCompile(`<c:numLit>.*?</c:numLit>`).ReplaceAllStringFunc(source, func(string) string { return numRef })
	return source
}
func TestNativeChartWorkbookSourceSeries(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for _, mode := range []struct{ connected, scatter bool }{{false, false}, {true, false}, {true, true}} {
			source := nativeWorkbookSeriesXML(strict, mode.connected, mode.scatter)
			node, e := parseNativeXML([]byte(source), "series.xml")
			if e != nil {
				t.Fatal(e)
			}
			var series *nativeChartWorkbookSeries
			var ok bool
			if mode.connected {
				series, ok = extractNativeChartWorkbookConnectedSeries(node, d, mode.scatter)
			} else {
				series, ok = extractNativeChartWorkbookBarSeries(node, d)
			}
			if !ok || series.ValueReference == nil || !series.ValueReference.CachePresent {
				t.Fatalf("source descriptors missing %v %v %s", strict, mode, source)
			}
			if mode.scatter && series.XReference == nil || !mode.scatter && series.CategoryReference == nil {
				t.Fatal("source axis reference kind lost")
			}
			if !mode.connected && (len(series.Colors) != 2 || series.Colors[0] != "#123456" || series.Colors[1] != "#ABCDEF") {
				t.Fatal("source indexed paint lost")
			}
			for _, pair := range [][2]string{{`<c:f>Sheet1!$B$2:$B$`, `<c:f>External!Name`}, {`srgbClr`, `schemeClr`}, {`</c:ser>`, `<c:extLst/></c:ser>`}, {`symbol val="none"`, `symbol val="circle"`}, {`invertIfNegative val="0"`, `invertIfNegative val="1"`}} {
				changed := strings.ReplaceAll(source, pair[0], pair[1])
				if changed == source {
					continue
				}
				node, e = parseNativeXML([]byte(changed), "bad.xml")
				if e != nil {
					continue
				}
				if mode.connected {
					_, ok = extractNativeChartWorkbookConnectedSeries(node, d, mode.scatter)
				} else {
					_, ok = extractNativeChartWorkbookBarSeries(node, d)
				}
				if ok {
					t.Fatalf("unqualified source series accepted %v", pair)
				}
			}
		}
	}
}
