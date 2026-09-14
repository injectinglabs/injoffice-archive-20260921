package pptxpatch

import (
	"strings"
	"testing"
)

func nativeWorkbookChartXML(strict bool, family string) string {
	source := nativeBarXML(strict, true)
	series := nativeBarSeriesXML(strict)
	connected := family == "line" || family == "scatter"
	if connected {
		source = nativeConnectedXML(strict, family == "scatter")
		series = nativeConnectedSeriesXML(family == "scatter")
	}
	source = strings.Replace(source, series, nativeWorkbookSeriesXML(strict, connected, family == "scatter"), 1)
	if family == "bar" {
		source = strings.NewReplacer(`barDir val="col"`, `barDir val="bar"`, `axPos val="b"`, `axPos val="l"`, `axPos val="l"`, `axPos val="b"`).Replace(source)
	}
	source = strings.Replace(source, `</c:chart>`, `<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart>`, 1)
	d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
	return strings.Replace(source, `</c:chartSpace>`, `<c:externalData xmlns:r="`+d.rels+`" r:id="workbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`, 1)
}
func TestNativeChartWorkbookSourceFamilies(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for _, family := range []string{"column", "bar", "line", "scatter"} {
			source := nativeWorkbookChartXML(strict, family)
			chart := extractNativeChartWorkbookSource([]byte(source), "chart.xml", d)
			if chart == nil || len(chart.Series) != 1 || chart.ExternalData == nil || chart.DispBlanksAs == nil || *chart.DispBlanksAs != "gap" || chart.Series[0].ValueReference == nil {
				t.Fatalf("missing %s source descriptor: %s", family, source)
			}
			if chart.XAxis.Position != "b" || chart.YAxis.Position != "l" {
				t.Fatal("physical axis identity changed")
			}
			for _, pair := range [][2]string{
				{`<c:plotVisOnly val="0"/>`, ``}, {`plotVisOnly val="0"`, `plotVisOnly val="1"`},
				{`<c:dispBlanksAs val="gap"/>`, `<c:dispBlanksAs val="guess"/>`},
				{`</c:plotArea>`, `</c:plotArea><c:legend/>`},
				{`</c:chartSpace>`, `<c:extLst/></c:chartSpace>`},
				{`grouping val="clustered"`, `grouping val="stacked"`},
				{`grouping val="standard"`, `grouping val="stacked"`},
				{`scatterStyle val="line"`, `scatterStyle val="smoothMarker"`},
				{`<c:autoTitleDeleted val="1"/>`, `<c:autoTitleDeleted val="0"/>`},
				{`<c:axId val="10"/>`, `<c:axId val="999"/>`},
			} {
				changed := strings.ReplaceAll(source, pair[0], pair[1])
				if changed == source {
					continue
				}
				if extractNativeChartWorkbookSource([]byte(changed), "bad.xml", d) != nil {
					t.Fatalf("unqualified source %s accepted %v", family, pair)
				}
			}
		}
	}
}
