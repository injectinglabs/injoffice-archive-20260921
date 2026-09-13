package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeChartScatterSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		chart := extractNativeChartScatter([]byte(nativeConnectedXML(strict, true)), "chart.xml", d)
		if chart == nil || chart.Family != "scatterChart" || len(chart.Categories) != 0 || strings.Join(chart.Series[0].XValues, ",") != "2,-1,2" || strings.Join(chart.Series[0].Values, ",") != "-1,2.5,1e0" || chart.XAxis.ID != 10 || chart.YAxis.ID != 20 {
			t.Fatalf("XY source reordered or missing: %#v", chart)
		}
	}
}
func TestNativeChartScatterRejectsOtherStylesAndAxes(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	original := nativeConnectedXML(false, true)
	for _, style := range []string{"lineMarker", "marker", "smooth", "smoothMarker", "none"} {
		if extractNativeChartScatter([]byte(strings.ReplaceAll(original, `scatterStyle val="line"`, `scatterStyle val="`+style+`"`)), "chart.xml", d) != nil {
			t.Fatal("non-straight source style accepted", style)
		}
	}
	if extractNativeChartScatter([]byte(strings.ReplaceAll(original, `</c:valAx>`, `<c:crossBetween val="between"/></c:valAx>`)), "chart.xml", d) != nil {
		t.Fatal("category-only axis clause accepted")
	}
}
