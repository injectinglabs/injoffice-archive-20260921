package pptxpatch

import (
	"strings"
	"testing"
)

func nativeBarSeriesXML(strict bool) string {
	c, a := nsChartTransitional, nsDrawingTransitional
	if strict {
		c, a = nsChartStrict, nsDrawingStrict
	}
	return `<c:ser xmlns:c="` + c + `" xmlns:a="` + a + `"><c:idx val="7"/><c:order val="0"/><c:tx><c:v>Signed series</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/><c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt><c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="1"><c:v>Same</c:v></c:pt><c:pt idx="0"><c:v>Same</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="1"><c:v>1.25e+1</c:v></c:pt><c:pt idx="0"><c:v>-0.5</c:v></c:pt></c:numLit></c:val></c:ser>`
}
func TestNativeChartLiteralSeries(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		node, err := parseNativeXML([]byte(nativeBarSeriesXML(strict)), "series.xml")
		if err != nil {
			t.Fatal(err)
		}
		series, ok := extractNativeChartLiteralSeries(node, d)
		if !ok || series.Index != 7 || series.Order != 0 || series.Title == nil || *series.Title != "Signed series" || series.Values[0] != "-0.5" || series.Values[1] != "1.25e+1" || series.Colors[0] != "#123456" || series.Colors[1] != "#ABCDEF" || series.Categories[0] != "Same" || series.Categories[1] != "Same" {
			t.Fatalf("incorrect indexed source: %#v", series)
		}
	}
}
func TestNativeChartLiteralSeriesRefusal(t *testing.T) {
	original := nativeBarSeriesXML(false)
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	for name, change := range map[string][2]string{
		"encoded category": {"Same", "_x0041_"}, "encoded title": {"Signed series", "_x0041_"},
		"missing index": {"<c:idx val=\"7\"/>", ""}, "unknown": {"</c:ser>", "<c:shape val=\"box\"/></c:ser>"}, "cache": {"numLit", "numCache"}, "reference": {"strLit", "strRef"}, "inversion": {"invertIfNegative val=\"0\"", "invertIfNegative val=\"1\""}, "implicit inversion": {"<c:invertIfNegative val=\"0\"/>", ""}, "point mismatch": {"<c:ptCount val=\"2\"/>", "<c:ptCount val=\"3\"/>"}, "duplicate index": {"<c:pt idx=\"1\">", "<c:pt idx=\"0\">"}, "noncanonical index": {"<c:pt idx=\"1\">", "<c:pt idx=\"01\">"}, "nonfinite": {"-0.5", "NaN"}, "foreign point": {"<c:pt idx=\"1\">", "<c:pt xmlns:c=\"urn:foreign\" idx=\"1\">"}, "theme paint": {"srgbClr", "schemeClr"}, "late paint": {"<c:invertIfNegative val=\"0\"/>", "<c:invertIfNegative val=\"0\"/><c:spPr/>"}, "title markup": {"Signed series", "<c:v/>"}, "missing base": {"<a:srgbClr val=\"123456\"/>", ""}, "partial override": {"<a:srgbClr val=\"ABCDEF\"/>", ""},
	} {
		t.Run(name, func(t *testing.T) {
			node, err := parseNativeXML([]byte(strings.ReplaceAll(original, change[0], change[1])), "series.xml")
			if err != nil {
				return
			}
			if _, ok := extractNativeChartLiteralSeries(node, d); ok {
				t.Fatal("unsupported source projected")
			}
		})
	}
}
