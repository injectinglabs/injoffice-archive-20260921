package pptxpatch

import (
	"strings"
	"testing"
)

func nativeConnectedSeriesXML(scatter bool) string {
	literal := func(numeric bool, values []string) string {
		name, format := "strLit", ""
		if numeric {
			name, format = "numLit", "<c:formatCode>General</c:formatCode>"
		}
		return `<c:` + name + `>` + format + `<c:ptCount val="3"/><c:pt idx="2"><c:v>` + values[2] + `</c:v></c:pt><c:pt idx="0"><c:v>` + values[0] + `</c:v></c:pt><c:pt idx="1"><c:v>` + values[1] + `</c:v></c:pt></c:` + name + `>`
	}
	xKind, yKind, values := "cat", "val", []string{"A", "B", "C"}
	if scatter {
		xKind, yKind, values = "xVal", "yVal", []string{"2", "-1", "2"}
	}
	return `<c:ser><c:idx val="5"/><c:order val="0"/><c:tx><c:v>Connected series</c:v></c:tx><c:spPr><a:noFill/><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:prstDash val="solid"/><a:round/><a:headEnd type="none"/><a:tailEnd type="none"/></a:ln></c:spPr><c:marker><c:symbol val="none"/></c:marker><c:` + xKind + `>` + literal(scatter, values) + `</c:` + xKind + `><c:` + yKind + `>` + literal(true, []string{"-1", "2.5", "1e0"}) + `</c:` + yKind + `><c:smooth val="0"/></c:ser>`
}
func nativeConnectedXML(strict, scatter bool) string {
	c, a := nsChartTransitional, nsDrawingTransitional
	if strict {
		c, a = nsChartStrict, nsDrawingStrict
	}
	family, style, flags := "lineChart", `<c:grouping val="standard"/>`, `<c:marker val="0"/><c:smooth val="0"/>`
	xAxis := nativeBarAxisXML(false, true)
	yAxis := nativeBarAxisXML(true, true)
	if scatter {
		family, style, flags = "scatterChart", `<c:scatterStyle val="line"/>`, ""
		xAxis = strings.NewReplacer(`axId val="20"`, `axId val="10"`, `crossAx val="10"`, `crossAx val="20"`, `axPos val="l"`, `axPos val="b"`, `<c:crossBetween val="between"/>`, "").Replace(yAxis)
		yAxis = strings.ReplaceAll(yAxis, `<c:crossBetween val="between"/>`, "")
	}
	transparent := `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`
	return `<c:chartSpace xmlns:c="` + c + `" xmlns:a="` + a + `"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:` + family + `>` + style + `<c:varyColors val="0"/>` + nativeConnectedSeriesXML(scatter) + flags + `<c:axId val="10"/><c:axId val="20"/></c:` + family + `>` + yAxis + xAxis + transparent + `</c:plotArea></c:chart>` + transparent + `</c:chartSpace>`
}
func TestNativeChartLineSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		line := extractNativeChartLine([]byte(nativeConnectedXML(strict, false)), "chart.xml", d)
		if line == nil || line.Family != "lineChart" || strings.Join(line.Categories, ",") != "A,B,C" || strings.Join(line.Series[0].Values, ",") != "-1,2.5,1e0" || line.Series[0].Color != "#123456" || line.Series[0].Width != 12700 {
			t.Fatalf("missing source line: %#v", line)
		}
	}
}
func TestNativeChartConnectedSourceRefusesUnsupported(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	for _, scatter := range []bool{false, true} {
		original := nativeConnectedXML(false, scatter)
		for name, pair := range map[string][2]string{"smooth": {"smooth val=\"0\"", "smooth val=\"1\""}, "implicit smooth": {"<c:smooth val=\"0\"/>", ""}, "marker": {"symbol val=\"none\"", "symbol val=\"circle\""}, "unknown join": {"<a:round/>", "<a:bevel/>"}, "theme": {"srgbClr", "schemeClr"}, "cache": {"numLit", "numCache"}, "sparse": {"ptCount val=\"3\"", "ptCount val=\"4\""}, "duplicate point": {"pt idx=\"2\"", "pt idx=\"1\""}, "implicit paint": {" w=\"12700\"", ""}, "trendline": {"</c:ser>", "<c:trendline/></c:ser>"}, "error bar": {"</c:ser>", "<c:errBars/></c:ser>"}, "nonfinite": {"2.5", "INF"}, "unresolved axes": {"crossAx val=\"20\"", "crossAx val=\"99\""}} {
			t.Run(name+map[bool]string{false: "Line", true: "Scatter"}[scatter], func(t *testing.T) {
				source := strings.ReplaceAll(original, pair[0], pair[1])
				if extractNativeChartConnected([]byte(source), "chart.xml", d, scatter) != nil {
					t.Fatal("unsupported source projected")
				}
			})
		}
	}
}

func TestNativeChartConnectedSeriesOrderAndCategoryEquality(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	for _, scatter := range []bool{false, true} {
		original := nativeConnectedXML(false, scatter)
		first := nativeConnectedSeriesXML(scatter)
		second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, `Connected series`, `Second series`, `123456`, `ABCDEF`).Replace(first)
		source := strings.Replace(original, first, second+first, 1)
		chart := extractNativeChartConnected([]byte(source), "chart.xml", d, scatter)
		if chart == nil || len(chart.Series) != 2 || chart.Series[0].Index != 9 || chart.Series[1].Index != 5 || chart.Series[0].Color != "#ABCDEF" {
			t.Fatalf("source series order lost: %#v", chart)
		}
		if !scatter {
			mismatch := strings.Replace(original, first, first+strings.Replace(second, `<c:v>B</c:v>`, `<c:v>Different</c:v>`, 1), 1)
			if extractNativeChartLine([]byte(mismatch), "chart.xml", d) != nil {
				t.Fatal("mismatched category lists accepted")
			}
		}
	}
}
