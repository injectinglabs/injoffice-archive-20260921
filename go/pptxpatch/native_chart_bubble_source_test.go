package pptxpatch

import (
	"strings"
	"testing"
)

func nativeBubbleXML(strict, workbook bool) string {
	source := nativeConnectedXML(strict, true)
	old := nativeConnectedSeriesXML(true)
	series := old
	start := strings.Index(series, "<c:spPr>")
	end := strings.Index(series, "</c:spPr>") + len("</c:spPr>")
	series = series[:start] + `<c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:invertIfNegative val="0"/>` + series[end:]
	series = strings.Replace(series, `<c:marker><c:symbol val="none"/></c:marker>`, "", 1)
	sizes := `<c:bubbleSize><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="2"><c:v>1e-100</c:v></c:pt><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:bubbleSize><c:bubble3D val="0"/>`
	series = strings.Replace(series, `<c:smooth val="0"/>`, sizes, 1)
	if workbook {
		for _, item := range []struct{ name, col string }{{"xVal", "A"}, {"yVal", "B"}, {"bubbleSize", "C"}} {
			start := strings.Index(series, "<c:"+item.name+">")
			end := strings.Index(series, "</c:"+item.name+">") + len("</c:"+item.name+">")
			ref := `<c:` + item.name + `><c:numRef><c:f>Data!` + item.col + `1:` + item.col + `3</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="999"/><c:pt idx="900"><c:v>999</c:v></c:pt></c:numCache></c:numRef></c:` + item.name + `>`
			series = series[:start] + ref + series[end:]
		}
	}
	source = strings.Replace(source, old, series, 1)
	source = strings.ReplaceAll(source, "scatterChart", "bubbleChart")
	source = strings.Replace(source, `<c:scatterStyle val="line"/>`, "", 1)
	scale := "100"
	if strict {
		scale += "%"
	}
	source = strings.Replace(source, `</c:ser><c:axId`, `</c:ser><c:bubble3D val="0"/><c:bubbleScale val="`+scale+`"/><c:showNegBubbles val="0"/><c:sizeRepresents val="area"/><c:axId`, 1)
	source = strings.Replace(source, `</c:plotArea></c:chart>`, `</c:plotArea><c:plotVisOnly val="0"/></c:chart>`, 1)
	if workbook {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		source = strings.Replace(source, `</c:chartSpace>`, `<c:externalData xmlns:r="`+d.rels+`" r:id="workbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`, 1)
	}
	return source
}
func TestNativeBubbleSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, workbook := range []bool{false, true} {
			d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
			input := nativeBubbleXML(strict, workbook)
			source := extractNativeChartBubbleSource([]byte(input), "chart.xml", d, workbook)
			if source == nil || source.BubbleScale != 100 || source.SizeRepresents != "area" || len(source.Series) != 1 {
				t.Fatalf("missing strict=%v workbook=%v", strict, workbook)
			}
			s := source.Series[0]
			if strings.Join(s.Colors, ",") != "#123456,#123456,#123456" {
				t.Fatal("paint lost")
			}
			if workbook {
				if len(s.Values) != 0 || s.SizeReference.Formula != "Data!C1:C3" || !s.SizeReference.CachePresent || source.ExternalData == nil {
					t.Fatal("cache authority")
				}
			} else if strings.Join(s.Sizes, ",") != "1,4,1e-100" || strings.Join(s.XValues, ",") != "2,-1,2" {
				t.Fatal("values/order lost")
			}
			if extractNativeChartBubbleSource([]byte(input), "chart.xml", d, !workbook) != nil {
				t.Fatal("origins mixed")
			}
		}
	}
}
func TestNativeBubbleRefusal(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	input := nativeBubbleXML(false, false)
	for name, pair := range map[string][2]string{"3D": {"bubble3D val=\"0\"", "bubble3D val=\"1\""}, "negativeflag": {"showNegBubbles val=\"0\"", "showNegBubbles val=\"1\""}, "negativevalue": {"<c:v>4</c:v>", "<c:v>-4</c:v>"}, "missingstyle": {"<c:sizeRepresents val=\"area\"/>", ""}, "largescale": {"bubbleScale val=\"100\"", "bubbleScale val=\"301\""}, "trendline": {"<c:xVal>", "<c:trendline/><c:xVal>"}, "theme": {"srgbClr", "schemeClr"}, "sparse": {"ptCount val=\"3\"", "ptCount val=\"4\""}, "duplicate": {"pt idx=\"2\"", "pt idx=\"1\""}, "hidden": {"plotVisOnly val=\"0\"", "plotVisOnly val=\"1\""}} {
		t.Run(name, func(t *testing.T) {
			if extractNativeChartBubbleSource([]byte(strings.ReplaceAll(input, pair[0], pair[1])), "chart.xml", d, false) != nil {
				t.Fatal("unqualified source")
			}
		})
	}
	for _, scale := range []string{"0", "300"} {
		changed := strings.Replace(input, `bubbleScale val="100"`, `bubbleScale val="`+scale+`"`, 1)
		if extractNativeChartBubbleSource([]byte(changed), "chart.xml", d, false) == nil {
			t.Fatal("scale boundary")
		}
	}
	if extractNativeChartBubbleSource([]byte(strings.Replace(input, `sizeRepresents val="area"`, `sizeRepresents val="w"`, 1)), "chart.xml", d, false) == nil {
		t.Fatal("width mode")
	}
}

func TestNativeBubbleSeriesPointPaintOrderAndBudgets(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	input := nativeBubbleXML(false, false)
	start := strings.Index(input, "<c:ser>")
	end := strings.Index(input, "</c:ser>") + len("</c:ser>")
	original := input[start:end]
	second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, `123456`, `ABCDEF`).Replace(original)
	combined := strings.Replace(input, original, second+original, 1)
	source := extractNativeChartBubbleSource([]byte(combined), "chart.xml", d, false)
	if source == nil || source.Series[0].Index != 5 || source.Series[1].Index != 9 || source.Series[1].Colors[0] != "#ABCDEF" {
		t.Fatal("series order or paint lost")
	}
	override := `<c:dPt><c:idx val="2"/><c:spPr><a:solidFill><a:srgbClr val="AABBCC"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt>`
	changed := strings.Replace(input, "<c:xVal>", override+"<c:xVal>", 1)
	source = extractNativeChartBubbleSource([]byte(changed), "chart.xml", d, false)
	if source == nil || strings.Join(source.Series[0].Colors, ",") != "#123456,#123456,#AABBCC" || strings.Join(source.Series[0].XValues, ",") != "2,-1,2" {
		t.Fatal("point override reorders source data")
	}
	for _, bad := range []string{strings.Replace(input, original, original+original, 1), strings.Replace(combined, `order val="1"`, `order val="2"`, 1), strings.Replace(changed, "<c:xVal>", override+"<c:xVal>", 1)} {
		if extractNativeChartBubbleSource([]byte(bad), "chart.xml", d, false) != nil {
			t.Fatal("duplicate/sparse source accepted")
		}
	}
}
