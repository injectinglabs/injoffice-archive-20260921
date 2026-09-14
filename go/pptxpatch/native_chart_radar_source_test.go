package pptxpatch

import (
	"strings"
	"testing"
)

func nativeRadarXML(strict, filled bool) string {
	s := nativeConnectedXML(strict, false)
	s = strings.ReplaceAll(s, "lineChart", "radarChart")
	style := "standard"
	if filled {
		style = "filled"
	}
	s = strings.Replace(s, `<c:grouping val="standard"/>`, `<c:radarStyle val="`+style+`"/>`, 1)
	for _, v := range []string{`<c:smooth val="0"/>`, `<c:marker val="0"/>`, `<c:crossBetween val="between"/>`} {
		s = strings.ReplaceAll(s, v, "")
	}
	s = strings.Replace(s, `</c:plotArea></c:chart>`, `</c:plotArea><c:plotVisOnly val="0"/></c:chart>`, 1)
	if filled {
		s = strings.Replace(s, `<c:spPr><a:noFill/><a:ln`, `<c:spPr><a:solidFill><a:srgbClr val="E53935"/></a:solidFill><a:ln`, 1)
	}
	return s
}
func TestNativeRadarSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, filled := range []bool{false, true} {
			d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
			input := nativeRadarXML(strict, filled)
			got := extractNativeChartRadarSource([]byte(input), "chart.xml", d, false)
			if got == nil || len(got.Series) != 1 {
				t.Fatalf("missing strict=%v filled=%v", strict, filled)
			}
			s := got.Series[0]
			if (s.Fill != nil) != filled || s.Color != "#123456" || len(s.Values) != 3 || len(s.Categories) != 3 {
				t.Fatal("source data or paint lost")
			}
			if extractNativeChartRadarSource([]byte(input), "chart.xml", d, true) != nil {
				t.Fatal("mixed authority")
			}
		}
	}
}
func TestNativeRadarSourceRefusesExtras(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	input := nativeRadarXML(false, false)
	for name, pair := range map[string][2]string{
		"markerStyle": {`radarStyle val="standard"`, `radarStyle val="marker"`}, "markers": {`symbol val="none"`, `symbol val="circle"`},
		"missingStyle": {`<c:radarStyle val="standard"/>`, ``}, "smooth": {`</c:ser>`, `<c:smooth val="0"/></c:ser>`},
		"pointOverride": {`<c:cat>`, `<c:dPt/><c:cat>`}, "theme": {`srgbClr`, `schemeClr`}, "hidden": {`plotVisOnly val="0"`, `plotVisOnly val="1"`}, "sparse": {`ptCount val="3"`, `ptCount val="4"`},
	} {
		t.Run(name, func(t *testing.T) {
			if extractNativeChartRadarSource([]byte(strings.ReplaceAll(input, pair[0], pair[1])), "chart.xml", d, false) != nil {
				t.Fatal("unqualified source")
			}
		})
	}
}

func TestNativeRadarWorkbookAndSourceOrder(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		input := nativeRadarXML(strict, true)
		for _, item := range []struct{ name, kind, formula string }{{"cat", "strRef", "Data!A1:A3"}, {"val", "numRef", "Data!B1:B3"}} {
			start := strings.Index(input, "<c:"+item.name+">")
			end := strings.Index(input, "</c:"+item.name+">") + len("</c:"+item.name+">")
			cache := "strCache"
			if item.kind == "numRef" {
				cache = "numCache"
			}
			ref := `<c:` + item.name + `><c:` + item.kind + `><c:f>` + item.formula + `</c:f><c:` + cache + `><c:ptCount val="999"/><c:pt idx="900"><c:v>999</c:v></c:pt></c:` + cache + `></c:` + item.kind + `></c:` + item.name + `>`
			input = input[:start] + ref + input[end:]
		}
		input = strings.Replace(input, `</c:chartSpace>`, `<c:externalData xmlns:r="`+d.rels+`" r:id="workbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`, 1)
		start, end := strings.Index(input, "<c:ser>"), strings.Index(input, "</c:ser>")+len("</c:ser>")
		first := input[start:end]
		second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`).Replace(first)
		input = input[:start] + second + first + input[end:]
		got := extractNativeChartRadarSource([]byte(input), "chart.xml", d, true)
		if got == nil || len(got.Series) != 2 || got.Series[0].Order != 1 || got.Series[1].Order != 0 || got.Series[0].Index != 9 {
			t.Fatal("source sequence/order metadata lost")
		}
		for _, s := range got.Series {
			if len(s.Values) != 0 || len(s.Categories) != 0 || s.ValueReference.Formula != "Data!B1:B3" || !s.ValueReference.CachePresent {
				t.Fatal("reference/cache authority changed")
			}
		}
		for _, bad := range []string{strings.Replace(input, `order val="1"`, `order val="0"`, 1), strings.Replace(input, `Data!B1:B3`, `Data!B1:B2`, 1), strings.Replace(input, `Data!A1:A3`, `Data!A2:A4`, 1), strings.Replace(input, `Data!B1:B3`, `[other.xlsx]Data!B1:B3`, 1)} {
			if extractNativeChartRadarSource([]byte(bad), "chart.xml", d, true) != nil {
				t.Fatal("mismatched or external source admitted")
			}
		}
	}
}
