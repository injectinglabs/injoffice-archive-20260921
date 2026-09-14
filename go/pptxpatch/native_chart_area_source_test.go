package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func nativeAreaXML(strict bool, grouping string) string {
	source := nativeConnectedXML(strict, false)
	start := strings.Index(source, "<c:spPr>")
	end := strings.Index(source[start:], "</c:spPr>") + start + len("</c:spPr>")
	source = source[:start] + `<c:spPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>` + source[end:]
	return strings.NewReplacer("lineChart", "areaChart", `grouping val="standard"`, `grouping val="`+grouping+`"`, `<c:marker><c:symbol val="none"/></c:marker>`, "", `<c:marker val="0"/>`, "", `<c:smooth val="0"/>`, "").Replace(source)
}
func TestNativeChartAreaSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for _, grouping := range []string{"standard", "stacked", "percentStacked"} {
			source := nativeAreaXML(strict, grouping)
			if grouping != "standard" {
				source = strings.ReplaceAll(source, "<c:v>-1</c:v>", "<c:v>1</c:v>")
			}
			payload := []byte(source)
			before := bytes.Clone(payload)
			area := extractNativeChartArea(payload, "chart.xml", d)
			if area == nil || area.Grouping != grouping || area.Series[0].Color != "#123456" || len(area.Categories) != 3 {
				t.Fatal("missing area", area)
			}
			if !bytes.Equal(payload, before) {
				t.Fatal("source mutated")
			}
		}
	}
}
func TestNativeChartAreaSourceRefusals(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	original := nativeAreaXML(false, "standard")
	for name, pair := range map[string][2]string{
		"duplicate grouping": {`<c:grouping val="standard"/>`, `<c:grouping val="standard"/><c:grouping val="standard"/>`},
		"point formatting":   {"</c:ser>", "<c:dPt/></c:ser>"},
		"error bars":         {"</c:ser>", "<c:errBars/></c:ser>"},
		"theme":              {"srgbClr", "schemeClr"},
		"sparse":             {`ptCount val="3"`, `ptCount val="4"`},
		"duplicate point":    {`pt idx="2"`, `pt idx="1"`},
		"line combination":   {"</c:areaChart>", "</c:areaChart><c:lineChart/>"},
		"cache":              {"numLit", "numCache"},
		"unknown attribute":  {"<c:areaChart>", `<c:areaChart unexpected="1">`},
	} {
		t.Run(name, func(t *testing.T) {
			if extractNativeChartArea([]byte(strings.ReplaceAll(original, pair[0], pair[1])), "chart.xml", d) != nil {
				t.Fatal("unsupported source admitted")
			}
		})
	}
}

func TestNativeChartAreaSourceSeriesOrder(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	original := nativeAreaXML(false, "standard")
	start := strings.Index(original, "<c:ser>")
	end := strings.Index(original, "</c:ser>") + len("</c:ser>")
	first := original[start:end]
	second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, `123456`, `ABCDEF`).Replace(first)
	source := strings.Replace(original, first, second+first, 1)
	area := extractNativeChartArea([]byte(source), "chart.xml", d)
	if area == nil || len(area.Series) != 2 || area.Series[0].Index != 9 || area.Series[1].Index != 5 || area.Series[0].Color != "#ABCDEF" {
		t.Fatal("authored order lost", area)
	}
	for _, invalidSecond := range []string{
		strings.Replace(second, `<c:v>B</c:v>`, `<c:v>other</c:v>`, 1),
		strings.Replace(second, `order val="1"`, `order val="2"`, 1),
		strings.Replace(second, `idx val="9"`, `idx val="5"`, 1),
	} {
		if extractNativeChartArea([]byte(strings.Replace(original, first, first+invalidSecond, 1)), "chart.xml", d) != nil {
			t.Fatal("ambiguous series admitted")
		}
	}
}
