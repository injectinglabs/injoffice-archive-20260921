package pptxpatch

import (
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func literalPiePaint(color string) string {
	return `<c:spPr><a:solidFill><a:srgbClr val="` + color + `"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>`
}
func literalPiePoint(index, color string) string {
	return `<c:dPt><c:idx val="` + index + `"/>` + literalPiePaint(color) + `</c:dPt>`
}
func literalPiePaintSource(strict bool, paint string, count int) string {
	source := nativeLiteralPieXML(strict)
	start, end := strings.Index(source, "<c:ser>"), strings.Index(source, "</c:ser>")+len("</c:ser>")
	series := `<c:ser><c:idx val="0"/><c:order val="0"/>` + paint + `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="` + fmt.Sprint(count) + `"/>`
	for i := 0; i < count; i++ {
		series += fmt.Sprintf(`<c:pt idx="%d"><c:v>%d</c:v></c:pt>`, i, i+1)
	}
	return source[:start] + series + `</c:numLit></c:val></c:ser>` + source[end:]
}

func TestNativeLiteralPieEffectivePaint(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for name, tc := range map[string]struct {
			paint  string
			colors []string
		}{
			"series":          {literalPiePaint("123abc"), []string{"#123ABC", "#123ABC", "#123ABC"}},
			"sparse":          {literalPiePaint("0000FF") + literalPiePoint("1", "FF0000"), []string{"#0000FF", "#FF0000", "#0000FF"}},
			"unordered":       {literalPiePaint("0000FF") + literalPiePoint("2", "00FF00") + literalPiePoint("0", "FF0000"), []string{"#FF0000", "#0000FF", "#00FF00"}},
			"complete points": {literalPiePoint("2", "00FF00") + literalPiePoint("0", "FF0000") + literalPiePoint("1", "0000FF"), []string{"#FF0000", "#0000FF", "#00FF00"}},
		} {
			t.Run(fmt.Sprintf("%s/strict=%v", name, strict), func(t *testing.T) {
				source := literalPiePaintSource(strict, tc.paint, 3)
				data := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: source})
				before := append([]byte(nil), data...)
				deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
				if err != nil {
					t.Fatal(err)
				}
				chart := nativeFixtureChart(t, deck.Slides[0])
				pie := chart.Chart.LiteralPie
				if pie == nil || !reflect.DeepEqual(pie.Colors, tc.colors) || !reflect.DeepEqual(pie.Values, []int64{1, 2, 3}) {
					t.Fatalf("wrong effective paint/value order: %#v", pie)
				}
				if chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || !reflect.DeepEqual(data, before) {
					t.Fatal("source ownership changed")
				}
				encoded, err := MarshalNativePPTXJSON(deck)
				if err != nil {
					t.Fatal(err)
				}
				if _, err = DecodeNativePPTXJSON(encoded); err != nil {
					t.Fatal(err)
				}
			})
		}
	}
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	paint := literalPiePaint("0000FF")
	for i := 63; i >= 0; i-- {
		paint += literalPiePoint(fmt.Sprint(i), "FF0000")
	}
	pie := extractNativeLiteralPie([]byte(literalPiePaintSource(false, paint, 64)), "chart.xml", d)
	if pie == nil || len(pie.Colors) != 64 || pie.Values[63] != 64 {
		t.Fatal("64-point budget boundary refused")
	}
}

func TestNativeLiteralPiePaintRefusal(t *testing.T) {
	base := literalPiePaint("0000FF")
	point := literalPiePoint("1", "FF0000")
	cases := map[string]string{
		"no paint": "", "sparse without series": point,
		"duplicate index":    base + point + point,
		"out of range":       base + literalPiePoint("3", "FF0000"),
		"negative":           base + literalPiePoint("-1", "FF0000"),
		"noncanonical":       base + literalPiePoint("01", "FF0000"),
		"huge":               base + literalPiePoint("9223372036854775808", "FF0000"),
		"series after point": point + base, "duplicate series": base + base,
		"empty series":            `<c:spPr/>` + point,
		"partial series":          strings.Replace(base, `<a:ln><a:noFill/></a:ln>`, "", 1) + point,
		"partial point":           base + strings.Replace(point, `<a:ln><a:noFill/></a:ln>`, "", 1),
		"empty point":             base + `<c:dPt><c:idx val="1"/><c:spPr/></c:dPt>`,
		"theme":                   strings.ReplaceAll(base, "srgbClr", "schemeClr"),
		"effect":                  strings.Replace(base, "</c:spPr>", "<a:effectLst/></c:spPr>", 1),
		"line":                    strings.Replace(base, "<a:noFill/>", "<a:solidFill><a:srgbClr val=\"000000\"/></a:solidFill>", 1),
		"qualified idx attr":      base + strings.Replace(point, `val="1"`, `c:val="1"`, 1),
		"foreign point namespace": base + strings.ReplaceAll(point, "c:dPt", "a:dPt"),
		"foreign paint namespace": strings.ReplaceAll(base, "c:spPr", "a:spPr"),
		"direct text":             strings.Replace(base, "<c:spPr>", "<c:spPr>unknown", 1),
		"rgb transform":           strings.Replace(base, `val="0000FF"/>`, `val="0000FF"><a:alpha val="50000"/></a:srgbClr>`, 1),
	}
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		for name, paint := range cases {
			t.Run(fmt.Sprintf("%s/strict=%v", name, strict), func(t *testing.T) {
				if extractNativeLiteralPie([]byte(literalPiePaintSource(strict, paint, 3)), "chart.xml", d) != nil {
					t.Fatal("ambiguous paint admitted")
				}
			})
		}
	}
}

func TestNativeLiteralPiePaintBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_LITERAL_PIE_PAINT_EVIDENCE")
	if dir == "" {
		t.Skip("optional external browser fixture")
	}
	for name, paint := range map[string]string{"sparse": literalPiePaint("0000FF") + literalPiePoint("1", "FF0000"), "refused": literalPiePoint("1", "FF0000")} {
		data := nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: literalPiePaintSource(false, paint, 3)})
		if err := os.WriteFile(filepath.Join(dir, name+".pptx"), data, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
