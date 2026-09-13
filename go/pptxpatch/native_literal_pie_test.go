package pptxpatch

import (
	"encoding/xml"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func xmlNamePresentation(strict bool) xml.Name {
	ns := nsPresentationTransitional
	if strict {
		ns = nsPresentationStrict
	}
	return xml.Name{Space: ns, Local: "presentation"}
}

func nativeLiteralPieXML(strict bool) string {
	c, a := nsChartTransitional, nsDrawingTransitional
	if strict {
		c, a = nsChartStrict, nsDrawingStrict
	}
	transparent := `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`
	point := func(idx, color string) string {
		return `<c:dPt><c:idx val="` + idx + `"/><c:spPr><a:solidFill><a:srgbClr val="` + color + `"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr></c:dPt>`
	}
	return `<c:chartSpace xmlns:c="` + c + `" xmlns:a="` + a + `"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:pieChart><c:varyColors val="0"/><c:ser><c:idx val="0"/><c:order val="0"/>` + point("0", "FF0000") + point("1", "00FF00") + `<c:val><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="2"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>3</c:v></c:pt></c:numLit></c:val></c:ser><c:firstSliceAng val="90"/></c:pieChart>` + transparent + `</c:plotArea></c:chart>` + transparent + `</c:chartSpace>`
}
func TestNativeLiteralPieSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		pie := extractNativeLiteralPie([]byte(nativeLiteralPieXML(strict)), "chart.xml", d)
		if pie == nil || pie.FirstSliceAngle != 90 || len(pie.Values) != 2 || pie.Values[1] != 3 || pie.Colors[0] != "#FF0000" {
			t.Fatalf("missing literals: %#v", pie)
		}
		bytes := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: nativeLiteralPieXML(strict)})
		deck, err := ExtractNativePPTX(bytes, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		chart := nativeFixtureChart(t, deck.Slides[0])
		if chart.Chart.LiteralPie == nil || chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatal("literal chart lost preserve-only ownership")
		}
		encoded, err := MarshalNativePPTXJSON(deck)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = DecodeNativePPTXJSON(encoded); err != nil {
			t.Fatal(err)
		}
	}
}
func TestNativeLiteralPieRefusesUnsupportedSource(t *testing.T) {
	original := nativeLiteralPieXML(false)
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	for name, change := range map[string][2]string{
		"cache": {"numLit", "numCache"}, "3d": {"pieChart", "pie3DChart"}, "count": {"ptCount val=\"2\"", "ptCount val=\"3\""}, "negative": {"<c:v>1</c:v>", "<c:v>-1</c:v>"}, "decimal": {"<c:v>1</c:v>", "<c:v>1.5</c:v>"}, "zero": {"<c:v>1</c:v>", "<c:v>0</c:v>"}, "nonfinite": {"<c:v>1</c:v>", "<c:v>NaN</c:v>"}, "duplicate point": {"<c:pt idx=\"1\">", "<c:pt idx=\"0\">"}, "theme color": {"srgbClr", "schemeClr"}, "rotation": {"firstSliceAng val=\"90\"", "firstSliceAng val=\"361\""}, "effect": {"<a:ln>", "<a:ln foo=\"1\">"}, "legend": {"</c:chart>", "<c:legend/></c:chart>"}, "label": {"</c:ser>", "<c:dLbls/></c:ser>"}, "extension": {"</c:chartSpace>", "<c:extLst/></c:chartSpace>"}, "explosion": {"</c:dPt>", "<c:explosion val=\"10\"/></c:dPt>"},
	} {
		t.Run(name, func(t *testing.T) {
			payload := strings.ReplaceAll(original, change[0], change[1])
			if extractNativeLiteralPie([]byte(payload), "chart.xml", d) != nil {
				t.Fatal("unsupported chart projected")
			}
		})
	}
}

// Optional local browser evidence export; generated source files stay outside git.
func TestNativeLiteralPieBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_LITERAL_PIE_EVIDENCE")
	if dir == "" {
		t.Skip("optional browser fixture export")
	}
	for name, source := range map[string]string{"literal": nativeLiteralPieXML(false), "refused": strings.ReplaceAll(nativeLiteralPieXML(false), "numLit", "numCache")} {
		data := nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: source})
		if err := os.WriteFile(filepath.Join(dir, name+".pptx"), data, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
