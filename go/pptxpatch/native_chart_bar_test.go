package pptxpatch

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func nativeBarAxisXML(value bool, visible bool) string {
	name, id, other, pos, scale, cross := "catAx", "10", "20", "b", `<c:orientation val="minMax"/>`, `<c:crosses val="min"/>`
	if value {
		name, id, other, pos, scale, cross = "valAx", "20", "10", "l", `<c:orientation val="minMax"/><c:max val="20"/><c:min val="-10"/>`, `<c:crossesAt val="0"/><c:crossBetween val="between"/>`
	}
	deleted, paint := "1", ""
	if visible {
		deleted = "0"
		paint = `<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="none"/><c:spPr><a:noFill/><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="111111"/></a:solidFill><a:prstDash val="solid"/><a:headEnd type="none"/><a:tailEnd type="none"/></a:ln></c:spPr>`
	}
	return `<c:` + name + `><c:axId val="` + id + `"/><c:scaling>` + scale + `</c:scaling><c:delete val="` + deleted + `"/><c:axPos val="` + pos + `"/>` + paint + `<c:crossAx val="` + other + `"/>` + cross + `</c:` + name + `>`
}
func nativeBarXML(strict, visible bool) string {
	c, a := nsChartTransitional, nsDrawingTransitional
	if strict {
		c, a = nsChartStrict, nsDrawingStrict
	}
	transparent := `<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>`
	result := `<c:chartSpace xmlns:c="` + c + `" xmlns:a="` + a + `"><c:chart><c:autoTitleDeleted val="1"/><c:plotArea><c:layout/><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>` + nativeBarSeriesXML(strict) + `<c:gapWidth val="150"/><c:overlap val="0"/><c:axId val="20"/><c:axId val="10"/></c:barChart>` + nativeBarAxisXML(true, visible) + nativeBarAxisXML(false, visible) + transparent + `</c:plotArea></c:chart>` + transparent + `</c:chartSpace>`
	if strict {
		result = strings.ReplaceAll(result, `gapWidth val="150"`, `gapWidth val="150%"`)
		result = strings.ReplaceAll(result, `overlap val="0"`, `overlap val="0%"`)
	}
	return result
}
func TestNativeChartBarSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, visible := range []bool{false, true} {
			d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
			bar := extractNativeChartBar([]byte(nativeBarXML(strict, visible)), "chart.xml", d)
			if bar == nil || bar.GapWidth != 150 || bar.Series[0].Values[0] != "-0.5" || bar.ValueAxis.Min != "-10" || bar.CategoryAxis.Deleted == visible {
				t.Fatalf("missing chart: %#v", bar)
			}
		}
	}
}
func TestNativeChartBarRefusesUnqualifiedSource(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	original := nativeBarXML(false, true)
	for name, pair := range map[string][2]string{
		"stacked": {"clustered", "stacked"}, "automatic gap": {"<c:gapWidth val=\"150\"/>", ""}, "overlap": {"overlap val=\"0\"", "overlap val=\"1\""}, "ids": {"<c:axId val=\"20\"/>", "<c:axId val=\"10\"/>"}, "missing reciprocal": {"crossAx val=\"20\"", "crossAx val=\"30\""}, "implicit label": {"<c:tickLblPos val=\"none\"/>", ""}, "labels": {"tickLblPos val=\"none\"", "tickLblPos val=\"nextTo\""}, "zero span": {"max val=\"20\"", "max val=\"-10\""}, "outside zero": {"min val=\"-10\"", "min val=\"1\""}, "log scale": {"<c:scaling>", "<c:scaling><c:logBase val=\"10\"/>"}, "cross midpoint": {"crossBetween val=\"between\"", "crossBetween val=\"midCat\""}, "cross nonzero": {"crossesAt val=\"0\"", "crossesAt val=\"2\""}, "implicit width": {" w=\"12700\"", ""}, "wrong direction": {"barDir val=\"col\"", "barDir val=\"bar\""}, "legend": {"</c:chart>", "<c:legend/></c:chart>"}, "manual layout": {"<c:layout/>", "<c:layout><c:manualLayout/></c:layout>"},
	} {
		t.Run(name, func(t *testing.T) {
			if extractNativeChartBar([]byte(strings.ReplaceAll(original, pair[0], pair[1])), "chart.xml", d) != nil {
				t.Fatal("unsupported chart projected")
			}
		})
	}
}

func TestNativeChartBarExtractionContract(t *testing.T) {
	for _, strict := range []bool{false, true} {
		payload := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: nativeBarXML(strict, true)})
		deck, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		chart := nativeFixtureChart(t, deck.Slides[0])
		if chart.Chart.LiteralBar == nil || !validNativeLiteralBar(chart.Chart.LiteralBar) || chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatal("missing read-only source chart")
		}
		encoded, err := MarshalNativePPTXJSON(deck)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := DecodeNativePPTXJSON(encoded); err != nil {
			t.Fatal(err)
		}
		original := chart.Chart.LiteralBar.ValueAxis.CrossesAt
		bad := "NaN"
		chart.Chart.LiteralBar.ValueAxis.CrossesAt = &bad
		if validNativeLiteralBar(chart.Chart.LiteralBar) {
			t.Fatal("accepted invalid axis")
		}
		chart.Chart.LiteralBar.ValueAxis.CrossesAt = original
	}
}

func TestNativeChartBarEvidence(t *testing.T) {
	dir := os.Getenv("INJOFFICE_BAR_EVIDENCE_DIR")
	if dir == "" {
		t.Skip("external evidence not requested")
	}
	series := nativeBarSeriesXML(false)
	second := strings.NewReplacer(`<c:idx val="7"/>`, `<c:idx val="9"/>`, `<c:order val="0"/>`, `<c:order val="1"/>`, `Signed series`, `Second series`, `-0.5`, `-7.5`, `1.25e+1`, `5.5`, `123456`, `FF0000`, `ABCDEF`, `0000FF`).Replace(series)
	column := strings.Replace(nativeBarXML(false, true), series, series+second, 1)
	bar := strings.NewReplacer(`barDir val="col"`, `barDir val="bar"`, `axPos val="b"`, `axPos val="l"`, `axPos val="l"`, `axPos val="b"`).Replace(column)
	for name, source := range map[string]string{"column": column, "bar": bar, "refused": strings.ReplaceAll(column, "numLit", "numCache")} {
		data := nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: source})
		if err := os.WriteFile(filepath.Join(dir, name+".pptx"), data, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
