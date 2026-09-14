package pptxpatch

import (
	"strings"
	"testing"
)

func nativeLabeledBarXML(strict bool) string {
	source := nativeBarXML(strict, true)
	source = strings.ReplaceAll(source, `<c:tickLblPos val="none"/>`, `<c:tickLblPos val="low"/>`)
	source = strings.ReplaceAll(source, `</c:spPr><c:crossAx`, `</c:spPr>`+nativeAxisTextXML(strict)+`<c:crossAx`)
	source = strings.Replace(source, `<c:axPos val="l"/>`, `<c:axPos val="l"/><c:numFmt formatCode="0.0" sourceLinked="0"/>`, 1)
	source = strings.Replace(source, `</c:valAx>`, `<c:majorUnit val="5"/></c:valAx>`, 1)
	offset := "0"
	if strict {
		offset = "0%"
	}
	source = strings.Replace(source, `</c:catAx>`, `<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="`+offset+`"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="1"/></c:catAx>`, 1)
	return source
}
func TestNativeChartAxisLabelsSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		source := nativeLabeledBarXML(strict)
		chart := extractNativeLiteralBar([]byte(source), "chart.xml", d)
		if chart == nil || !validNativeLiteralBar(chart) || chart.CategoryAxis.Labels == nil || chart.ValueAxis.Labels == nil || *chart.ValueAxis.Labels.MajorUnit != "5" {
			t.Fatalf("source labels rejected strict=%v: %+v", strict, chart)
		}
		for _, pair := range [][2]string{
			{`sourceLinked="0"`, `sourceLinked="1"`}, {`formatCode="0.0"`, `formatCode="General"`},
			{`<c:majorUnit val="5"/>`, ``}, {`<c:majorUnit val="5"/>`, `<c:majorUnit val="3"/>`},
			{`<c:majorUnit val="5"/>`, `<c:majorUnit val="0.001"/>`},
			{`<c:tickLblSkip val="1"/>`, `<c:tickLblSkip val="2"/>`}, {`<c:noMultiLvlLbl val="1"/>`, ``},
			{`<c:majorTickMark val="none"/>`, `<c:majorTickMark val="out"/>`},
			{`<c:tickLblPos val="low"/>`, `<c:tickLblPos val="nextTo"/>`},
			{`<c:minorTickMark val="none"/>`, `<c:minorTickMark val="in"/>`},
			{`kern="0"`, `kern="1200"`}, {`<c:delete val="0"/>`, `<c:delete val="1"/>`},
		} {
			if got := extractNativeLiteralBar([]byte(strings.ReplaceAll(source, pair[0], pair[1])), "bad.xml", d); got != nil {
				t.Fatalf("unqualified labels accepted %v", pair)
			}
		}
		boundary := strings.ReplaceAll(source, `<c:min val="-10"/>`, `<c:min val="0"/>`)
		boundary = strings.ReplaceAll(boundary, `<c:majorTickMark val="none"/>`, `<c:majorTickMark val="out"/>`)
		if got := extractNativeLiteralBar([]byte(boundary), "edge.xml", d); got == nil || !validNativeLiteralBar(got) {
			t.Fatal("explicit plot-edge ticks rejected")
		}
		chart.CategoryAxis.Labels.MajorUnit = chart.ValueAxis.Labels.MajorUnit
		if validNativeLiteralBar(chart) {
			t.Fatal("category tick unit accepted")
		}
	}
}

func nativeLabeledConnectedXML(strict, scatter bool) string {
	source := nativeConnectedXML(strict, scatter)
	source = strings.ReplaceAll(source, `<c:tickLblPos val="none"/>`, `<c:tickLblPos val="high"/>`)
	source = strings.ReplaceAll(source, `</c:spPr><c:crossAx`, `</c:spPr>`+nativeAxisTextXML(strict)+`<c:crossAx`)
	for _, pos := range []string{"l", "b"} {
		if pos == "b" && !scatter {
			continue
		}
		source = strings.Replace(source, `<c:axPos val="`+pos+`"/>`, `<c:axPos val="`+pos+`"/><c:numFmt formatCode="0.0" sourceLinked="0"/>`, 1)
	}
	source = strings.ReplaceAll(source, `</c:valAx>`, `<c:majorUnit val="5"/></c:valAx>`)
	offset := "0"
	if strict {
		offset = "0%"
	}
	source = strings.Replace(source, `</c:catAx>`, `<c:auto val="0"/><c:lblAlgn val="ctr"/><c:lblOffset val="`+offset+`"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/><c:noMultiLvlLbl val="1"/></c:catAx>`, 1)
	return source
}
func TestNativeChartAxisLabelsExtractionOwnership(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, source := range []string{nativeLabeledBarXML(strict), nativeLabeledConnectedXML(strict, false), nativeLabeledConnectedXML(strict, true)} {
			payload := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: source})
			deck, e := ExtractNativePPTX(payload, nativeTestExtractOptions())
			if e != nil {
				t.Fatal(e)
			}
			chart := nativeFixtureChart(t, deck.Slides[0])
			if chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || chart.Chart.LiteralBar == nil && chart.Chart.LiteralConnected == nil {
				t.Fatal("source labels lost read-only profile")
			}
			encoded, e := MarshalNativePPTXJSON(deck)
			if e != nil {
				t.Fatal(e)
			}
			if _, e = DecodeNativePPTXJSON(encoded); e != nil {
				t.Fatal(e)
			}
		}
	}
}
