package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeChartAreaPublicContract(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, source := range []string{nativeAreaXML(strict, "standard"), nativeAreaPercentLexemeXML(strict)} {
			payload := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: source})
			before := bytes.Clone(payload)
			deck, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			chart := nativeFixtureChart(t, deck.Slides[0])
			area := chart.Chart.LiteralArea
			if !validNativeLiteralArea(area) || chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || chart.Chart.LiteralConnected != nil || chart.Chart.LiteralBar != nil || chart.Chart.LiteralPie != nil || chart.Chart.LiteralDoughnut != nil {
				t.Fatal("missing exclusive preserved area profile")
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatal(err)
			}
			if _, err = DecodeNativePPTXJSON(encoded); err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(payload, before) {
				t.Fatal("source bytes changed")
			}
			areaOffset := bytes.Index(encoded, []byte(`"literalArea"`))
			if areaOffset < 0 {
				t.Fatal("fixture area missing")
			}
			brace := areaOffset + bytes.IndexByte(encoded[areaOffset:], '{') + 1
			malformed := append(bytes.Clone(encoded[:brace]), []byte(`"unknown":true,`)...)
			malformed = append(malformed, encoded[brace:]...)

			if _, err = DecodeNativePPTXJSON(malformed); err == nil {
				t.Fatal("unknown area property admitted")
			}
			if area.Grouping == "percentStacked" && strings.Join(area.Series[0].Values, ",") != "2e-100,0.0,18014398509481986" {
				t.Fatal("source lexemes lost in public profile")
			}
			chart.Chart.LiteralPie = &NativeLiteralPie{Profile: "literal-pie-v1", Values: []int64{1}, Colors: []string{"#000000"}}
			if _, err = MarshalNativePPTXJSON(deck); err == nil {
				t.Fatal("competing chart family admitted")
			}
		}
	}
}

func TestNativeChartAreaPublicRefusals(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	for _, mutate := range []func(*NativeLiteralArea){
		func(a *NativeLiteralArea) { a.Grouping = "stacked" },
		func(a *NativeLiteralArea) { a.DataOrigin = "embedded-workbook" },
		func(a *NativeLiteralArea) { a.Series[0].Values = []string{"1"} },
		func(a *NativeLiteralArea) { a.YAxis.CrossesAt = stringPointer("1") },
		func(a *NativeLiteralArea) { a.XAxis.CrossAxisID = 99 },
	} {
		a := extractNativeLiteralArea([]byte(nativeAreaXML(false, "standard")), "chart.xml", d)
		mutate(a)
		if validNativeLiteralArea(a) {
			t.Fatal("invalid public area admitted")
		}
	}
}
