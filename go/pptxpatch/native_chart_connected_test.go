package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeChartConnectedPublicContract(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, scatter := range []bool{false, true} {
			payload := nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: nativeConnectedXML(strict, scatter)})
			deck, e := ExtractNativePPTX(payload, nativeTestExtractOptions())
			if e != nil {
				t.Fatal(e)
			}
			chart := nativeFixtureChart(t, deck.Slides[0])
			c := chart.Chart.LiteralConnected
			if !validNativeLiteralConnected(c) || chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || chart.Chart.LiteralBar != nil || chart.Chart.LiteralPie != nil || chart.Chart.LiteralDoughnut != nil {
				t.Fatal("missing exclusive preserved profile")
			}
			encoded, e := MarshalNativePPTXJSON(deck)
			if e != nil {
				t.Fatal(e)
			}
			if _, e = DecodeNativePPTXJSON(encoded); e != nil {
				t.Fatal(e)
			}
			if scatter {
				if len(c.Categories) != 0 || strings.Join(c.Series[0].XValues, ",") != "2,-1,2" {
					t.Fatal("source x order lost")
				}
			} else {
				c.Series[0].XValues = []string{}
				if validNativeLiteralConnected(c) {
					t.Fatal("line xValues accepted")
				}
				c.Series[0].XValues = nil
			}
			c.Profile = "literal-scatter-v1"
			c.Categories = []string{}
			c.Series[0].XValues = []string{"1"}
			if validNativeLiteralConnected(c) {
				t.Fatal("mismatched XY accepted")
			}
		}
	}
}
