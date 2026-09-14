package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeLiteralBubbleAttachment(t *testing.T) {
	for _, strict := range []bool{false, true} {
		source := nativeBubbleXML(strict, false)
		deck, err := ExtractNativePPTX(nativeChartFixture(t, nativeChartFixtureOptions{strict: strict, omitPreview: true, chartXML: source}), nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		element := nativeFixtureChart(t, deck.Slides[0])
		bubble := element.Chart.LiteralBubble
		if bubble == nil || !validNativeLiteralBubble(bubble) || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || bubble.Series[0].Sizes[2] != "1e-100" {
			t.Fatal("source bubble not preserved/attached")
		}
		encoded, err := MarshalNativePPTXJSON(deck)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = DecodeNativePPTXJSON(encoded); err != nil {
			t.Fatal(err)
		}
		bubble.Series[0].Sizes[0] = "-1"
		if len(ValidateNativePPTX(deck)) == 0 {
			t.Fatal("negative sizes accepted")
		}
		bubble.Series[0].Sizes[0] = "1"
		element.Chart.LiteralArea = &NativeLiteralArea{}
		if len(ValidateNativePPTX(deck)) == 0 {
			t.Fatal("multiple literal families accepted")
		}
		zero := strings.ReplaceAll(source, `bubbleScale val="100`, `bubbleScale val="0`)
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		b := extractNativeLiteralBubble([]byte(zero), "chart.xml", d)
		if b == nil || b.BubbleScale != 0 || !validNativeLiteralBubble(b) {
			t.Fatal("explicit scale zero lost")
		}
	}
}
func TestNativeWorkbookBubbleDescriptorAttachment(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		raw := nativeBubbleXML(strict, true)
		source := extractNativeChartWorkbookSource([]byte(raw), "chart.xml", d)
		if source == nil || source.Family != "bubbleChart" || source.BubbleScale == nil || *source.BubbleScale != 100 {
			t.Fatal("workbook bubble source missing")
		}
		out := nativeWorkbookPublicSource(source)
		if out.Family != "bubble" || out.Series[0].SizeReference.Formula != "Data!C1:C3" || !out.Series[0].SizeReference.CachePresent || out.Series[0].WidthEMU != nil || len(out.Series[0].Colors) != 3 {
			t.Fatal("numeric size source/paint provenance lost")
		}
		if extractNativeLiteralBubble([]byte(raw), "chart.xml", d) != nil {
			t.Fatal("workbook rebranded as literal")
		}
	}
}
