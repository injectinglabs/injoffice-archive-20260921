package pptxpatch

import (
	"bytes"
	"slices"
	"strings"
	"testing"
)

// This source fixture intentionally orders XML points and series differently
// from their authored indices/orders. It is also reusable by the connected
// native/WASM fixture after the shared source attachment is available.
func nativeAreaPercentLexemeXML(strict bool) string {
	source := nativeAreaXML(strict, "percentStacked")
	start := strings.Index(source, "<c:ser>")
	end := strings.Index(source, "</c:ser>") + len("</c:ser>")
	first := source[start:end]
	first = strings.NewReplacer("<c:v>-1</c:v>", "<c:v>1e-100</c:v>", "<c:v>2.5</c:v>", "<c:v>-0</c:v>", "<c:v>1e0</c:v>", "<c:v>9007199254740993</c:v>").Replace(first)
	second := strings.NewReplacer(`idx val="5"`, `idx val="9"`, `order val="0"`, `order val="1"`, "123456", "ABCDEF", "<c:v>1e-100</c:v>", "<c:v>2e-100</c:v>", "<c:v>-0</c:v>", "<c:v>0.0</c:v>", "<c:v>9007199254740993</c:v>", "<c:v>18014398509481986</c:v>").Replace(first)
	return source[:start] + second + first + source[end:]
}

func TestNativeChartAreaPercentSourceLexemesAndZeroTotal(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		payload := []byte(nativeAreaPercentLexemeXML(strict))
		before := bytes.Clone(payload)
		area := extractNativeChartArea(payload, "chart.xml", d)
		if area == nil || len(area.Series) != 2 || !slices.Equal(area.Categories, []string{"A", "B", "C"}) {
			t.Fatal("qualified lexical source fixture was refused")
		}
		expected := [][]string{{"2e-100", "0.0", "18014398509481986"}, {"1e-100", "-0", "9007199254740993"}}
		series := make([]nativeChartStackSeries, 2)
		for i, item := range area.Series {
			if !slices.Equal(item.Values, expected[i]) {
				t.Fatal("source point order or numeric spelling changed", item.Values)
			}
			series[i] = nativeChartStackSeries{Index: item.Index, Order: item.Order, Values: item.Values}
		}
		bands, err := nativeChartStackBands(series, area.Grouping)
		if err != nil {
			t.Fatal(err)
		}
		for _, point := range []int{0, 2} {
			if bands[0].Upper[point].RatString() != "2/3" || bands[1].Lower[point].RatString() != "2/3" || bands[1].Upper[point].RatString() != "1" {
				t.Fatal("tiny/large exact source values lost their common ratio")
			}
		}
		for _, band := range bands {
			if band.Lower[1].Sign() != 0 || band.Upper[1].Sign() != 0 {
				t.Fatal("zero-total source category acquired a percentage segment")
			}
		}
		if !bytes.Equal(payload, before) {
			t.Fatal("source payload changed")
		}
	}
}
