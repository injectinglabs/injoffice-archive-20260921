package pptxpatch

import (
	"encoding/json"
	"testing"
)

func TestNativeStackedRecordValidation(t *testing.T) {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
	bar := extractNativeLiteralBar([]byte(nativeBarXML(false, true)), "chart.xml", d)
	if bar == nil {
		t.Fatal("missing baseline")
	}
	bar.Profile = "literal-stacked-bar-v1"
	bar.Grouping = "percentStacked"
	bar.Overlap = 100
	second := bar.Series[0]
	second.Index = 9
	second.Order = 1
	bar.Series = append([]NativeLiteralBarSeries{second}, bar.Series...)
	before, _ := json.Marshal(bar)
	if !validNativeStackedBarRecord(bar) {
		t.Fatal("signed stack refused")
	}
	after, _ := json.Marshal(bar)
	if string(before) != string(after) {
		t.Fatal("validation mutated source record")
	}
	bad := *bar
	bad.Overlap = 0
	if validNativeStackedBarRecord(&bad) {
		t.Fatal("partial overlap admitted")
	}
	bad = *bar
	bad.DataOrigin = "reference"
	if validNativeStackedBarRecord(&bad) {
		t.Fatal("cache authority upgraded")
	}
	source := extractNativeChartLine([]byte(nativeConnectedXML(false, false)), "chart.xml", d)
	if source == nil {
		t.Fatal("missing line source")
	}
	line := &nativeStackedLineRecord{Grouping: "stacked", NativeLiteralConnected: NativeLiteralConnected{Profile: "literal-stacked-line-v1", DataOrigin: "literal", Categories: source.Categories, XAxis: nativeLiteralChartAxis(source.XAxis, false), YAxis: nativeLiteralChartAxis(source.YAxis, true)}}
	for _, s := range source.Series {
		line.Series = append(line.Series, NativeLiteralConnectedSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, Color: s.Color, WidthEMU: s.Width})
	}
	if !validNativeStackedLineRecord(line) {
		t.Fatal("signed cumulative line refused")
	}
	next := line.Series[0]
	next.Index = 9
	next.Order = 1
	line.Series = append([]NativeLiteralConnectedSeries{next}, line.Series...)
	before, _ = json.Marshal(line)
	if !validNativeStackedLineRecord(line) {
		t.Fatal("XML-order line refused")
	}
	after, _ = json.Marshal(line)
	if string(before) != string(after) {
		t.Fatal("original order metadata mutated")
	}
	line.Grouping = "standard"
	if validNativeStackedLineRecord(line) {
		t.Fatal("wrong profile admitted")
	}
}
