package pptxpatch

func extractNativeLiteralArea(payload []byte, part string, d nativeExtractDialect) *NativeLiteralArea {
	source := extractNativeChartArea(payload, part, d)
	if source == nil || source.XAxis.CategoryCrossing != "min" {
		return nil
	}
	out := &NativeLiteralArea{Profile: "literal-area-v1", SourceBaseline: &NativeAreaSourceBaseline{Crossing: source.XAxis.CategoryCrossing, Value: source.YAxis.Min}, DataOrigin: "literal", Grouping: source.Grouping, Categories: source.Categories, Series: []NativeLiteralAreaSeries{}, XAxis: nativeLiteralChartAxis(source.XAxis, false), YAxis: nativeLiteralChartAxis(source.YAxis, true)}
	for _, s := range source.Series {
		out.Series = append(out.Series, NativeLiteralAreaSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, Color: s.Color})
	}
	return out
}
