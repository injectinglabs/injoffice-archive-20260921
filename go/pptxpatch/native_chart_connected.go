package pptxpatch

func extractNativeLiteralConnected(payload []byte, part string, d nativeExtractDialect) *NativeLiteralConnected {
	source := extractNativeChartLine(payload, part, d)
	profile := "literal-line-v1"
	if source == nil {
		source = extractNativeChartScatter(payload, part, d)
		profile = "literal-scatter-v1"
	}
	if source == nil {
		return nil
	}

	out := &NativeLiteralConnected{Profile: profile, DataOrigin: "literal", Categories: source.Categories, Series: []NativeLiteralConnectedSeries{}, XAxis: nativeLiteralChartAxis(source.XAxis, profile == "literal-scatter-v1"), YAxis: nativeLiteralChartAxis(source.YAxis, true)}
	for _, s := range source.Series {
		out.Series = append(out.Series, NativeLiteralConnectedSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, XValues: s.XValues, Color: s.Color, WidthEMU: s.Width})
	}
	return out
}
