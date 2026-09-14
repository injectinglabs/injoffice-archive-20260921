package pptxpatch

func extractNativeLiteralBubble(payload []byte, part string, d nativeExtractDialect) *NativeLiteralBubble {
	source := extractNativeChartBubbleSource(payload, part, d, false)
	if source == nil {
		return nil
	}
	out := &NativeLiteralBubble{Profile: "literal-bubble-v1", DataOrigin: "literal", BubbleScale: source.BubbleScale, SizeRepresents: source.SizeRepresents, Series: []NativeLiteralBubbleSeries{}, XAxis: nativeLiteralChartAxis(source.XAxis, true), YAxis: nativeLiteralChartAxis(source.YAxis, true)}
	for _, s := range source.Series {
		out.Series = append(out.Series, NativeLiteralBubbleSeries{Index: s.Index, Order: s.Order, Title: s.Title, XValues: s.XValues, Values: s.Values, Sizes: s.Sizes, Colors: s.Colors})
	}
	return out
}
