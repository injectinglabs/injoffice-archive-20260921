package pptxpatch

func extractNativeLiteralPie(payload []byte, part string, d nativeExtractDialect) *NativeLiteralPie {
	source := extractNativeLiteralCircular(payload, part, d, "pieChart")
	if source == nil {
		return nil
	}
	return &NativeLiteralPie{Profile: "literal-pie-v1", FirstSliceAngle: source.FirstSliceAngle, Values: source.Values, Colors: source.Colors}
}
