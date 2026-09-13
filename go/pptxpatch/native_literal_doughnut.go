package pptxpatch

func extractNativeLiteralDoughnut(payload []byte, part string, d nativeExtractDialect) *NativeLiteralDoughnut {
	source := extractNativeLiteralCircular(payload, part, d, "doughnutChart")
	if source == nil {
		return nil
	}
	return &NativeLiteralDoughnut{Profile: "literal-doughnut-v1", FirstSliceAngle: source.FirstSliceAngle, HoleSize: source.HoleSize, Values: source.Values, Colors: source.Colors}
}
