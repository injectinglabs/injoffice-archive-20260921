package pptxpatch

// Connected XY scatter keeps literal x/y point order, including duplicate or
// decreasing x. It does not sort points or reinterpret them as categories.
func extractNativeChartScatter(payload []byte, part string, d nativeExtractDialect) *nativeChartLine {
	return extractNativeChartConnected(payload, part, d, true)
}
