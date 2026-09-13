package pptxpatch

// Axis attachment follows IDs and axis position, never XML list order. Numeric
// XY axes each qualify explicit zero crossings and no category-only clauses.
func extractNativeChartConnectedAxes(plot *nativeChartCursor, d nativeExtractDialect, scatter bool, firstID, secondID int64) ([2]nativeChartAxis, bool) {
	var result [2]nativeChartAxis
	seen := map[string]bool{}
	for plot.has("catAx") || plot.has("valAx") {
		isValue := plot.has("valAx")
		kind := "catAx"
		if isValue {
			kind = "valAx"
		}
		if scatter && !isValue {
			return result, false
		}
		axis, ok := extractNativeChartAxisWithCrossBetween(plot.take(kind), d, isValue, !scatter)
		if !ok {
			return result, false
		}
		if axis.Position != "b" && axis.Position != "l" || seen[axis.Position] {
			return result, false
		}
		seen[axis.Position] = true
		if !scatter && ((axis.Position == "b" && isValue) || (axis.Position == "l" && !isValue)) {
			return result, false
		}
		if axis.Position == "b" {
			result[0] = *axis
		} else {
			result[1] = *axis
		}
	}
	x, y := result[0], result[1]
	if len(seen) != 2 || x.ID == y.ID || x.CrossAxisID != y.ID || y.CrossAxisID != x.ID || !((x.ID == firstID && y.ID == secondID) || (x.ID == secondID && y.ID == firstID)) {
		return result, false
	}
	return result, true
}
