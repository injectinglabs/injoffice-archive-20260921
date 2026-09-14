package pptxpatch

// Geometry qualification remains narrower than syntax extraction: labels and
// out-of-scale points retain their source but require subsequent radial policy.
func validNativeRadarGeometrySource(c *nativeChartRadarSource) bool {
	if c == nil || c.Workbook || (c.Style != "standard" && c.Style != "filled") || len(c.Series) < 1 || len(c.Series) > 16 {
		return false
	}
	x, y := c.CategoryAxis, c.ValueAxis
	if x.Labels != nil || y.Labels != nil || x.Position != "b" || y.Position != "l" || x.ID == y.ID || x.CrossAxisID != y.ID || y.CrossAxisID != x.ID {
		return false
	}
	low, err := parseNativeChartDecimal(y.Min)
	if err != nil {
		return false
	}
	high, err := parseNativeChartDecimal(y.Max)
	if err != nil || low.compare(high) >= 0 {
		return false
	}
	for _, s := range c.Series {
		if len(s.Values) < 3 || len(s.Values) > 256 || len(s.Values) != len(s.Categories) || s.CategoryReference != nil || s.ValueReference != nil || s.TitleReference != nil {
			return false
		}
		for _, raw := range s.Values {
			value, err := parseNativeChartDecimal(raw)
			if err != nil || value.compare(low) < 0 || value.compare(high) > 0 {
				return false
			}
		}
	}
	return true
}
