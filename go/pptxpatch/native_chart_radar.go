package pptxpatch

import "strings"

func extractNativeLiteralRadar(payload []byte, part string, d nativeExtractDialect) *NativeLiteralRadar {
	source := extractNativeChartRadarSource(payload, part, d, false)
	if source == nil {
		return nil
	}
	out := &NativeLiteralRadar{Profile: "literal-radar-v1", DataOrigin: "literal", Style: source.Style, Categories: source.Series[0].Categories, Series: []NativeLiteralRadarSeries{}, CategoryAxis: nativeLiteralChartAxis(source.CategoryAxis, false), ValueAxis: nativeLiteralChartAxis(source.ValueAxis, true)}
	for _, s := range source.Series {
		out.Series = append(out.Series, NativeLiteralRadarSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, Color: s.Color, WidthEMU: s.Width, Fill: s.Fill})
	}
	if !validNativeLiteralRadar(out) {
		return nil
	}
	return out
}

// Reuse the exact category/value/axis subset through an immutable value
// projection; no source bytes, metadata or series sequence is normalized.
func validNativeLiteralRadar(c *NativeLiteralRadar) bool {
	if c == nil || c.Profile != "literal-radar-v1" || c.DataOrigin != "literal" || (c.Style != "standard" && c.Style != "filled") || len(c.Categories) < 3 || len(c.Categories) > 256 || len(c.Series) < 1 || len(c.Series) > 16 || c.CategoryAxis.Labels != nil || c.ValueAxis.Labels != nil {
		return false
	}
	projected := &NativeLiteralArea{Profile: "literal-area-v1", DataOrigin: "literal", Grouping: "standard", Categories: c.Categories, XAxis: c.CategoryAxis, YAxis: c.ValueAxis, Series: []NativeLiteralAreaSeries{}}
	for _, s := range c.Series {
		if s.WidthEMU < 1 || s.WidthEMU > 20116800 {
			return false
		}
		if c.Style == "standard" && s.Fill != nil {
			return false
		}
		if c.Style == "filled" && (s.Fill == nil || len(*s.Fill) != 7 || (*s.Fill)[0] != '#' || !inspectionRGB.MatchString((*s.Fill)[1:]) || strings.ToUpper(*s.Fill) != *s.Fill) {
			return false
		}
		projected.Series = append(projected.Series, NativeLiteralAreaSeries{Index: s.Index, Order: s.Order, Title: s.Title, Values: s.Values, Color: s.Color})
	}
	if !validNativeLiteralArea(projected) {
		return false
	}
	low, _ := parseNativeChartDecimal(*c.ValueAxis.Min)
	high, _ := parseNativeChartDecimal(*c.ValueAxis.Max)
	for _, s := range c.Series {
		for _, raw := range s.Values {
			v, e := parseNativeChartDecimal(raw)
			if e != nil || v.compare(low) < 0 || v.compare(high) > 0 {
				return false
			}
		}
	}
	return true
}
