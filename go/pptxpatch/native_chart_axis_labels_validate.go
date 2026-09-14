package pptxpatch

import (
	"strings"
	"unicode"
	"unicode/utf8"
)

func validNativeChartAxisLabels(axis, perpendicular NativeLiteralBarAxis, value bool) bool {
	labels := axis.Labels
	if labels == nil {
		return true
	}
	if axis.Deleted || labels.Profile != "explicit-axis-labels-v1" || (labels.Position != "low" && labels.Position != "high") || (labels.MajorTickMark != "none" && labels.MajorTickMark != "out") {
		return false
	}
	s := labels.Style
	if s.FontFamily == "" || len(s.FontFamily) > 128 || !utf8.ValidString(s.FontFamily) || strings.HasPrefix(s.FontFamily, "+") || strings.IndexFunc(s.FontFamily, unicode.IsControl) >= 0 || s.FontSize < 1 || s.FontSize > 400000 || !validNativeLanguage(s.Language) || !inspectionRGB.MatchString(s.Color) || strings.ToUpper(s.Color) != s.Color {
		return false
	}
	if value {
		if labels.MajorUnit == nil || labels.NumberFormat == nil || axis.Min == nil || axis.Max == nil || !nativeChartFixedFormat(*labels.NumberFormat) {
			return false
		}
		if _, ok := nativeChartAxisTickCount(*axis.Min, *axis.Max, *labels.MajorUnit); !ok {
			return false
		}
	} else if labels.MajorUnit != nil || labels.NumberFormat != nil {
		return false
	}
	// Outside ticks require an actual plot-edge axis. An interior zero crossing
	// has no unambiguous outward direction in this bounded preview policy.
	if labels.MajorTickMark == "out" && perpendicular.Min != nil {
		if perpendicular.Max == nil {
			return false
		}
		low, e := parseNativeChartDecimal(*perpendicular.Min)
		if e != nil {
			return false
		}
		high, e := parseNativeChartDecimal(*perpendicular.Max)
		if e != nil {
			return false
		}
		if !low.isZero() && !high.isZero() {
			return false
		}
	}
	return true
}
