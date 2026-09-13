package pptxpatch

import (
	"strings"
	"unicode/utf16"
)

func validNativeLiteralBar(bar *NativeLiteralBar) bool {
	if bar == nil || bar.Profile != "literal-bar-v1" || bar.Grouping != "clustered" || bar.DataOrigin != "literal" || (bar.BarDirection != "column" && bar.BarDirection != "bar") || bar.GapWidth < 0 || bar.GapWidth > 500 || bar.Overlap != 0 || len(bar.Categories) < 1 || len(bar.Categories) > nativeChartMaxCategories || len(bar.Series) < 1 || len(bar.Series) > nativeChartMaxSeries {
		return false
	}
	units := 0
	for _, category := range bar.Categories {
		units += len(utf16.Encode([]rune(category)))
		if units > nativeChartMaxCategoryUnits {
			return false
		}
	}
	color := func(c string) bool {
		return len(c) == 7 && c[0] == '#' && inspectionRGB.MatchString(c[1:]) && strings.ToUpper(c) == c
	}
	ca, va := bar.CategoryAxis, bar.ValueAxis
	for _, axis := range []NativeLiteralBarAxis{ca, va} {
		if axis.ID < 0 || axis.ID > 4294967295 || axis.CrossAxisID < 0 || axis.CrossAxisID > 4294967295 || (axis.Orientation != "minMax" && axis.Orientation != "maxMin") {
			return false
		}
		if axis.Deleted {
			if axis.Color != nil || axis.WidthEMU != nil {
				return false
			}
		} else {
			if axis.Color == nil || !color(*axis.Color) || axis.WidthEMU == nil || *axis.WidthEMU < 1 || *axis.WidthEMU > 20116800 {
				return false
			}
		}
	}
	if ca.ID == va.ID || ca.CrossAxisID != va.ID || va.CrossAxisID != ca.ID || ca.Min != nil || ca.Max != nil || ca.CrossesAt != nil || va.Min == nil || va.Max == nil || va.CrossesAt == nil {
		return false
	}
	if bar.BarDirection == "column" {
		if ca.Position != "b" || va.Position != "l" {
			return false
		}
	} else {
		if ca.Position != "l" || va.Position != "b" {
			return false
		}
	}
	minimum, err := parseNativeChartDecimal(*va.Min)
	if err != nil {
		return false
	}
	maximum, err := parseNativeChartDecimal(*va.Max)
	if err != nil {
		return false
	}
	zero, _ := parseNativeChartDecimal("0")
	cross, err := parseNativeChartDecimal(*va.CrossesAt)
	if err != nil || !cross.isZero() || minimum.compare(maximum) >= 0 || minimum.compare(zero) > 0 || maximum.compare(zero) < 0 {
		return false
	}
	indices := map[int64]bool{}
	for i, series := range bar.Series {
		if series.Index < 0 || series.Index > 4294967295 || indices[series.Index] || series.Order != int64(i) || len(series.Values) != len(bar.Categories) || len(series.Colors) != len(bar.Categories) || (series.Title != nil && len(utf16.Encode([]rune(*series.Title))) > 1024) {
			return false
		}
		indices[series.Index] = true
		for j, value := range series.Values {
			if _, err := parseNativeChartDecimal(value); err != nil || !color(series.Colors[j]) {
				return false
			}
		}
	}
	return true
}
