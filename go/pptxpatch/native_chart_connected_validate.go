package pptxpatch

import (
	"strings"
	"unicode/utf16"
)

func validNativeLiteralConnected(c *NativeLiteralConnected) bool {
	if c == nil || (c.Profile != "literal-line-v1" && c.Profile != "literal-scatter-v1") || c.DataOrigin != "literal" || len(c.Series) < 1 || len(c.Series) > 16 || c.Categories == nil {
		return false
	}
	scatter := c.Profile == "literal-scatter-v1"
	if scatter {
		if len(c.Categories) != 0 {
			return false
		}
	} else if len(c.Categories) < 1 || len(c.Categories) > 256 {
		return false
	}
	units := 0
	for _, v := range c.Categories {
		units += len(utf16.Encode([]rune(v)))
		if units > 32768 {
			return false
		}
	}
	color := func(s string) bool {
		return len(s) == 7 && s[0] == '#' && inspectionRGB.MatchString(s[1:]) && strings.ToUpper(s) == s
	}
	x, y := c.XAxis, c.YAxis
	if x.ID == y.ID || x.CrossAxisID != y.ID || y.CrossAxisID != x.ID || x.Position != "b" || y.Position != "l" {
		return false
	}
	for i, a := range []NativeLiteralBarAxis{x, y} {
		if a.ID < 0 || a.ID > 4294967295 || a.CrossAxisID < 0 || a.CrossAxisID > 4294967295 || (a.Orientation != "minMax" && a.Orientation != "maxMin") {
			return false
		}
		if a.Deleted {
			if a.Color != nil || a.WidthEMU != nil {
				return false
			}
		} else if a.Color == nil || !color(*a.Color) || a.WidthEMU == nil || *a.WidthEMU < 1 || *a.WidthEMU > 20116800 {
			return false
		}
		if i == 0 && !scatter {
			if a.Min != nil || a.Max != nil || a.CrossesAt != nil {
				return false
			}
			continue
		}
		if a.Min == nil || a.Max == nil || a.CrossesAt == nil {
			return false
		}
		low, e := parseNativeChartDecimal(*a.Min)
		if e != nil {
			return false
		}
		high, e := parseNativeChartDecimal(*a.Max)
		if e != nil {
			return false
		}
		cross, e := parseNativeChartDecimal(*a.CrossesAt)
		if e != nil || !cross.isZero() {
			return false
		}
		zero, _ := parseNativeChartDecimal("0")
		if low.compare(high) >= 0 || low.compare(zero) > 0 || high.compare(zero) < 0 {
			return false
		}
	}
	seen := map[int64]bool{}
	for i, s := range c.Series {
		if s.Index < 0 || s.Index > 4294967295 || seen[s.Index] || s.Order != int64(i) || len(s.Values) < 1 || len(s.Values) > 256 || !color(s.Color) || s.WidthEMU < 1 || s.WidthEMU > 20116800 || (s.Title != nil && len(utf16.Encode([]rune(*s.Title))) > 1024) {
			return false
		}
		seen[s.Index] = true
		if scatter {
			if len(s.XValues) != len(s.Values) {
				return false
			}
		} else if s.XValues != nil || len(s.Values) != len(c.Categories) {
			return false
		}
		for _, v := range s.Values {
			if _, e := parseNativeChartDecimal(v); e != nil {
				return false
			}
		}
		for _, v := range s.XValues {
			if _, e := parseNativeChartDecimal(v); e != nil {
				return false
			}
		}
	}
	return true
}
