package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

var nativeThemeColorSlots = []string{
	"dk1", "lt1", "dk2", "lt2",
	"accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
	"hlink", "folHlink",
}

var nativeThemeColorMapSlots = []string{
	"bg1", "tx1", "bg2", "tx2",
	"accent1", "accent2", "accent3", "accent4", "accent5", "accent6",
	"hlink", "folHlink",
}

type nativeThemeFonts struct {
	majorLatin string
	minorLatin string
	majorEA    string
	minorEA    string
	majorCS    string
	minorCS    string
}

type nativeResolvedTheme struct {
	fonts  nativeThemeFonts
	colors map[string]string
	clrMap map[string]string
}

func (theme nativeResolvedTheme) resolveTypeface(value string) (string, error) {
	if value == "" {
		return "", fmt.Errorf("empty typeface")
	}
	if !strings.HasPrefix(value, "+") {
		return value, nil
	}
	resolved := ""
	switch value {
	case "+mj-lt":
		resolved = theme.fonts.majorLatin
	case "+mn-lt":
		resolved = theme.fonts.minorLatin
	case "+mj-ea":
		resolved = firstNonEmptyNativeTypeface(theme.fonts.majorEA, theme.fonts.majorLatin)
	case "+mn-ea":
		resolved = firstNonEmptyNativeTypeface(theme.fonts.minorEA, theme.fonts.minorLatin)
	case "+mj-cs":
		resolved = firstNonEmptyNativeTypeface(theme.fonts.majorCS, theme.fonts.majorLatin)
	case "+mn-cs":
		resolved = firstNonEmptyNativeTypeface(theme.fonts.minorCS, theme.fonts.minorLatin)
	default:
		return "", fmt.Errorf("unknown theme font token %q", value)
	}
	if resolved == "" {
		return "", fmt.Errorf("theme font token %q has no exact typeface", value)
	}
	return resolved, nil
}

func (theme nativeResolvedTheme) resolveSchemeColor(name string) (string, error) {
	if name == "" || name == "phClr" {
		return "", fmt.Errorf("placeholder or empty scheme color is not exact")
	}
	slot := name
	if mapped, ok := theme.clrMap[name]; ok {
		slot = mapped
	}
	color, ok := theme.colors[slot]
	if !ok || color == "" {
		return "", fmt.Errorf("scheme color %q has no exact sRGB snapshot", name)
	}
	return color, nil
}

func firstNonEmptyNativeTypeface(values ...string) string {
	for _, value := range values {
		if value != "" {
			return value
		}
	}
	return ""
}

func resolveNativeTheme(graph nativeSlideDependencyGraph, dialect nativeExtractDialect) (nativeResolvedTheme, error) {
	theme := nativeResolvedTheme{
		colors: map[string]string{},
		clrMap: map[string]string{},
	}
	if graph.themeRoot != nil {
		if err := parseNativeThemeElements(graph.themeRoot, dialect, &theme); err != nil {
			return nativeResolvedTheme{}, err
		}
	}
	if graph.masterRoot != nil {
		if err := parseNativeMasterColorMap(graph.masterRoot, dialect, &theme); err != nil {
			return nativeResolvedTheme{}, err
		}
	}
	return theme, nil
}

func parseNativeThemeElements(root *nativeXMLNode, dialect nativeExtractDialect, theme *nativeResolvedTheme) error {
	elements, err := nativeSingleton(root, dialect.drawing, "themeElements", false)
	if err != nil {
		return err
	}
	if elements == nil {
		return nil
	}
	if requireOnlyNativeAttrs(elements) != nil {
		return nil
	}
	scheme, err := nativeSingleton(elements, dialect.drawing, "clrScheme", false)
	if err != nil {
		return err
	}
	fonts, err := nativeSingleton(elements, dialect.drawing, "fontScheme", false)
	if err != nil {
		return err
	}
	if scheme != nil {
		if parseErr := parseNativeThemeColorScheme(scheme, dialect, theme); parseErr != nil {
			theme.colors = map[string]string{}
		}
	}
	if fonts != nil {
		if parseErr := parseNativeThemeFontScheme(fonts, dialect, theme); parseErr != nil {
			theme.fonts = nativeThemeFonts{}
		}
	}
	return nil
}

func parseNativeThemeColorScheme(node *nativeXMLNode, dialect nativeExtractDialect, theme *nativeResolvedTheme) error {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "name"}) != nil {
		return fmt.Errorf("theme color scheme attributes are outside the exact subset")
	}
	allowed := make([]xml.Name, 0, len(nativeThemeColorSlots))
	for _, slot := range nativeThemeColorSlots {
		allowed = append(allowed, xml.Name{Space: dialect.drawing, Local: slot})
	}
	if requireOnlyNativeChildren(node, allowed...) != nil {
		return fmt.Errorf("theme color scheme contains unmodeled slots")
	}
	colors := map[string]string{}
	for _, slot := range nativeThemeColorSlots {
		child, err := nativeSingleton(node, dialect.drawing, slot, true)
		if err != nil {
			return err
		}
		color, colorErr := exactNativeThemeSlotColor(child, dialect)
		if colorErr != nil {
			return colorErr
		}
		colors[slot] = color
	}
	theme.colors = colors
	return nil
}

func exactNativeThemeSlotColor(node *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	if requireOnlyNativeAttrs(node) != nil {
		return "", fmt.Errorf("theme color slot has unmodeled attributes")
	}
	srgb, err := nativeSingleton(node, dialect.drawing, "srgbClr", false)
	if err != nil {
		return "", err
	}
	sys, err := nativeSingleton(node, dialect.drawing, "sysClr", false)
	if err != nil {
		return "", err
	}
	if (srgb == nil) == (sys == nil) {
		return "", fmt.Errorf("theme color slot requires exactly one sRGB or sysClr snapshot")
	}
	if srgb != nil {
		return exactNativeSRGBColor(srgb)
	}
	if requireOnlyNativeAttrs(sys, xml.Name{Local: "val"}, xml.Name{Local: "lastClr"}) != nil || requireOnlyNativeChildren(sys) != nil {
		return "", fmt.Errorf("theme sysClr is not an exact lastClr snapshot")
	}
	last, ok := exactNativeAttr(sys, "", "lastClr")
	last = strings.ToUpper(last)
	if !ok || !colorPattern.MatchString(last) {
		return "", fmt.Errorf("theme sysClr is missing an exact lastClr snapshot")
	}
	return last, nil
}

func exactNativeSRGBColor(node *nativeXMLNode) (string, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(node) != nil {
		return "", fmt.Errorf("sRGB color is transformed or unmodeled")
	}
	value, ok := exactNativeAttr(node, "", "val")
	value = strings.ToUpper(value)
	if !ok || !colorPattern.MatchString(value) {
		return "", fmt.Errorf("invalid sRGB color")
	}
	return value, nil
}

func parseNativeThemeFontScheme(node *nativeXMLNode, dialect nativeExtractDialect, theme *nativeResolvedTheme) error {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "name"}) != nil {
		return fmt.Errorf("theme font scheme attributes are outside the exact subset")
	}
	if requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "majorFont"},
		xml.Name{Space: dialect.drawing, Local: "minorFont"}) != nil {
		return fmt.Errorf("theme font scheme contains unmodeled children")
	}
	major, err := nativeSingleton(node, dialect.drawing, "majorFont", true)
	if err != nil {
		return err
	}
	minor, err := nativeSingleton(node, dialect.drawing, "minorFont", true)
	if err != nil {
		return err
	}
	majorLatin, majorEA, majorCS, majorErr := parseNativeThemeFontSet(major, dialect)
	if majorErr != nil {
		return majorErr
	}
	minorLatin, minorEA, minorCS, minorErr := parseNativeThemeFontSet(minor, dialect)
	if minorErr != nil {
		return minorErr
	}
	if majorLatin == "" || minorLatin == "" {
		return fmt.Errorf("theme latin typefaces are empty")
	}
	theme.fonts = nativeThemeFonts{
		majorLatin: majorLatin, minorLatin: minorLatin,
		majorEA: majorEA, minorEA: minorEA,
		majorCS: majorCS, minorCS: minorCS,
	}
	return nil
}

func parseNativeThemeFontSet(node *nativeXMLNode, dialect nativeExtractDialect) (latin, ea, cs string, err error) {
	if requireOnlyNativeAttrs(node) != nil {
		return "", "", "", fmt.Errorf("theme font set has unmodeled attributes")
	}
	if requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "latin"},
		xml.Name{Space: dialect.drawing, Local: "ea"},
		xml.Name{Space: dialect.drawing, Local: "cs"},
		xml.Name{Space: dialect.drawing, Local: "font"}) != nil {
		return "", "", "", fmt.Errorf("theme font set contains unmodeled children")
	}
	latinNode, err := nativeSingleton(node, dialect.drawing, "latin", true)
	if err != nil {
		return "", "", "", err
	}
	eaNode, err := nativeSingleton(node, dialect.drawing, "ea", false)
	if err != nil {
		return "", "", "", err
	}
	csNode, err := nativeSingleton(node, dialect.drawing, "cs", false)
	if err != nil {
		return "", "", "", err
	}
	latin, err = exactNativeThemeTypeface(latinNode, true)
	if err != nil {
		return "", "", "", err
	}
	if eaNode != nil {
		ea, err = exactNativeThemeTypeface(eaNode, false)
		if err != nil {
			return "", "", "", err
		}
	}
	if csNode != nil {
		cs, err = exactNativeThemeTypeface(csNode, false)
		if err != nil {
			return "", "", "", err
		}
	}
	for _, child := range node.Children {
		if child.Name != (xml.Name{Space: dialect.drawing, Local: "font"}) {
			continue
		}
		if requireOnlyNativeAttrs(child, xml.Name{Local: "script"}, xml.Name{Local: "typeface"}) != nil || requireOnlyNativeChildren(child) != nil {
			return "", "", "", fmt.Errorf("theme script font mapping is malformed")
		}
		script, scriptOK := exactNativeAttr(child, "", "script")
		typeface, typefaceOK := exactNativeAttr(child, "", "typeface")
		if !scriptOK || script == "" || !typefaceOK || utf16CodeUnitLengthBounded(typeface, 1025) > 1024 {
			return "", "", "", fmt.Errorf("theme script font mapping is incomplete")
		}
	}
	return latin, ea, cs, nil
}

func exactNativeThemeTypeface(node *nativeXMLNode, requireNonEmpty bool) (string, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "typeface"}, xml.Name{Local: "panose"}) != nil || requireOnlyNativeChildren(node) != nil {
		return "", fmt.Errorf("theme typeface markup is outside the exact subset")
	}
	value, ok := exactNativeAttr(node, "", "typeface")
	if !ok {
		return "", fmt.Errorf("theme typeface is missing")
	}
	if requireNonEmpty && value == "" {
		return "", fmt.Errorf("theme latin typeface is empty")
	}
	if utf16CodeUnitLengthBounded(value, 1025) > 1024 {
		return "", fmt.Errorf("theme typeface exceeds the native contract bound")
	}
	return value, nil
}

func parseNativeMasterColorMap(root *nativeXMLNode, dialect nativeExtractDialect, theme *nativeResolvedTheme) error {
	mapping, err := nativeSingleton(root, dialect.presentation, "clrMap", false)
	if err != nil {
		return err
	}
	if mapping == nil {
		return nil
	}
	allowed := make([]xml.Name, 0, len(nativeThemeColorMapSlots))
	for _, slot := range nativeThemeColorMapSlots {
		allowed = append(allowed, xml.Name{Local: slot})
	}
	if requireOnlyNativeAttrs(mapping, allowed...) != nil || requireOnlyNativeChildren(mapping) != nil {
		return nil
	}
	clrMap := map[string]string{}
	for _, slot := range nativeThemeColorMapSlots {
		value, ok := exactNativeAttr(mapping, "", slot)
		if !ok || value == "" {
			return nil
		}
		clrMap[slot] = value
	}
	theme.clrMap = clrMap
	return nil
}

func exactNativeSolidColor(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (string, error) {
	if node == nil || requireOnlyNativeAttrs(node) != nil {
		return "", fmt.Errorf("invalid solid fill")
	}
	srgb, err := nativeSingleton(node, dialect.drawing, "srgbClr", false)
	if err != nil {
		return "", err
	}
	scheme, err := nativeSingleton(node, dialect.drawing, "schemeClr", false)
	if err != nil {
		return "", err
	}
	if (srgb == nil) == (scheme == nil) || requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "srgbClr"},
		xml.Name{Space: dialect.drawing, Local: "schemeClr"}) != nil {
		return "", fmt.Errorf("solid fill is not one exact sRGB or documented theme color")
	}
	if srgb != nil {
		return exactNativeSRGBColor(srgb)
	}
	return exactNativeSchemeColor(scheme, dialect, theme)
}

func exactNativeSchemeColor(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (string, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "val"}) != nil {
		return "", fmt.Errorf("scheme color attributes are unmodeled")
	}
	allowed := []xml.Name{
		{Space: dialect.drawing, Local: "tint"},
		{Space: dialect.drawing, Local: "shade"},
		{Space: dialect.drawing, Local: "lumMod"},
		{Space: dialect.drawing, Local: "lumOff"},
		{Space: dialect.drawing, Local: "alpha"},
		{Space: dialect.drawing, Local: "alphaMod"},
		{Space: dialect.drawing, Local: "alphaOff"},
	}
	if requireOnlyNativeChildren(node, allowed...) != nil {
		return "", fmt.Errorf("scheme color has unmodeled transforms")
	}
	name, ok := exactNativeAttr(node, "", "val")
	if !ok {
		return "", fmt.Errorf("scheme color is missing")
	}
	base, err := theme.resolveSchemeColor(name)
	if err != nil {
		return "", err
	}
	if len(node.Children) == 0 {
		return base, nil
	}
	return applyNativeSchemeColorTransforms(base, node, dialect)
}

const (
	nativeColorPercent     = int64(100000)
	nativeColorHueMax      = int64(21600000)
	nativePositiveFixedPct = int64(100000)
)

type nativeSRGBColor struct {
	r, g, b int64
}

func parseNativeSRGBHex(value string) (nativeSRGBColor, error) {
	if !colorPattern.MatchString(value) {
		return nativeSRGBColor{}, fmt.Errorf("invalid sRGB color")
	}
	parsed, err := strconv.ParseUint(value, 16, 32)
	if err != nil {
		return nativeSRGBColor{}, fmt.Errorf("invalid sRGB color")
	}
	return nativeSRGBColor{r: int64(parsed >> 16), g: int64((parsed >> 8) & 255), b: int64(parsed & 255)}, nil
}

func (color nativeSRGBColor) hex() string {
	return fmt.Sprintf("%02X%02X%02X", color.r, color.g, color.b)
}

func nativeRoundDiv(num, den int64) int64 {
	if den == 0 {
		return 0
	}
	if den < 0 {
		num, den = -num, -den
	}
	if num >= 0 {
		return (num + den/2) / den
	}
	return -((-num + den/2) / den)
}

func nativeClampInt64(value, minimum, maximum int64) int64 {
	if value < minimum {
		return minimum
	}
	if value > maximum {
		return maximum
	}
	return value
}

func applyNativeSchemeColorTransforms(base string, node *nativeXMLNode, dialect nativeExtractDialect) (string, error) {
	color, err := parseNativeSRGBHex(base)
	if err != nil {
		return "", err
	}
	applied := false
	for _, child := range node.Children {
		switch child.Name.Local {
		case "alpha", "alphaMod", "alphaOff":
			return "", fmt.Errorf("alpha color transforms are not represented in native sRGB")
		case "tint":
			value, valueErr := parseNativeColorTransformPercent(child, 0, nativePositiveFixedPct)
			if valueErr != nil {
				return "", valueErr
			}
			color = applyNativeTint(color, value)
			applied = true
		case "shade":
			value, valueErr := parseNativeColorTransformPercent(child, 0, nativePositiveFixedPct)
			if valueErr != nil {
				return "", valueErr
			}
			color = applyNativeShade(color, value)
			applied = true
		case "lumMod":
			value, valueErr := parseNativeColorTransformPercent(child, 0, nativeMaxDrawingPercentage)
			if valueErr != nil {
				return "", valueErr
			}
			h, s, l := nativeSRGBToHSL(color)
			l = nativeClampInt64(nativeRoundDiv(l*value, nativeColorPercent), 0, nativeColorPercent)
			if l == 0 || l == nativeColorPercent {
				s = 0
			}
			color = nativeHSLToSRGB(h, s, l)
			applied = true
		case "lumOff":
			value, valueErr := parseNativeColorTransformPercent(child, -nativeMaxDrawingPercentage, nativeMaxDrawingPercentage)
			if valueErr != nil {
				return "", valueErr
			}
			h, s, l := nativeSRGBToHSL(color)
			l = nativeClampInt64(l+value, 0, nativeColorPercent)
			if l == 0 || l == nativeColorPercent {
				s = 0
			}
			color = nativeHSLToSRGB(h, s, l)
			applied = true
		default:
			return "", fmt.Errorf("scheme color has unmodeled transforms")
		}
	}
	if !applied {
		return "", fmt.Errorf("scheme color transforms are not an exact documented sRGB formula")
	}
	return color.hex(), nil
}

func parseNativeColorTransformPercent(node *nativeXMLNode, minimum, maximum int64) (int64, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(node) != nil {
		return 0, fmt.Errorf("color transform is malformed")
	}
	value, ok := exactNativeAttr(node, "", "val")
	if !ok {
		return 0, fmt.Errorf("color transform is missing val")
	}
	parsed, err := parseCanonicalNativeInt(value, minimum, maximum)
	if err != nil {
		return 0, fmt.Errorf("color transform val is not a canonical percentage")
	}
	return parsed, nil
}

func applyNativeTint(color nativeSRGBColor, tint int64) nativeSRGBColor {
	// ECMA-376: a 10% tint is 10% of the input color combined with 90% white.
	mix := func(channel int64) int64 {
		return nativeClampInt64(nativeRoundDiv(channel*tint+255*(nativePositiveFixedPct-tint), nativePositiveFixedPct), 0, 255)
	}
	return nativeSRGBColor{r: mix(color.r), g: mix(color.g), b: mix(color.b)}
}

func applyNativeShade(color nativeSRGBColor, shade int64) nativeSRGBColor {
	mix := func(channel int64) int64 {
		return nativeClampInt64(nativeRoundDiv(channel*shade, nativePositiveFixedPct), 0, 255)
	}
	return nativeSRGBColor{r: mix(color.r), g: mix(color.g), b: mix(color.b)}
}

func nativeSRGBToHSL(color nativeSRGBColor) (h, s, l int64) {
	minC, maxC := color.r, color.r
	if color.g < minC {
		minC = color.g
	}
	if color.b < minC {
		minC = color.b
	}
	if color.g > maxC {
		maxC = color.g
	}
	if color.b > maxC {
		maxC = color.b
	}
	sum := minC + maxC
	l = nativeRoundDiv(sum*nativeColorPercent, 510)
	if maxC == minC {
		return 0, 0, l
	}
	delta := maxC - minC
	if l == 0 || l == nativeColorPercent {
		s = 0
	} else if l <= nativeColorPercent/2 {
		s = nativeRoundDiv(delta*nativeColorPercent, sum)
	} else {
		s = nativeRoundDiv(delta*nativeColorPercent, 510-sum)
	}
	var hue int64
	switch maxC {
	case color.r:
		hue = nativeRoundDiv((color.g-color.b)*60*60000, delta)
		if hue < 0 {
			hue += nativeColorHueMax
		}
	case color.g:
		hue = nativeRoundDiv((color.b-color.r)*60*60000, delta) + 120*60000
	default:
		hue = nativeRoundDiv((color.r-color.g)*60*60000, delta) + 240*60000
	}
	if hue < 0 {
		hue += nativeColorHueMax
	}
	h = hue % nativeColorHueMax
	return h, s, l
}

func nativeHSLToSRGB(h, s, l int64) nativeSRGBColor {
	if s < 0 {
		s = 0
	}
	if s > nativeColorPercent {
		s = nativeColorPercent
	}
	l = nativeClampInt64(l, 0, nativeColorPercent)
	if s == 0 || l == 0 || l == nativeColorPercent {
		value := nativeClampInt64(nativeRoundDiv(l*255, nativeColorPercent), 0, 255)
		if l == 0 {
			value = 0
		}
		if l == nativeColorPercent {
			value = 255
		}
		return nativeSRGBColor{r: value, g: value, b: value}
	}
	const scale int64 = 1_000_000_000
	sat := nativeRoundDiv(s*scale, nativeColorPercent)
	lum := nativeRoundDiv(l*scale, nativeColorPercent)
	hue6 := nativeRoundDiv(h*6*scale, nativeColorHueMax)
	if hue6 >= 6*scale {
		hue6 = 0
	}
	var red, green, blue int64
	switch {
	case hue6 <= scale:
		red, green, blue = scale, hue6, 0
	case hue6 <= 2*scale:
		red, green, blue = 2*scale-hue6, scale, 0
	case hue6 <= 3*scale:
		red, green, blue = 0, scale, hue6-2*scale
	case hue6 <= 4*scale:
		red, green, blue = 0, 4*scale-hue6, scale
	case hue6 <= 5*scale:
		red, green, blue = hue6-4*scale, 0, scale
	default:
		red, green, blue = scale, 0, 6*scale-hue6
	}
	channel := func(primary int64) int64 {
		component := nativeRoundDiv((primary-scale/2)*sat, scale) + scale/2
		shift := 2*lum - scale
		if shift < 0 {
			component = nativeRoundDiv(component*(shift+scale), scale)
		} else if shift > 0 {
			component = scale - nativeRoundDiv((scale-component)*(scale-shift), scale)
		}
		return nativeClampInt64(nativeRoundDiv(component*255, scale), 0, 255)
	}
	return nativeSRGBColor{r: channel(red), g: channel(green), b: channel(blue)}
}
