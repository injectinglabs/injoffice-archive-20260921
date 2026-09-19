package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

// Slide background resolution.
//
// ECMA-376 makes p:bg optional at every level: a slide with none inherits its
// layout's, a layout with none inherits its master's, and a package with none at
// all falls back to the implicit default, schemeClr bg1. Previously only an
// explicit p:bg on the slide part itself was modelled, so a deck whose
// background lives on its layout or master painted white.
//
// Every scheme reference here resolves through the slide's EFFECTIVE colour map
// (master -> layout p:clrMapOvr -> slide p:clrMapOvr), not the map of the part
// the p:bg was authored in. A master background written as schemeClr bg1 under a
// slide that overrides bg1 to dk1 paints dark, which is what PowerPoint does.
//
// Anything outside the resolvable subset returns "" with no error: the caller
// keeps its existing refusal and the source stays preserved rather than guessed.

// bgRef indices are 1-based within their list; >= bgRefBackgroundBase selects
// a:bgFillStyleLst, 1..999 selects a:fillStyleLst, and 0 or 1000 mean no fill.
const bgRefBackgroundBase = 1001

// Bounds mirror $defs/linearGradient in schemas/pptx-native-v1.schema.json.
const nativeMaxGradientStops = 64
const nativeMaxPositiveFixedAngle = 21599999

// nativeResolvedBackground is one slide background. At most one field is set;
// both empty means this tier does not represent the authored background.
type nativeResolvedBackground struct {
	color    string
	gradient *NativeLinearGradient
}

func (background nativeResolvedBackground) empty() bool {
	return background.color == "" && background.gradient == nil
}

func resolveNativeInheritedBackground(graph nativeSlideDependencyGraph, slideRoot *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) nativeResolvedBackground {
	for _, root := range []*nativeXMLNode{slideRoot, graph.layoutRoot, graph.masterRoot} {
		background, found := nativePartBackground(root, dialect, theme)
		if !found {
			continue
		}
		// The nearest authored background wins even when this tier cannot
		// represent it; falling through would paint a more distant part's
		// background, which is worse than refusing.
		return background
	}
	// No p:bg at any level: the implicit default is bg1 through the effective map.
	color, err := theme.resolveSchemeColor("bg1")
	if err != nil {
		return nativeResolvedBackground{}
	}
	return nativeResolvedBackground{color: color}
}

// nativePartBackground reports whether this part authors a background at all,
// and the paint when one can be resolved.
func nativePartBackground(root *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (nativeResolvedBackground, bool) {
	if root == nil {
		return nativeResolvedBackground{}, false
	}
	common, err := nativeSingleton(root, dialect.presentation, "cSld", false)
	if err != nil || common == nil {
		return nativeResolvedBackground{}, false
	}
	background, err := nativeSingleton(common, dialect.presentation, "bg", false)
	if err != nil || background == nil {
		return nativeResolvedBackground{}, false
	}
	return nativeBackgroundNodePaint(background, dialect, theme), true
}

// nativeBackgroundNodePaint resolves one p:bg element, whichever part it is on.
func nativeBackgroundNodePaint(background *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) nativeResolvedBackground {
	if background == nil || requireOnlyNativeAttrs(background) != nil {
		return nativeResolvedBackground{}
	}
	if properties, err := nativeSingleton(background, dialect.presentation, "bgPr", false); err == nil && properties != nil {
		return nativeBackgroundProperties(properties, dialect, theme)
	}
	if reference, err := nativeSingleton(background, dialect.presentation, "bgRef", false); err == nil && reference != nil {
		return nativeResolvedBackground{color: nativeBackgroundReference(reference, dialect, theme)}
	}
	return nativeResolvedBackground{}
}

// nativeBackgroundProperties resolves p:bgPr, an explicit DrawingML fill. A
// solid fill and a linear a:gradFill are approximated; every other fill is
// preserved without one. bwMode is a display hint and a:effectLst is
// accepted when empty, because PowerPoint writes one on essentially every
// authored background and an empty effect list changes no pixel.
func nativeBackgroundProperties(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) nativeResolvedBackground {
	none := nativeResolvedBackground{}
	if requireOnlyNativeAttrs(node, xml.Name{Local: "bwMode"}) != nil {
		return none
	}
	for _, child := range node.Children {
		if child.Name.Space != dialect.drawing {
			return none
		}
		switch child.Name.Local {
		case "solidFill", "gradFill":
		case "effectLst":
			if len(child.Children) != 0 || requireOnlyNativeAttrs(child) != nil {
				return none
			}
		default:
			// blipFill, pattFill, noFill: not approximated.
			return none
		}
	}
	if fill, err := nativeSingleton(node, dialect.drawing, "gradFill", false); err == nil && fill != nil {
		if solid, solidErr := nativeSingleton(node, dialect.drawing, "solidFill", false); solidErr != nil || solid != nil {
			return none
		}
		gradient, gradientErr := nativeLinearGradientFill(fill, dialect, theme)
		if gradientErr != nil {
			return none
		}
		return nativeResolvedBackground{gradient: gradient}
	}
	fill, err := nativeSingleton(node, dialect.drawing, "solidFill", false)
	if err != nil || fill == nil {
		return none
	}
	color, err := exactNativeSolidColor(fill, dialect, theme)
	if err != nil {
		return none
	}
	return nativeResolvedBackground{color: color}
}

// nativeLinearGradientFill projects an a:gradFill whose direction is an a:lin.
//
// ECMA-376 Part 1 §20.1.8.33 defines a:gradFill, §20.1.8.36 the a:gsLst stop
// list, §20.1.8.41 the a:lin linear direction with @ang in 1/60000 of a degree
// clockwise from the positive x axis, and §20.1.8.55 @rotWithShape. Only the
// linear shading form is modeled: a:path (§20.1.8.46) shades along a rectangle
// or shape outline, and a:tileRect, flip and gradient rotation change the
// mapping in ways this tier does not represent, so each one refuses instead of
// painting a straight interpolation that would not be the authored fill.
//
// @scaled (§20.1.8.41) selects whether the angle is measured in the shape's
// own scaled space or unscaled; it is projected as the disclosed angle either
// way because a slide background fills the page and its two interpretations
// coincide only for a square page. That is why this projection is confined to
// the slide background, where the painted box is exactly the page.
func nativeLinearGradientFill(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*NativeLinearGradient, error) {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "rotWithShape"}, xml.Name{Local: "flip"}) != nil {
		return nil, fmt.Errorf("gradient fill attributes are outside the exact native subset")
	}
	if _, ok := exactNativeAttr(node, "", "flip"); ok {
		return nil, fmt.Errorf("flipped gradients are preserved but not approximated")
	}
	if requireOnlyNativeChildren(node,
		xml.Name{Space: dialect.drawing, Local: "gsLst"},
		xml.Name{Space: dialect.drawing, Local: "lin"}) != nil || !onlyNativeXMLSpace(node.Text) {
		return nil, fmt.Errorf("only a linear gradient with an explicit stop list is representable")
	}
	stopList, err := nativeSingleton(node, dialect.drawing, "gsLst", true)
	if err != nil {
		return nil, err
	}
	direction, err := nativeSingleton(node, dialect.drawing, "lin", true)
	if err != nil {
		return nil, err
	}
	if requireOnlyNativeAttrs(direction, xml.Name{Local: "ang"}, xml.Name{Local: "scaled"}) != nil || requireOnlyNativeChildren(direction) != nil {
		return nil, fmt.Errorf("linear gradient direction is outside the exact native subset")
	}
	angleValue, ok := exactNativeAttr(direction, "", "ang")
	if !ok {
		return nil, fmt.Errorf("linear gradient declares no angle")
	}
	angle, err := parseCanonicalNativeInt(angleValue, 0, nativeMaxPositiveFixedAngle)
	if err != nil {
		return nil, err
	}
	if requireOnlyNativeAttrs(stopList) != nil || !onlyNativeXMLSpace(stopList.Text) {
		return nil, fmt.Errorf("gradient stop list is outside the exact native subset")
	}
	if len(stopList.Children) < 2 || len(stopList.Children) > nativeMaxGradientStops {
		return nil, fmt.Errorf("gradient stop list must hold 2..%d stops", nativeMaxGradientStops)
	}
	stops := make([]NativeLinearGradientStop, 0, len(stopList.Children))
	previous := int64(-1)
	for _, child := range stopList.Children {
		if child.Name != (xml.Name{Space: dialect.drawing, Local: "gs"}) {
			return nil, fmt.Errorf("gradient stop list contains unmodeled markup")
		}
		if requireOnlyNativeAttrs(child, xml.Name{Local: "pos"}) != nil || !onlyNativeXMLSpace(child.Text) {
			return nil, fmt.Errorf("gradient stop attributes are outside the exact native subset")
		}
		positionValue, ok := exactNativeAttr(child, "", "pos")
		if !ok {
			return nil, fmt.Errorf("gradient stop declares no position")
		}
		position, err := parseCanonicalNativeInt(positionValue, 0, nativePositiveFixedPct)
		if err != nil {
			return nil, err
		}
		if position <= previous {
			return nil, fmt.Errorf("gradient stops are not in strictly increasing position order")
		}
		previous = position
		// a:gs carries @pos alongside its one color child, so the color is read
		// through an owned attribute-free view of the same children.
		color, err := exactNativeSolidColor(&nativeXMLNode{Children: child.Children, Text: child.Text}, dialect, theme)
		if err != nil {
			return nil, err
		}
		stops = append(stops, NativeLinearGradientStop{PositionPct: position, Color: color})
	}
	return &NativeLinearGradient{Angle: angle, Stops: stops}, nil
}

// nativeBackgroundReference resolves p:bgRef, a style-matrix reference: an index
// into the theme's background fill styles plus the colour substituted for phClr.
func nativeBackgroundReference(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) string {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "idx"}) != nil {
		return ""
	}
	raw, ok := exactNativeAttr(node, "", "idx")
	if !ok {
		return ""
	}
	index, err := strconv.Atoi(raw)
	if err != nil || strconv.Itoa(index) != raw || index < bgRefBackgroundBase {
		// 0 and 1000 are no-fill; 1..999 indexes a:fillStyleLst, which this tier
		// does not model for backgrounds.
		return ""
	}
	slot := index - bgRefBackgroundBase
	if slot >= len(theme.bgFillStyles) || theme.bgFillStyles[slot] == nil {
		return ""
	}
	color, err := nativeBackgroundReferenceColor(node, dialect, theme)
	if err != nil {
		return ""
	}
	// The selected entry is a solid fill whose colour is a:phClr, and the
	// reference's own colour is substituted for it. An entry naming anything
	// else is a literal this tier does not approximate.
	placeholder, err := nativeSingleton(theme.bgFillStyles[slot], dialect.drawing, "schemeClr", false)
	if err != nil || placeholder == nil {
		return ""
	}
	if value, attrOK := exactNativeAttr(placeholder, "", "val"); !attrOK || value != "phClr" {
		return ""
	}
	if len(placeholder.Children) != 0 {
		// Transforms on the placeholder would have to be applied to the
		// substituted colour; not modelled, so refuse rather than approximate.
		return ""
	}
	return color
}

func nativeBackgroundReferenceColor(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (string, error) {
	if scheme, err := nativeSingleton(node, dialect.drawing, "schemeClr", false); err == nil && scheme != nil {
		value, ok := exactNativeAttr(scheme, "", "val")
		if !ok {
			return "", fmt.Errorf("background reference scheme color has no value")
		}
		if len(scheme.Children) != 0 {
			return "", fmt.Errorf("background reference color transforms are not approximated")
		}
		return theme.resolveSchemeColor(value)
	}
	if srgb, err := nativeSingleton(node, dialect.drawing, "srgbClr", false); err == nil && srgb != nil {
		value, ok := exactNativeAttr(srgb, "", "val")
		if !ok || !colorPattern.MatchString(value) {
			return "", fmt.Errorf("background reference sRGB color is not exact")
		}
		return value, nil
	}
	return "", fmt.Errorf("background reference has no representable color")
}
