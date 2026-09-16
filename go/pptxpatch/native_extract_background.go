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

func resolveNativeInheritedBackground(graph nativeSlideDependencyGraph, slideRoot *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) string {
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
		return ""
	}
	return color
}

// nativePartBackground reports whether this part authors a background at all,
// and the colour when one can be resolved.
func nativePartBackground(root *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (string, bool) {
	if root == nil {
		return "", false
	}
	common, err := nativeSingleton(root, dialect.presentation, "cSld", false)
	if err != nil || common == nil {
		return "", false
	}
	background, err := nativeSingleton(common, dialect.presentation, "bg", false)
	if err != nil || background == nil {
		return "", false
	}
	return nativeBackgroundNodeColor(background, dialect, theme), true
}

// nativeBackgroundNodeColor resolves one p:bg element, whichever part it is on.
func nativeBackgroundNodeColor(background *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) string {
	if background == nil || requireOnlyNativeAttrs(background) != nil {
		return ""
	}
	if properties, err := nativeSingleton(background, dialect.presentation, "bgPr", false); err == nil && properties != nil {
		return nativeBackgroundProperties(properties, dialect, theme)
	}
	if reference, err := nativeSingleton(background, dialect.presentation, "bgRef", false); err == nil && reference != nil {
		return nativeBackgroundReference(reference, dialect, theme)
	}
	return ""
}

// nativeBackgroundProperties resolves p:bgPr, an explicit DrawingML fill. Only a
// solid fill is approximated. bwMode is a display hint and a:effectLst is
// accepted when empty, because PowerPoint writes one on essentially every
// authored background and an empty effect list changes no pixel.
func nativeBackgroundProperties(node *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) string {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "bwMode"}) != nil {
		return ""
	}
	for _, child := range node.Children {
		if child.Name.Space != dialect.drawing {
			return ""
		}
		switch child.Name.Local {
		case "solidFill":
		case "effectLst":
			if len(child.Children) != 0 || requireOnlyNativeAttrs(child) != nil {
				return ""
			}
		default:
			// gradFill, blipFill, pattFill, noFill: not approximated.
			return ""
		}
	}
	fill, err := nativeSingleton(node, dialect.drawing, "solidFill", false)
	if err != nil || fill == nil {
		return ""
	}
	color, err := exactNativeSolidColor(fill, dialect, theme)
	if err != nil {
		return ""
	}
	return color
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
