package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

// Resolve a bounded solid fill/line style matrix into an owned paint view.
// Raw shape XML and source anchors are never rewritten by this projection.
func resolveNativeShapeStyle(properties, style, themeRoot *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*nativeXMLNode, error) {
	if style == nil {
		return properties, nil
	}
	if requireOnlyNativeAttrs(style) != nil || requireOnlyNativeChildren(style, xml.Name{Space: dialect.drawing, Local: "lnRef"}, xml.Name{Space: dialect.drawing, Local: "fillRef"}, xml.Name{Space: dialect.drawing, Local: "effectRef"}, xml.Name{Space: dialect.drawing, Local: "fontRef"}) != nil {
		return nil, fmt.Errorf("unmodeled shape style markup")
	}
	refs := map[string]*nativeXMLNode{}
	colors := map[string]string{}
	for _, name := range []string{"lnRef", "fillRef", "effectRef", "fontRef"} {
		ref, err := nativeSingleton(style, dialect.drawing, name, true)
		if err != nil {
			return nil, err
		}
		if requireOnlyNativeAttrs(ref, xml.Name{Local: "idx"}) != nil {
			return nil, fmt.Errorf("unmodeled style reference attributes")
		}
		color, err := exactNativeSolidColor(&nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "solidFill"}, Children: ref.Children, Text: ref.Text}, dialect, theme)
		if err != nil {
			return nil, err
		}
		refs[name] = ref
		colors[name] = color
	}
	fontIndex, _ := exactNativeAttr(refs["fontRef"], "", "idx")
	if fontIndex != "major" && fontIndex != "minor" {
		return nil, fmt.Errorf("unsupported font style reference")
	}
	effectIndex, _ := exactNativeAttr(refs["effectRef"], "", "idx")
	if effectIndex != "0" {
		return nil, fmt.Errorf("nonzero effect matrix reference remains unsupported")
	}
	elements, err := nativeSingleton(themeRoot, dialect.drawing, "themeElements", true)
	if err != nil {
		return nil, err
	}
	matrix, err := nativeSingleton(elements, dialect.drawing, "fmtScheme", true)
	if err != nil {
		return nil, err
	}
	result := *properties
	result.Children = append([]*nativeXMLNode(nil), properties.Children...)
	for _, item := range []struct {
		ref, list string
		names     []string
	}{{"fillRef", "fillStyleLst", []string{"solidFill", "noFill", "gradFill", "pattFill", "blipFill", "grpFill"}}, {"lnRef", "lnStyleLst", []string{"ln"}}} {
		value, _ := exactNativeAttr(refs[item.ref], "", "idx")
		index, err := parseCanonicalNativeInt(value, 1, 3)
		if err != nil {
			return nil, fmt.Errorf("style matrix index outside bounded first three entries")
		}
		list, err := nativeSingleton(matrix, dialect.drawing, item.list, true)
		if err != nil {
			return nil, err
		}
		if requireOnlyNativeAttrs(list) != nil || !onlyNativeXMLSpace(list.Text) || index > int64(len(list.Children)) {
			return nil, fmt.Errorf("style matrix entry unavailable")
		}
		selected := list.Children[index-1]
		if selected.Name.Space != dialect.drawing || item.ref == "fillRef" && (selected.Name.Local != "solidFill" && selected.Name.Local != "noFill") || item.ref == "lnRef" && selected.Name.Local != "ln" {
			return nil, fmt.Errorf("only exact solid fill and line matrix entries are supported")
		}
		paint, err := nativeStylePlaceholderColor(selected, dialect, colors[item.ref])
		if err != nil {
			return nil, err
		}
		// Validate even an overridden matrix entry before local precedence can
		// conceal malformed or duplicate source markup.
		paintGaps := nativeShapeGapSet{}
		wrapper := &nativeXMLNode{Children: []*nativeXMLNode{paint}}
		if item.ref == "fillRef" {
			validateNativeAutoShapeFill(wrapper, dialect, theme, &paintGaps)
		} else {
			if _, err := validateNativeAutoShapeLine(wrapper, dialect, theme, false, &paintGaps); err != nil {
				return nil, err
			}
		}
		if len(paintGaps.values) != 0 {
			return nil, fmt.Errorf("unmodeled selected style matrix paint")
		}
		found := false
		for _, name := range item.names {
			child, err := nativeSingleton(properties, dialect.drawing, name, false)
			if err != nil {
				return nil, err
			}
			found = found || child != nil
		}
		if !found {
			result.Children = append(result.Children, paint)
		}
	}
	return &result, nil
}

func nativeStylePlaceholderColor(node *nativeXMLNode, dialect nativeExtractDialect, color string) (*nativeXMLNode, error) {
	if node.Name == (xml.Name{Space: dialect.drawing, Local: "schemeClr"}) {
		value, _ := exactNativeAttr(node, "", "val")
		if value == "phClr" {
			if requireOnlyNativeAttrs(node, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(node) != nil {
				return nil, fmt.Errorf("theme placeholder color transforms remain unsupported")
			}
			return &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "srgbClr"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: color}}}, nil
		}
	}
	result := *node
	result.Attrs = append([]xml.Attr(nil), node.Attrs...)
	result.Children = make([]*nativeXMLNode, 0, len(node.Children))
	for _, child := range node.Children {
		copy, err := nativeStylePlaceholderColor(child, dialect, color)
		if err != nil {
			return nil, err
		}
		result.Children = append(result.Children, copy)
	}
	return &result, nil
}
