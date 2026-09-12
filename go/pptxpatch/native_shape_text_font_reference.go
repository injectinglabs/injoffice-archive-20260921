package pptxpatch

import "encoding/xml"

// Absence must be proven from the actual source, not inferred from an omitted
// native field. Foreign lookalikes and duplicate containers are not absence.
// This first slice accepts only absent/empty external text style layers.
func nativeShapeReferenceEmptyLayer(root *nativeXMLNode, namespace, name string, master bool, dialect nativeExtractDialect) bool {
	if root == nil {
		return true
	}
	var layer *nativeXMLNode
	for _, child := range root.Children {
		if child.Name.Local != name {
			continue
		}
		if layer != nil || child.Name.Space != namespace {
			return false
		}
		layer = child
	}
	if layer == nil {
		return true
	}
	if !master {
		return requireEmptyNativeElement(layer) == nil
	}
	if requireOnlyNativeAttrs(layer) != nil || requireOnlyNativeChildren(layer, xml.Name{Space: dialect.presentation, Local: "titleStyle"}, xml.Name{Space: dialect.presentation, Local: "bodyStyle"}, xml.Name{Space: dialect.presentation, Local: "otherStyle"}) != nil {
		return false
	}
	for _, local := range []string{"titleStyle", "bodyStyle", "otherStyle"} {
		child, err := nativeSingleton(layer, dialect.presentation, local, false)
		if err != nil || child != nil && requireEmptyNativeElement(child) != nil {
			return false
		}
	}
	return true
}

// Resolve only source-authored shape font/color into an owned text projection.
// Local/list properties retain precedence; no size, weight or layout defaults
// are supplied. Source XML and mutation anchors remain untouched.
func resolveNativeShapeTextFontReference(body, style *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*nativeXMLNode, bool, error) {
	if style == nil || body == nil {
		return body, false, nil
	}
	ref, err := nativeSingleton(style, dialect.drawing, "fontRef", true)
	if err != nil {
		return nil, false, err
	}
	if err = requireOnlyNativeAttrs(ref, xml.Name{Local: "idx"}); err != nil {
		return nil, false, err
	}
	idx, _ := exactNativeAttr(ref, "", "idx")
	token := ""
	switch idx {
	case "minor":
		token = "+mn-lt"
	case "major":
		token = "+mj-lt"
	default:
		return nil, false, unsupportedNativeTextContent("unsupported shape font reference")
	}
	color, err := exactNativeSolidColor(&nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "solidFill"}, Children: ref.Children, Text: ref.Text}, dialect, theme)
	if err != nil {
		return nil, false, err
	}
	projected, err := resolveNativeLocalTextStyles(body, dialect, theme)
	if err != nil {
		return nil, false, err
	}
	used := false
	result := *projected
	result.Children = append([]*nativeXMLNode(nil), projected.Children...)
	for pi, paragraph := range projected.Children {
		if paragraph.Name != (xml.Name{Space: dialect.drawing, Local: "p"}) {
			continue
		}
		pcopy := *paragraph
		pcopy.Children = append([]*nativeXMLNode(nil), paragraph.Children...)
		for ri, run := range paragraph.Children {
			if run.Name != (xml.Name{Space: dialect.drawing, Local: "r"}) {
				continue
			}
			props, e := nativeSingleton(run, dialect.drawing, "rPr", true)
			if e != nil {
				return nil, false, e
			}
			latin, e := nativeSingleton(props, dialect.drawing, "latin", false)
			if e != nil {
				return nil, false, e
			}
			fill, e := nativeSingleton(props, dialect.drawing, "solidFill", false)
			if e != nil {
				return nil, false, e
			}
			if latin != nil && fill != nil {
				continue
			}
			text, e := nativeSingleton(run, dialect.drawing, "t", true)
			if e != nil {
				return nil, false, e
			}
			// A Latin theme slot is not a font substitution policy for other scripts.
			for _, cp := range text.Text {
				if cp < 0x20 || cp > 0x7e {
					return nil, false, unsupportedNativeTextContent("shape font reference requires qualified ASCII Latin text")
				}
			}
			owned := *props
			owned.Children = append([]*nativeXMLNode(nil), props.Children...)
			if latin == nil {
				family, e := theme.resolveTypeface(token)
				if e != nil {
					return nil, false, e
				}
				owned.Children = append(owned.Children, &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "latin"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "typeface"}, Value: family}}})
			}
			if fill == nil {
				owned.Children = append(owned.Children, &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "solidFill"}, Children: []*nativeXMLNode{{Name: xml.Name{Space: dialect.drawing, Local: "srgbClr"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: color}}}}})
			}
			rcopy := *run
			rcopy.Children = append([]*nativeXMLNode(nil), run.Children...)
			for ci, child := range rcopy.Children {
				if child == props {
					rcopy.Children[ci] = &owned
				}
			}
			pcopy.Children[ri] = &rcopy
			used = true
		}
		result.Children[pi] = &pcopy
	}
	return &result, used, nil
}
