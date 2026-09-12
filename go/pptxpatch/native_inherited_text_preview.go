package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

const nativeInheritedTextPreviewCode = "pptx.source-inherited-text-approximate"

// This is a declared preview policy, NOT qualified Office cascade semantics.
// All nodes are owned projections; the package and mutation anchors stay raw.
func (e *nativeExtractor) inheritedTextPreview(body, style *nativeXMLNode, shape bool, d nativeExtractDialect) (*nativeXMLNode, error) {
	if body == nil {
		return nil, unsupportedNativeTextContent("missing inherited preview text body")
	}
	for _, p := range body.Children {
		for _, r := range p.Children {
			if r.Name == (xml.Name{Space: d.drawing, Local: "r"}) {
				t, err := nativeSingleton(r, d.drawing, "t", true)
				if err != nil {
					return nil, err
				}
				for _, cp := range t.Text {
					if cp < 32 || cp > 126 {
						return nil, unsupportedNativeTextContent("inherited preview requires graphic ASCII Latin text")
					}
				}
			}
		}
	}
	layers := []*nativeXMLNode{e.presentationTextPreviewStyle}
	if shape && e.slideDependencies.masterRoot != nil {
		styles, err := nativeSingleton(e.slideDependencies.masterRoot, d.presentation, "txStyles", false)
		if err != nil {
			return nil, err
		}
		if styles != nil {
			if requireOnlyNativeAttrs(styles) != nil || requireOnlyNativeChildren(styles, xml.Name{Space: d.presentation, Local: "titleStyle"}, xml.Name{Space: d.presentation, Local: "bodyStyle"}, xml.Name{Space: d.presentation, Local: "otherStyle"}) != nil {
				return nil, unsupportedNativeTextContent("unmodeled master text styles")
			}
			other, err := nativeSingleton(styles, d.presentation, "otherStyle", false)
			if err != nil {
				return nil, err
			}
			layers = append(layers, other)
		}
	}
	local, err := nativeSingleton(body, d.drawing, "lstStyle", true)
	if err != nil {
		return nil, err
	}
	if shape && style != nil {
		ref, err := nativeSingleton(style, d.drawing, "fontRef", true)
		if err != nil {
			return nil, err
		}
		if requireOnlyNativeAttrs(ref, xml.Name{Local: "idx"}) != nil {
			return nil, unsupportedNativeTextContent("invalid inherited shape font reference")
		}
		idx, _ := exactNativeAttr(ref, "", "idx")
		token := "+mn-lt"
		if idx == "major" {
			token = "+mj-lt"
		} else if idx != "minor" {
			return nil, unsupportedNativeTextContent("unsupported inherited shape font reference")
		}
		family, err := e.theme.resolveTypeface(token)
		if err != nil {
			return nil, err
		}
		color, err := exactNativeSolidColor(&nativeXMLNode{Children: ref.Children, Text: ref.Text}, d, e.theme)
		if err != nil {
			return nil, err
		}
		latin := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "latin"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "typeface"}, Value: family}}}
		rgb := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "srgbClr"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: color}}}
		fill := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "solidFill"}, Children: []*nativeXMLNode{rgb}}
		run := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "defRPr"}, Children: []*nativeXMLNode{latin, fill}}
		paragraph := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "defPPr"}, Children: []*nativeXMLNode{run}}
		layers = append(layers, &nativeXMLNode{Children: []*nativeXMLNode{paragraph}})
	}
	layers = append(layers, local)
	styles := make([]map[string]*nativeXMLNode, 0, len(layers))
	for _, layer := range layers {
		levels := map[string]*nativeXMLNode{}
		if layer != nil {
			if requireOnlyNativeAttrs(layer) != nil || !onlyNativeXMLSpace(layer.Text) {
				return nil, unsupportedNativeTextContent("unmodeled inherited text style container")
			}
			for _, child := range layer.Children {
				n := child.Name.Local
				if child.Name.Space != d.drawing || (n != "defPPr" && (len(n) != 7 || n[:3] != "lvl" || n[3] < '1' || n[3] > '9' || n[4:] != "pPr")) {
					return nil, unsupportedNativeTextContent("unmodeled inherited style level")
				}
				if levels[n] != nil {
					return nil, fmt.Errorf("duplicate inherited preview level")
				}
				clean, err := sanitizeNativeInheritedPreviewProperties(child, d, true, e.theme)
				if err != nil {
					return nil, err
				}
				levels[n] = clean
			}
		}
		styles = append(styles, levels)
	}
	result := *body
	result.Children = append([]*nativeXMLNode(nil), body.Children...)
	for pi, p := range body.Children {
		if p == local {
			result.Children[pi] = &nativeXMLNode{Name: local.Name}
			continue
		}
		if p.Name != (xml.Name{Space: d.drawing, Local: "p"}) {
			continue
		}
		if _, err := nativeSingleton(p, d.drawing, "endParaRPr", false); err != nil {
			return nil, err
		}
		ppr, err := nativeSingleton(p, d.drawing, "pPr", false)
		if err != nil {
			return nil, err
		}
		level := int64(0)
		if ppr != nil {
			if v, ok := exactNativeAttr(ppr, "", "lvl"); ok {
				level, err = parseCanonicalNativeInt(v, 0, 8)
				if err != nil {
					return nil, err
				}
			}
		}
		var merged *nativeXMLNode
		for _, ls := range styles {
			base := mergeNativeStyleNodes(ls["defPPr"], ls[fmt.Sprintf("lvl%dpPr", level+1)], d)
			merged = mergeNativeStyleNodes(merged, base, d)
		}
		clean, err := sanitizeNativeInheritedPreviewProperties(ppr, d, true, e.theme)
		if err != nil {
			return nil, err
		}
		merged = mergeNativeStyleNodes(merged, clean, d)
		merged.Name = xml.Name{Space: d.drawing, Local: "pPr"}
		if _, ok := exactNativeAttr(merged, "", "algn"); !ok {
			merged.Attrs = append(merged.Attrs, xml.Attr{Name: xml.Name{Local: "algn"}, Value: "l"})
		}
		if nativeChild(merged, d.drawing, "buChar") == nil && nativeChild(merged, d.drawing, "buNone") == nil {
			merged.Children = append(merged.Children, &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "buNone"}})
		}
		pc := *p
		pc.Children = []*nativeXMLNode{merged}
		for _, r := range p.Children {
			if r == ppr {
				continue
			}
			if r.Name == (xml.Name{Space: d.drawing, Local: "endParaRPr"}) {
				if err := validateNativeInheritedPreviewEnd(r, d); err != nil {
					return nil, err
				}
				continue
			}
			if r.Name != (xml.Name{Space: d.drawing, Local: "r"}) {
				pc.Children = append(pc.Children, r)
				continue
			}
			rpr, err := nativeSingleton(r, d.drawing, "rPr", false)
			if err != nil {
				return nil, err
			}
			clean, err := sanitizeNativeInheritedPreviewProperties(rpr, d, false, e.theme)
			if err != nil {
				return nil, err
			}
			if clean == nil {
				clean = &nativeXMLNode{}
			}
			clean.Name = xml.Name{Space: d.drawing, Local: "rPr"}
			rc := *r
			rc.Children = []*nativeXMLNode{clean}
			for _, child := range r.Children {
				if child != rpr {
					rc.Children = append(rc.Children, child)
				}
			}
			pc.Children = append(pc.Children, &rc)
		}
		result.Children[pi] = &pc
	}
	// Resolve validated source layers before applying documented false b/i defaults.
	resolved, err := resolveNativeLocalTextStyles(&result, d, e.theme)
	if err != nil {
		return nil, err
	}
	for _, p := range resolved.Children {
		for _, r := range p.Children {
			if r.Name != (xml.Name{Space: d.drawing, Local: "r"}) {
				continue
			}
			props, err := nativeSingleton(r, d.drawing, "rPr", true)
			if err != nil {
				return nil, err
			}
			for _, name := range []string{"b", "i"} {
				if _, ok := exactNativeAttr(props, "", name); !ok {
					props.Attrs = append(props.Attrs, xml.Attr{Name: xml.Name{Local: name}, Value: "0"})
				}
			}
		}
	}
	return resolved, nil
}

func sanitizeNativeInheritedPreviewProperties(node *nativeXMLNode, d nativeExtractDialect, paragraph bool, theme nativeResolvedTheme) (*nativeXMLNode, error) {
	if node == nil {
		return nil, nil
	}
	if duplicateNativeAttrs(node.Attrs) {
		return nil, unsupportedNativeTextContent("duplicate inherited text attribute")
	}
	out := *node
	out.Attrs = nil
	out.Children = nil
	for _, name := range []string{"ea", "cs", "defRPr"} {
		if _, err := nativeSingleton(node, d.drawing, name, false); err != nil {
			return nil, err
		}
	}
	for _, a := range node.Attrs {
		if a.Name.Space != "" {
			return nil, unsupportedNativeTextContent("foreign inherited text attribute")
		}
		drop := false
		if paragraph {
			switch a.Name.Local {
			case "defTabSz":
				if _, err := parseCanonicalNativeInt(a.Value, 0, 51206400); err != nil {
					return nil, err
				}
				drop = true
			case "rtl", "eaLnBrk", "latinLnBrk", "hangingPunct":
				v, err := nativeBool(a.Value)
				if err != nil {
					return nil, err
				}
				if a.Name.Local == "rtl" && v {
					return nil, unsupportedNativeTextContent("RTL inherited preview unsupported")
				}
				drop = true
			}
		} else if a.Name.Local == "kern" {
			if _, err := parseCanonicalNativeInt(a.Value, 0, 400000); err != nil {
				return nil, err
			}
			drop = true
		}
		if !drop {
			out.Attrs = append(out.Attrs, a)
		}
	}
	for _, child := range node.Children {
		if child.Name.Space == d.drawing && paragraph && child.Name.Local == "defRPr" {
			clean, err := sanitizeNativeInheritedPreviewProperties(child, d, false, theme)
			if err != nil {
				return nil, err
			}
			out.Children = append(out.Children, clean)
			continue
		}
		if child.Name.Space == d.drawing && !paragraph && (child.Name.Local == "ea" || child.Name.Local == "cs") {
			if requireOnlyNativeAttrs(child, xml.Name{Local: "typeface"}) != nil || requireOnlyNativeChildren(child) != nil {
				return nil, unsupportedNativeTextContent("unmodeled inactive font slot")
			}
			face, ok := exactNativeAttr(child, "", "typeface")
			if !ok || strings.TrimSpace(face) == "" {
				return nil, unsupportedNativeTextContent("invalid inactive font slot")
			}
			continue
		}
		out.Children = append(out.Children, child)
	}
	// Exact existing validators still enforce all active properties and duplicates.
	if err := validateNativeTextStyleProperties(&out, d, paragraph, theme); err != nil {
		return nil, err
	}
	return &out, nil
}

func validateNativeInheritedPreviewEnd(node *nativeXMLNode, d nativeExtractDialect) error {
	if duplicateNativeAttrs(node.Attrs) {
		return unsupportedNativeTextContent("duplicate terminal metadata")
	}
	if requireOnlyNativeAttrs(node, xml.Name{Local: "lang"}, xml.Name{Local: "dirty"}, xml.Name{Local: "smtClean"}, xml.Name{Local: "err"}) != nil || requireOnlyNativeChildren(node) != nil {
		return unsupportedNativeTextContent("active terminal text formatting remains unsupported")
	}
	for _, a := range node.Attrs {
		if a.Name.Local == "lang" {
			if !validNativeLanguage(a.Value) {
				return unsupportedNativeTextContent("invalid terminal language")
			}
		} else if _, err := nativeBool(a.Value); err != nil {
			return err
		}
	}
	return nil
}

func nativeMarkInheritedTextPreview(element *NativeElement) {
	if element.Compatibility.Status == NativeCompatibilityStatusEditable {
		element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	}
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: nativeInheritedTextPreviewCode, Message: "Read-only source-latin-inheritance-approximate-v1: declared source style ordering, false bold/italic defaults, Latin-only font slots; authored kerning disabled and terminal language/checking metadata preserved without layout. Font metrics, wrapping and terminal metrics may differ from PowerPoint."})
}
