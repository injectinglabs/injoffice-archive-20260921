package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
)

const (
	nativeInheritedTextPreviewCode   = "pptx.source-inherited-text-approximate"
	nativeInheritedTextOmissionsCode = "pptx.inherited-text-properties-omitted"
	nativeMaxInheritedPreviewBreaks  = 1000
	nativeMaxInheritedOmissionNames  = 64
	nativeXMLSpaceNamespace          = "http://www.w3.org/XML/1998/namespace"
)

// nativeInheritedTextOmissions records source properties that the read-only
// inherited preview validated but could not carry into native v1 paint. The
// vocabulary is fixed by the sanitizer, so the set stays small and bounded.
type nativeInheritedTextOmissions struct {
	seen map[string]bool
}

func (o *nativeInheritedTextOmissions) add(name string) {
	if o == nil {
		return
	}
	if o.seen == nil {
		o.seen = map[string]bool{}
	}
	if len(o.seen) >= nativeMaxInheritedOmissionNames && !o.seen[name] {
		o.seen["…"] = true
		return
	}
	o.seen[name] = true
}

func (o *nativeInheritedTextOmissions) names() []string {
	if o == nil || len(o.seen) == 0 {
		return nil
	}
	result := make([]string, 0, len(o.seen))
	for name := range o.seen {
		result = append(result, name)
	}
	sort.Strings(result)
	return result
}

// This is a declared preview policy, NOT qualified Office cascade semantics.
// All nodes are owned projections; the package and mutation anchors stay raw.
func (e *nativeExtractor) inheritedTextPreview(body, style *nativeXMLNode, shape bool, d nativeExtractDialect) (*nativeXMLNode, *nativeInheritedTextOmissions, []nativeParagraphSpacingSource, error) {
	if body == nil {
		return nil, nil, nil, unsupportedNativeTextContent("missing inherited preview text body")
	}
	layers := []*nativeXMLNode{e.presentationTextPreviewStyle}
	if shape && e.slideDependencies.masterRoot != nil {
		styles, err := nativeSingleton(e.slideDependencies.masterRoot, d.presentation, "txStyles", false)
		if err != nil {
			return nil, nil, nil, err
		}
		if styles != nil {
			if requireOnlyNativeAttrs(styles) != nil || requireOnlyNativeChildren(styles, xml.Name{Space: d.presentation, Local: "titleStyle"}, xml.Name{Space: d.presentation, Local: "bodyStyle"}, xml.Name{Space: d.presentation, Local: "otherStyle"}) != nil {
				return nil, nil, nil, unsupportedNativeTextContent("unmodeled master text styles")
			}
			other, err := nativeSingleton(styles, d.presentation, "otherStyle", false)
			if err != nil {
				return nil, nil, nil, err
			}
			layers = append(layers, other)
		}
	}
	if shape && style != nil {
		ref, err := nativeSingleton(style, d.drawing, "fontRef", true)
		if err != nil {
			return nil, nil, nil, err
		}
		if requireOnlyNativeAttrs(ref, xml.Name{Local: "idx"}) != nil {
			return nil, nil, nil, unsupportedNativeTextContent("invalid inherited shape font reference")
		}
		idx, _ := exactNativeAttr(ref, "", "idx")
		token := "+mn-lt"
		if idx == "major" {
			token = "+mj-lt"
		} else if idx != "minor" {
			return nil, nil, nil, unsupportedNativeTextContent("unsupported inherited shape font reference")
		}
		family, err := e.theme.resolveTypeface(token)
		if err != nil {
			return nil, nil, nil, err
		}
		latin := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "latin"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "typeface"}, Value: family}}}
		run := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "defRPr"}, Children: []*nativeXMLNode{latin}}
		// A font reference without an authored color contributes only its
		// typeface; the color then comes from the remaining source layers.
		if len(ref.Children) != 0 || !onlyNativeXMLSpace(ref.Text) {
			color, err := exactNativeSolidColor(&nativeXMLNode{Children: ref.Children, Text: ref.Text}, d, e.theme)
			if err != nil {
				return nil, nil, nil, err
			}
			rgb := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "srgbClr"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: color}}}
			run.Children = append(run.Children, &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "solidFill"}, Children: []*nativeXMLNode{rgb}})
		}
		paragraph := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "defPPr"}, Children: []*nativeXMLNode{run}}
		layers = append(layers, &nativeXMLNode{Children: []*nativeXMLNode{paragraph}})
	}
	return e.inheritedTextPreviewLayers(body, layers, d)
}

// inheritedTextPreviewLayers applies the declared layer order (earliest layer
// loses) followed by the body's own list style and local properties. Layers
// may be nil. Validated-but-unmodeled properties are collected as omissions
// for disclosure; unknown or paint-active unsupported source still refuses.
func (e *nativeExtractor) inheritedTextPreviewLayers(body *nativeXMLNode, layers []*nativeXMLNode, d nativeExtractDialect) (*nativeXMLNode, *nativeInheritedTextOmissions, []nativeParagraphSpacingSource, error) {
	if body == nil {
		return nil, nil, nil, unsupportedNativeTextContent("missing inherited preview text body")
	}
	omit := &nativeInheritedTextOmissions{}
	spacing := []nativeParagraphSpacingSource{}
	for _, p := range body.Children {
		for _, r := range p.Children {
			if r.Name == (xml.Name{Space: d.drawing, Local: "r"}) {
				t, err := nativeSingleton(r, d.drawing, "t", true)
				if err != nil {
					return nil, nil, nil, err
				}
				for _, cp := range t.Text {
					if cp < 32 || cp > 126 {
						return nil, nil, nil, unsupportedNativeTextContent("inherited preview requires graphic ASCII Latin text")
					}
				}
			}
		}
	}
	local, err := nativeSingleton(body, d.drawing, "lstStyle", true)
	if err != nil {
		return nil, nil, nil, err
	}
	layers = append(append([]*nativeXMLNode(nil), layers...), local)
	styles := make([]map[string]*nativeXMLNode, 0, len(layers))
	for _, layer := range layers {
		levels := map[string]*nativeXMLNode{}
		if layer != nil {
			if requireOnlyNativeAttrs(layer) != nil || !onlyNativeXMLSpace(layer.Text) {
				return nil, nil, nil, unsupportedNativeTextContent("unmodeled inherited text style container")
			}
			for _, child := range layer.Children {
				n := child.Name.Local
				if child.Name.Space != d.drawing || (n != "defPPr" && (len(n) != 7 || n[:3] != "lvl" || n[3] < '1' || n[3] > '9' || n[4:] != "pPr")) {
					return nil, nil, nil, unsupportedNativeTextContent("unmodeled inherited style level")
				}
				if levels[n] != nil {
					return nil, nil, nil, fmt.Errorf("duplicate inherited preview level")
				}
				clean, err := sanitizeNativeInheritedPreviewProperties(child, d, true, e.theme, omit)
				if err != nil {
					return nil, nil, nil, err
				}
				levels[n] = clean
			}
		}
		styles = append(styles, levels)
	}
	result := *body
	result.Children = make([]*nativeXMLNode, 0, len(body.Children))
	breaks := 0
	for _, p := range body.Children {
		if p == local {
			result.Children = append(result.Children, &nativeXMLNode{Name: local.Name})
			continue
		}
		if p.Name != (xml.Name{Space: d.drawing, Local: "p"}) {
			result.Children = append(result.Children, p)
			continue
		}
		projected, projectedSpacing, err := e.inheritedPreviewParagraph(p, styles, d, omit, &breaks)
		if err != nil {
			return nil, nil, nil, err
		}
		result.Children = append(result.Children, projected...)
		spacing = append(spacing, projectedSpacing...)
	}
	// Resolve validated source layers before applying documented false b/i defaults.
	resolved, err := resolveNativeLocalTextStyles(&result, d, e.theme)
	if err != nil {
		return nil, nil, nil, err
	}
	for _, p := range resolved.Children {
		for _, r := range p.Children {
			if r.Name != (xml.Name{Space: d.drawing, Local: "r"}) {
				continue
			}
			props, err := nativeSingleton(r, d.drawing, "rPr", true)
			if err != nil {
				return nil, nil, nil, err
			}
			for _, name := range []string{"b", "i"} {
				if _, ok := exactNativeAttr(props, "", name); !ok {
					props.Attrs = append(props.Attrs, xml.Attr{Name: xml.Name{Local: name}, Value: "0"})
				}
			}
		}
	}
	return resolved, omit, spacing, nil
}

// inheritedPreviewParagraph projects one source paragraph. An authored a:br
// continues in a bullet-free paragraph at the same left margin, and a paragraph
// without runs becomes one blank space run carrying its end-mark metrics so
// the line still occupies height. Both projections are disclosed omissions.
func (e *nativeExtractor) inheritedPreviewParagraph(p *nativeXMLNode, styles []map[string]*nativeXMLNode, d nativeExtractDialect, omit *nativeInheritedTextOmissions, breaks *int) ([]*nativeXMLNode, []nativeParagraphSpacingSource, error) {
	if requireOnlyNativeAttrs(p) != nil || !onlyNativeXMLSpace(p.Text) {
		return nil, nil, unsupportedNativeTextContent("unmodeled inherited paragraph markup")
	}
	end, err := nativeSingleton(p, d.drawing, "endParaRPr", false)
	if err != nil {
		return nil, nil, err
	}
	ppr, err := nativeSingleton(p, d.drawing, "pPr", false)
	if err != nil {
		return nil, nil, err
	}
	level := int64(0)
	if ppr != nil {
		if v, ok := exactNativeAttr(ppr, "", "lvl"); ok {
			level, err = parseCanonicalNativeInt(v, 0, 8)
			if err != nil {
				return nil, nil, err
			}
		}
	}
	var merged *nativeXMLNode
	for _, ls := range styles {
		base := mergeNativeStyleNodes(ls["defPPr"], ls[fmt.Sprintf("lvl%dpPr", level+1)], d)
		merged = mergeNativeStyleNodes(merged, base, d)
	}
	clean, err := sanitizeNativeInheritedPreviewProperties(ppr, d, true, e.theme, omit)
	if err != nil {
		return nil, nil, err
	}
	merged = mergeNativeStyleNodes(merged, clean, d)
	merged = nativeWithoutBulletTextMarkers(merged, d)
	merged.Name = xml.Name{Space: d.drawing, Local: "pPr"}
	if _, ok := exactNativeAttr(merged, "", "algn"); !ok {
		merged.Attrs = append(merged.Attrs, xml.Attr{Name: xml.Name{Local: "algn"}, Value: "l"})
	}
	if nativeChild(merged, d.drawing, "buChar") == nil && nativeChild(merged, d.drawing, "buNone") == nil {
		merged.Children = append(merged.Children, &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "buNone"}})
	}
	if nativeChild(merged, d.drawing, "buNone") != nil && nativeChild(merged, d.drawing, "buFont") != nil {
		// A later layer removed the bullet; the earlier layer's bullet font no
		// longer applies to any marker (exact DrawingML semantics, not a guess).
		kept := make([]*nativeXMLNode, 0, len(merged.Children))
		for _, child := range merged.Children {
			if child.Name != (xml.Name{Space: d.drawing, Local: "buFont"}) {
				kept = append(kept, child)
			}
		}
		merged.Children = kept
	}
	// The merged a:pPr now holds the cascaded a:lnSpc/a:spcBef/a:spcAft. They
	// leave the projected paint XML here so the exact paragraph extractor still
	// sees the node shape it already accepts; the values travel beside it.
	spacingSource, err := nativeReadParagraphSpacing(merged, d)
	if err != nil {
		return nil, nil, err
	}
	merged = nativeWithoutParagraphSpacing(merged, d)
	merged.Name = xml.Name{Space: d.drawing, Local: "pPr"}
	var endProperties *nativeXMLNode
	if end != nil {
		endProperties, err = sanitizeNativeInheritedPreviewEnd(end, d, e.theme, omit)
		if err != nil {
			return nil, nil, err
		}
	}
	newParagraph := func(properties *nativeXMLNode) *nativeXMLNode {
		pc := *p
		pc.Children = []*nativeXMLNode{properties}
		return &pc
	}
	current := newParagraph(merged)
	runs := 0
	result := []*nativeXMLNode{current}
	for _, r := range p.Children {
		if r == ppr || r == end {
			continue
		}
		switch r.Name {
		case xml.Name{Space: d.drawing, Local: "br"}:
			if requireOnlyNativeAttrs(r) != nil || requireOnlyNativeChildren(r, xml.Name{Space: d.drawing, Local: "rPr"}) != nil || !onlyNativeXMLSpace(r.Text) {
				return nil, nil, unsupportedNativeTextContent("unmodeled inherited line break markup")
			}
			breakProperties, err := nativeSingleton(r, d.drawing, "rPr", false)
			if err != nil {
				return nil, nil, err
			}
			if _, err := sanitizeNativeInheritedPreviewProperties(breakProperties, d, false, e.theme, omit); err != nil {
				return nil, nil, err
			}
			*breaks++
			if *breaks > nativeMaxInheritedPreviewBreaks {
				return nil, nil, unsupportedNativeTextContent("inherited preview line-break budget exceeded")
			}
			if runs == 0 {
				current.Children = append(current.Children, nativeInheritedBlankRun(endProperties, d))
			}
			omit.add("a:br→continuation-paragraph")
			current = newParagraph(nativeInheritedContinuationProperties(merged, d))
			runs = 0
			result = append(result, current)
		case xml.Name{Space: d.drawing, Local: "r"}:
			rpr, err := nativeSingleton(r, d.drawing, "rPr", false)
			if err != nil {
				return nil, nil, err
			}
			clean, err := sanitizeNativeInheritedPreviewProperties(rpr, d, false, e.theme, omit)
			if err != nil {
				return nil, nil, err
			}
			if clean == nil {
				clean = &nativeXMLNode{}
			}
			clean.Name = xml.Name{Space: d.drawing, Local: "rPr"}
			rc := *r
			rc.Children = []*nativeXMLNode{clean}
			for _, child := range r.Children {
				if child != rpr {
					rc.Children = append(rc.Children, nativeInheritedPreservedText(child, d, omit))
				}
			}
			current.Children = append(current.Children, &rc)
			runs++
		default:
			current.Children = append(current.Children, r)
			runs++
		}
	}
	if runs == 0 {
		omit.add("empty-paragraph→blank-line")
		current.Children = append(current.Children, nativeInheritedBlankRun(endProperties, d))
	}
	// One source paragraph may project as several: space before belongs to the
	// first fragment and space after to the last, so a hard break never repeats
	// either gap inside what PowerPoint authored as one paragraph.
	spacing := make([]nativeParagraphSpacingSource, len(result))
	for index := range spacing {
		spacing[index] = nativeParagraphSpacingSource{
			present:                spacingSource.present,
			lineSpacingPercent1000: spacingSource.lineSpacingPercent1000,
			lineSpacingHundredthPt: spacingSource.lineSpacingHundredthPt,
		}
	}
	if len(spacing) > 0 {
		spacing[0].spaceBeforePercent1000 = spacingSource.spaceBeforePercent1000
		spacing[0].spaceBeforeHundredthPt = spacingSource.spaceBeforeHundredthPt
		last := len(spacing) - 1
		spacing[last].spaceAfterPercent1000 = spacingSource.spaceAfterPercent1000
		spacing[last].spaceAfterHundredthPt = spacingSource.spaceAfterHundredthPt
	}
	return result, spacing, nil
}

// nativeInheritedContinuationProperties derives the paragraph properties for
// text after an authored a:br: same level, alignment and left margin, no bullet
// and no hanging first-line indent.
func nativeInheritedContinuationProperties(properties *nativeXMLNode, d nativeExtractDialect) *nativeXMLNode {
	result := &nativeXMLNode{Name: properties.Name}
	for _, attr := range properties.Attrs {
		if attr.Name == (xml.Name{Local: "indent"}) {
			continue
		}
		result.Attrs = append(result.Attrs, attr)
	}
	result.Attrs = append(result.Attrs, xml.Attr{Name: xml.Name{Local: "indent"}, Value: "0"})
	for _, child := range properties.Children {
		if child.Name.Space == d.drawing && (child.Name.Local == "buChar" || child.Name.Local == "buNone" || child.Name.Local == "buFont") {
			continue
		}
		result.Children = append(result.Children, child)
	}
	result.Children = append(result.Children, &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "buNone"}})
	return result
}

// nativeInheritedPreservedText projects an a:t whose authored edge whitespace
// lacks xml:space="preserve". DrawingML text content is literal in PowerPoint,
// so the owned copy declares preservation and the assumption is disclosed.
func nativeInheritedPreservedText(node *nativeXMLNode, d nativeExtractDialect, omit *nativeInheritedTextOmissions) *nativeXMLNode {
	if node.Name != (xml.Name{Space: d.drawing, Local: "t"}) || !hasNativeEdgeXMLSpace(node.Text) {
		return node
	}
	if _, declared := exactNativeAttr(node, nativeXMLSpaceNamespace, "space"); declared {
		return node
	}
	copy := *node
	copy.Attrs = append(append([]xml.Attr(nil), node.Attrs...), xml.Attr{Name: xml.Name{Space: nativeXMLSpaceNamespace, Local: "space"}, Value: "preserve"})
	omit.add("a:t@xml:space=preserve-assumed")
	return &copy
}

// nativeInheritedBlankRun is the owned projection of a run-less line: one
// preserved space whose run properties are the sanitized end-mark metrics.
func nativeInheritedBlankRun(endProperties *nativeXMLNode, d nativeExtractDialect) *nativeXMLNode {
	properties := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "rPr"}}
	if endProperties != nil {
		properties.Attrs = append(properties.Attrs, endProperties.Attrs...)
		properties.Children = append(properties.Children, endProperties.Children...)
	}
	text := &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "t"}, Attrs: []xml.Attr{{Name: xml.Name{Space: nativeXMLSpaceNamespace, Local: "space"}, Value: "preserve"}}, Text: " "}
	return &nativeXMLNode{Name: xml.Name{Space: d.drawing, Local: "r"}, Children: []*nativeXMLNode{properties, text}}
}

// sanitizeNativeInheritedPreviewEnd validates end-mark properties like a run.
// Only the blank-line projection uses them; a paragraph with runs omits them.
func sanitizeNativeInheritedPreviewEnd(node *nativeXMLNode, d nativeExtractDialect, theme nativeResolvedTheme, omit *nativeInheritedTextOmissions) (*nativeXMLNode, error) {
	if duplicateNativeAttrs(node.Attrs) {
		return nil, unsupportedNativeTextContent("duplicate terminal metadata")
	}
	if requireOnlyNativeAttrs(node, xml.Name{Local: "lang"}, xml.Name{Local: "dirty"}, xml.Name{Local: "smtClean"}, xml.Name{Local: "err"}) != nil || requireOnlyNativeChildren(node) != nil {
		omit.add("a:endParaRPr")
	}
	clean, err := sanitizeNativeInheritedPreviewProperties(node, d, false, theme, omit)
	if err != nil {
		return nil, unsupportedNativeTextContent("active terminal text formatting remains unsupported: " + err.Error())
	}
	if clean == nil {
		clean = &nativeXMLNode{}
	}
	clean.Name = xml.Name{Space: d.drawing, Local: "rPr"}
	return clean, nil
}

// nativeInheritedSpacingValue validates an a:lnSpc/a:spcBef/a:spcAft container
// as one canonical percentage or point value. The value is never laid out.
func nativeInheritedSpacingValue(node *nativeXMLNode, d nativeExtractDialect) error {
	if requireOnlyNativeAttrs(node) != nil || !onlyNativeXMLSpace(node.Text) || requireOnlyNativeChildren(node, xml.Name{Space: d.drawing, Local: "spcPct"}, xml.Name{Space: d.drawing, Local: "spcPts"}) != nil || len(node.Children) != 1 {
		return unsupportedNativeTextContent("unmodeled paragraph spacing markup")
	}
	value := node.Children[0]
	if requireOnlyNativeAttrs(value, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(value) != nil || !onlyNativeXMLSpace(value.Text) {
		return unsupportedNativeTextContent("unmodeled paragraph spacing value")
	}
	raw, ok := exactNativeAttr(value, "", "val")
	if !ok {
		return unsupportedNativeTextContent("paragraph spacing value is missing")
	}
	maximum := int64(13200000)
	if value.Name.Local == "spcPts" {
		maximum = 158400
	}
	if _, err := parseCanonicalNativeInt(raw, 0, maximum); err != nil {
		return unsupportedNativeTextContent("paragraph spacing value is not canonical")
	}
	return nil
}

func nativeInheritedBulletSizeValue(node *nativeXMLNode, minimum, maximum int64) error {
	if requireOnlyNativeAttrs(node, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(node) != nil || !onlyNativeXMLSpace(node.Text) {
		return unsupportedNativeTextContent("unmodeled bullet size markup")
	}
	raw, ok := exactNativeAttr(node, "", "val")
	if !ok {
		return unsupportedNativeTextContent("bullet size value is missing")
	}
	if _, err := parseCanonicalNativeInt(raw, minimum, maximum); err != nil {
		return unsupportedNativeTextContent("bullet size value is not canonical")
	}
	return nil
}

func nativeInheritedTabList(node *nativeXMLNode, d nativeExtractDialect) error {
	if requireOnlyNativeAttrs(node) != nil || !onlyNativeXMLSpace(node.Text) || requireOnlyNativeChildren(node, xml.Name{Space: d.drawing, Local: "tab"}) != nil || len(node.Children) > 32 {
		return unsupportedNativeTextContent("unmodeled tab stop list")
	}
	for _, tab := range node.Children {
		if requireOnlyNativeAttrs(tab, xml.Name{Local: "pos"}, xml.Name{Local: "algn"}) != nil || requireOnlyNativeChildren(tab) != nil || !onlyNativeXMLSpace(tab.Text) {
			return unsupportedNativeTextContent("unmodeled tab stop")
		}
		if value, ok := exactNativeAttr(tab, "", "pos"); ok {
			if _, err := parseCanonicalNativeInt(value, -51206400, 51206400); err != nil {
				return unsupportedNativeTextContent("tab stop position is not canonical")
			}
		}
		if value, ok := exactNativeAttr(tab, "", "algn"); ok && value != "l" && value != "ctr" && value != "r" && value != "dec" {
			return unsupportedNativeTextContent("tab stop alignment is invalid")
		}
	}
	return nil
}

// sanitizeNativeInheritedPreviewProperties validates every source property and
// returns an owned copy containing only the properties native v1 paints.
// Non-layout metadata is dropped silently as before; properties that would
// affect Office layout but are outside the contract (paragraph spacing,
// bullet color/size, character spacing, tab stops, strike=noStrike, panose)
// are validated, dropped, and recorded in omit. Everything else still refuses.
func sanitizeNativeInheritedPreviewProperties(node *nativeXMLNode, d nativeExtractDialect, paragraph bool, theme nativeResolvedTheme, omit *nativeInheritedTextOmissions) (*nativeXMLNode, error) {
	if node == nil {
		return nil, nil
	}
	if duplicateNativeAttrs(node.Attrs) {
		return nil, unsupportedNativeTextContent("duplicate inherited text attribute")
	}
	out := *node
	out.Attrs = nil
	out.Children = nil
	for _, name := range []string{"ea", "cs", "defRPr", "lnSpc", "spcBef", "spcAft", "buClr", "buSzPct", "buSzPts", "buSzTx", "buFontTx", "buClrTx", "tabLst", "buFont"} {
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
			case "fontAlgn":
				switch a.Value {
				case "auto", "t", "ctr", "base", "b":
				default:
					return nil, unsupportedNativeTextContent("invalid font alignment")
				}
				omit.add("a:pPr@fontAlgn")
				drop = true
			}
		} else {
			switch a.Name.Local {
			case "kern":
				if _, err := parseCanonicalNativeInt(a.Value, 0, 400000); err != nil {
					return nil, err
				}
				drop = true
			case "spc":
				if _, err := parseCanonicalNativeInt(a.Value, -400000, 400000); err != nil {
					return nil, err
				}
				omit.add("a:rPr@spc")
				drop = true
			case "strike":
				if a.Value != "noStrike" {
					if a.Value != "sngStrike" && a.Value != "dblStrike" {
						return nil, unsupportedNativeTextContent("invalid strike value")
					}
					return nil, unsupportedNativeTextContent("strikethrough is outside the inherited preview subset")
				}
				omit.add("a:rPr@strike=noStrike")
				drop = true
			case "noProof":
				if _, err := nativeBool(a.Value); err != nil {
					return nil, err
				}
				drop = true
			case "altLang":
				if !validNativeLanguage(a.Value) {
					return nil, unsupportedNativeTextContent("invalid alternate language tag")
				}
				drop = true
			}
		}
		if !drop {
			out.Attrs = append(out.Attrs, a)
		}
	}
	for _, child := range node.Children {
		if child.Name.Space == d.drawing && paragraph {
			switch child.Name.Local {
			case "defRPr":
				clean, err := sanitizeNativeInheritedPreviewProperties(child, d, false, theme, omit)
				if err != nil {
					return nil, err
				}
				out.Children = append(out.Children, clean)
				continue
			case "lnSpc", "spcBef", "spcAft":
				// Kept in the owned copy so the cascade merge resolves the
				// authored spacing; nativeWithoutParagraphSpacing removes it
				// again once the projected paragraph has recorded the values.
				if err := nativeInheritedSpacingValue(child, d); err != nil {
					return nil, err
				}
				out.Children = append(out.Children, child)
				continue
			case "buClr":
				if _, err := exactNativeSolidColor(&nativeXMLNode{Children: child.Children, Text: child.Text, Attrs: child.Attrs}, d, theme); err != nil {
					return nil, unsupportedNativeTextContent("unmodeled bullet color: " + err.Error())
				}
				omit.add("a:buClr")
				continue
			case "buSzPct":
				if err := nativeInheritedBulletSizeValue(child, 25000, 400000); err != nil {
					return nil, err
				}
				omit.add("a:buSzPct")
				continue
			case "buSzPts":
				if err := nativeInheritedBulletSizeValue(child, 100, 400000); err != nil {
					return nil, err
				}
				omit.add("a:buSzPts")
				continue
			case "buSzTx", "buFontTx", "buClrTx":
				// "Follow text" markers cancel an inherited buFont/buSz*/buClr from
				// a lower layer during the merge, then leave the projection.
				if err := requireEmptyNativeElement(child); err != nil {
					return nil, unsupportedNativeTextContent("unmodeled bullet inheritance marker")
				}
				omit.add("a:" + child.Name.Local)
				out.Children = append(out.Children, child)
				continue
			case "tabLst":
				if err := nativeInheritedTabList(child, d); err != nil {
					return nil, err
				}
				omit.add("a:tabLst")
				continue
			case "buFont":
				if _, ok := exactNativeAttr(child, "", "panose"); ok {
					value, _ := exactNativeAttr(child, "", "panose")
					if len(value) > 20 || strings.IndexFunc(value, func(r rune) bool { return !(r >= '0' && r <= '9' || r >= 'A' && r <= 'F' || r >= 'a' && r <= 'f') }) >= 0 {
						return nil, unsupportedNativeTextContent("invalid buFont panose metadata")
					}
					copy := *child
					copy.Attrs = nil
					for _, attr := range child.Attrs {
						if attr.Name != (xml.Name{Local: "panose"}) {
							copy.Attrs = append(copy.Attrs, attr)
						}
					}
					omit.add("a:buFont@panose")
					out.Children = append(out.Children, &copy)
					continue
				}
			}
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
	if err := validateNativeTextStyleProperties(nativeWithoutParagraphSpacing(nativeWithoutBulletTextMarkers(&out, d), d), d, paragraph, theme); err != nil {
		return nil, err
	}
	return &out, nil
}

func nativeIsBulletTextMarker(node *nativeXMLNode, d nativeExtractDialect) bool {
	return node != nil && node.Name.Space == d.drawing && (node.Name.Local == "buFontTx" || node.Name.Local == "buSzTx" || node.Name.Local == "buClrTx")
}

// nativeWithoutBulletTextMarkers returns an owned copy without a:buFontTx,
// a:buSzTx and a:buClrTx. After the merge these markers only mean "no
// inherited bullet font/size/color", which the exact extractor expresses by
// absence.
func nativeWithoutBulletTextMarkers(node *nativeXMLNode, d nativeExtractDialect) *nativeXMLNode {
	if node == nil {
		return nil
	}
	copy := *node
	copy.Children = make([]*nativeXMLNode, 0, len(node.Children))
	for _, child := range node.Children {
		if !nativeIsBulletTextMarker(child, d) {
			copy.Children = append(copy.Children, child)
		}
	}
	return &copy
}

func nativeMarkInheritedTextPreview(element *NativeElement) {
	if element.Compatibility.Status == NativeCompatibilityStatusEditable {
		element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	}
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: nativeInheritedTextPreviewCode, Message: "Read-only source-latin-inheritance-approximate-v1: declared source style ordering, false bold/italic defaults, Latin-only font slots; authored kerning disabled and terminal language/checking metadata preserved without layout. Font metrics, wrapping and terminal metrics may differ from PowerPoint."})
}

// nativeMarkInheritedTextOmissions discloses every validated source property
// the inherited preview dropped from layout. Nothing is disclosed when the
// projection carried every authored property.
func nativeMarkInheritedTextOmissions(element *NativeElement, omit *nativeInheritedTextOmissions) {
	names := omit.names()
	if element == nil || len(names) == 0 {
		return
	}
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning,
		Code:     nativeInheritedTextOmissionsCode,
		Message:  "Read-only inherited text preview validated these source properties but omits them from native v1 layout: " + strings.Join(names, ", ") + ". Bullet color/size, character spacing, tab stops, terminal run metrics, hard-break continuation and blank-line projections are approximations, not PowerPoint layout.",
	})
}
