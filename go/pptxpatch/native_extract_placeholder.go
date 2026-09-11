package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

type nativePlaceholderIdentity struct {
	index int64
	kind  string
}

// Ancestor placeholder text is an authoring prompt, not slide content. Admit
// prompts without copying their text into the source slide. Paragraph defaults
// require an explicit level; run/end-mark formatting remains unqualified.
func validateNativePlaceholderPrompt(paragraph *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) error {
	if requireOnlyNativeAttrs(paragraph) != nil || requireOnlyNativeChildren(paragraph, xml.Name{Space: dialect.drawing, Local: "pPr"}, xml.Name{Space: dialect.drawing, Local: "r"}, xml.Name{Space: dialect.drawing, Local: "endParaRPr"}) != nil || !onlyNativeXMLSpace(paragraph.Text) {
		return fmt.Errorf("pptxpatch: unmodeled ancestor placeholder prompt")
	}
	for _, name := range []string{"pPr", "endParaRPr"} {
		property, err := nativeSingleton(paragraph, dialect.drawing, name, false)
		if err != nil {
			return err
		}
		if name == "pPr" && property != nil && (len(property.Attrs) != 0 || len(property.Children) != 0) {
			if _, explicit := exactNativeAttr(property, "", "lvl"); !explicit {
				return fmt.Errorf("pptxpatch: styled ancestor prompt requires explicit list level")
			}
			if err := validateNativeTextStyleProperties(property, dialect, true, theme); err != nil {
				return err
			}
			continue
		}
		if property != nil && (requireOnlyNativeAttrs(property) != nil || requireOnlyNativeChildren(property) != nil || !onlyNativeXMLSpace(property.Text)) {
			return fmt.Errorf("pptxpatch: ancestor placeholder paragraph styling remains outside inheritance subset")
		}
	}
	for _, run := range paragraph.Children {
		if run.Name.Local != "r" {
			continue
		}
		if requireOnlyNativeAttrs(run) != nil || requireOnlyNativeChildren(run, xml.Name{Space: dialect.drawing, Local: "rPr"}, xml.Name{Space: dialect.drawing, Local: "t"}) != nil || !onlyNativeXMLSpace(run.Text) {
			return fmt.Errorf("pptxpatch: unmodeled ancestor placeholder run")
		}
		properties, err := nativeSingleton(run, dialect.drawing, "rPr", false)
		if err != nil {
			return err
		}
		if properties != nil && (requireOnlyNativeAttrs(properties) != nil || requireOnlyNativeChildren(properties) != nil || !onlyNativeXMLSpace(properties.Text)) {
			return fmt.Errorf("pptxpatch: ancestor placeholder run styling remains outside inheritance subset")
		}
		text, err := nativeSingleton(run, dialect.drawing, "t", true)
		if err != nil {
			return err
		}
		if requireOnlyNativeAttrs(text) != nil || len(text.Children) != 0 {
			return fmt.Errorf("pptxpatch: unmodeled ancestor placeholder prompt text")
		}
	}
	return nil
}

// Competing defaults are not guessed: only equal or disjoint qualified
// properties may be projected from an ancestor prompt's explicit list level.
func nativePromptStylesCompete(base, prompt *nativeXMLNode) bool {
	if base == nil || prompt == nil {
		return false
	}
	for _, a := range base.Attrs {
		for _, b := range prompt.Attrs {
			if a.Name == b.Name && a.Value != b.Value {
				return true
			}
		}
	}
	for _, a := range base.Children {
		for _, b := range prompt.Children {
			if base.Name.Local == "solidFill" && a.Name != b.Name {
				return true
			}
			if (a.Name.Local == "buNone" || a.Name.Local == "buChar") && (b.Name.Local == "buNone" || b.Name.Local == "buChar") && a.Name != b.Name {
				return true
			}
			if a.Name == b.Name && nativePromptStylesCompete(a, b) {
				return true
			}
		}
	}
	return !onlyNativeXMLSpace(base.Text) && base.Text != prompt.Text
}

func mergeNativePlaceholderPromptStyles(list, body *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*nativeXMLNode, error) {
	seen := map[int64]bool{}
	for _, paragraph := range body.Children {
		if paragraph.Name != (xml.Name{Space: dialect.drawing, Local: "p"}) {
			continue
		}
		property, err := nativeSingleton(paragraph, dialect.drawing, "pPr", false)
		if err != nil {
			return nil, err
		}
		if property == nil || (len(property.Attrs) == 0 && len(property.Children) == 0) {
			continue
		}
		value, _ := exactNativeAttr(property, "", "lvl")
		level, err := parseCanonicalNativeInt(value, 0, 8)
		if err != nil {
			return nil, err
		}
		if seen[level] {
			return nil, fmt.Errorf("pptxpatch: duplicate ancestor prompt list level")
		}
		seen[level] = true
		projected := *property
		projected.Name = xml.Name{Space: dialect.drawing, Local: fmt.Sprintf("lvl%dpPr", level+1)}
		projected.Attrs = nil
		for _, attr := range property.Attrs {
			if attr.Name != (xml.Name{Local: "lvl"}) {
				projected.Attrs = append(projected.Attrs, attr)
			}
		}
		if list != nil {
			for _, existing := range list.Children {
				if (existing.Name.Local == "defPPr" || existing.Name == projected.Name) && nativePromptStylesCompete(existing, &projected) {
					return nil, fmt.Errorf("pptxpatch: competing ancestor prompt defaults require explicit qualification")
				}
			}
		}
		list, err = mergeNativeListStyles(list, &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "lstStyle"}, Children: []*nativeXMLNode{&projected}}, dialect, theme)
		if err != nil {
			return nil, err
		}
	}
	return list, nil
}

func nativeTextPlaceholder(node *nativeXMLNode, dialect nativeExtractDialect) (*nativePlaceholderIdentity, error) {
	identity, err := nativePlaceholderMetadata(node, dialect)
	if err != nil || identity == nil {
		return identity, err
	}
	if identity.kind != "" && identity.kind != "title" && identity.kind != "body" {
		return nil, fmt.Errorf("pptxpatch: only title/body placeholder inheritance is qualified")
	}
	return identity, nil
}

func nativePlaceholderMetadata(node *nativeXMLNode, dialect nativeExtractDialect) (*nativePlaceholderIdentity, error) {
	nv, err := nativeSingleton(node, dialect.presentation, "nvSpPr", true)
	if err != nil {
		return nil, err
	}
	properties, err := nativeSingleton(nv, dialect.presentation, "nvPr", true)
	if err != nil {
		return nil, err
	}
	ph, err := nativeSingleton(properties, dialect.presentation, "ph", false)
	if err != nil || ph == nil {
		return nil, err
	}
	if requireOnlyNativeAttrs(properties) != nil || requireOnlyNativeChildren(properties, xml.Name{Space: dialect.presentation, Local: "ph"}) != nil || requireOnlyNativeAttrs(ph, xml.Name{Local: "type"}, xml.Name{Local: "idx"}) != nil || requireOnlyNativeChildren(ph) != nil {
		return nil, fmt.Errorf("pptxpatch: placeholder metadata is outside the exact subset")
	}
	identity := &nativePlaceholderIdentity{}
	identity.kind, _ = exactNativeAttr(ph, "", "type")
	if value, ok := exactNativeAttr(ph, "", "idx"); ok {
		identity.index, err = parseCanonicalNativeInt(value, 0, 4294967295)
		if err != nil {
			return nil, err
		}
	}
	return identity, nil
}

func nativeMatchingPlaceholder(root *nativeXMLNode, wanted nativePlaceholderIdentity, matchIndex bool, dialect nativeExtractDialect) (*nativeXMLNode, *nativePlaceholderIdentity, error) {
	if root == nil {
		return nil, nil, fmt.Errorf("pptxpatch: missing relationship-bound placeholder source")
	}
	common, err := nativeSingleton(root, dialect.presentation, "cSld", true)
	if err != nil {
		return nil, nil, err
	}
	tree, err := nativeSingleton(common, dialect.presentation, "spTree", true)
	if err != nil {
		return nil, nil, err
	}
	var matched *nativeXMLNode
	var identity *nativePlaceholderIdentity
	for _, shape := range tree.Children {
		if shape.Name != (xml.Name{Space: dialect.presentation, Local: "sp"}) {
			continue
		}
		candidate, err := nativePlaceholderMetadata(shape, dialect)
		if err != nil {
			return nil, nil, err
		}
		if candidate == nil {
			continue
		}
		if (matchIndex && candidate.index != wanted.index) || (!matchIndex && candidate.kind != wanted.kind) {
			continue
		}
		if matched != nil {
			return nil, nil, fmt.Errorf("pptxpatch: ambiguous placeholder inheritance match")
		}
		matched, identity = shape, candidate
	}
	if matched == nil {
		return nil, nil, fmt.Errorf("pptxpatch: placeholder has no exact layout/master match")
	}
	return matched, identity, nil
}

func mergeNativeListStyles(base, override *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*nativeXMLNode, error) {
	result := &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "lstStyle"}}
	for _, source := range []*nativeXMLNode{base, override} {
		if source == nil {
			continue
		}
		if requireOnlyNativeAttrs(source) != nil || !onlyNativeXMLSpace(source.Text) {
			return nil, unsupportedNativeTextContent("unmodeled inherited list style")
		}
		seen := map[xml.Name]bool{}
		for _, level := range source.Children {
			if seen[level.Name] {
				return nil, fmt.Errorf("pptxpatch: duplicate inherited list level")
			}
			seen[level.Name] = true
			if err := validateNativeTextStyleProperties(level, dialect, true, theme); err != nil {
				return nil, err
			}
			found := false
			for i, previous := range result.Children {
				if previous.Name == level.Name {
					merged := mergeNativeStyleNodes(previous, level, dialect)
					merged.Name = level.Name
					result.Children[i] = merged
					found = true
					break
				}
			}
			if !found {
				result.Children = append(result.Children, level)
			}
		}
	}
	return result, nil
}

// Resolve only an unambiguous title/body chain. The relationship graph has
// already validated namespaces, part ownership and unique internal routes.
// The result retains original raw offsets; inherited nodes are a read-only view.
func (extractor *nativeExtractor) resolveNativePlaceholder(node *nativeXMLNode, dialect nativeExtractDialect) (*nativeXMLNode, *NativePlaceholderType, error) {
	identity, err := nativeTextPlaceholder(node, dialect)
	if err != nil || identity == nil {
		return node, nil, err
	}
	layout, layoutIdentity, err := nativeMatchingPlaceholder(extractor.slideDependencies.layoutRoot, *identity, true, dialect)
	if err != nil {
		return nil, nil, err
	}
	if (layoutIdentity.kind != "title" && layoutIdentity.kind != "body") || (identity.kind != "" && identity.kind != layoutIdentity.kind) {
		return nil, nil, fmt.Errorf("pptxpatch: placeholder type is missing or conflicts with layout")
	}
	master, _, err := nativeMatchingPlaceholder(extractor.slideDependencies.masterRoot, *layoutIdentity, false, dialect)
	if err != nil {
		return nil, nil, err
	}
	var transform *nativeXMLNode
	var bodyProperties *nativeXMLNode
	var list *nativeXMLNode
	styles, err := nativeSingleton(extractor.slideDependencies.masterRoot, dialect.presentation, "txStyles", false)
	if err != nil {
		return nil, nil, err
	}
	if styles != nil {
		styleName := "bodyStyle"
		if layoutIdentity.kind == "title" {
			styleName = "titleStyle"
		}
		style, err := nativeSingleton(styles, dialect.presentation, styleName, false)
		if err != nil {
			return nil, nil, err
		}
		list, err = mergeNativeListStyles(nil, style, dialect, extractor.theme)
		if err != nil {
			return nil, nil, err
		}
	}
	for _, shape := range []*nativeXMLNode{master, layout, node} {
		if shape != node {
			nv, err := nativeSingleton(shape, dialect.presentation, "nvSpPr", true)
			if err != nil {
				return nil, nil, err
			}
			if requireOnlyNativeAttrs(nv) != nil || requireOnlyNativeChildren(nv, xml.Name{Space: dialect.presentation, Local: "cNvPr"}, xml.Name{Space: dialect.presentation, Local: "cNvSpPr"}, xml.Name{Space: dialect.presentation, Local: "nvPr"}) != nil {
				return nil, nil, fmt.Errorf("pptxpatch: unmodeled inherited nonvisual metadata")
			}
			for _, name := range []string{"cNvPr", "cNvSpPr"} {
				metadata, err := nativeSingleton(nv, dialect.presentation, name, true)
				if err != nil {
					return nil, nil, err
				}
				attrs := []xml.Name{{Local: "id"}, {Local: "name"}}
				if name == "cNvSpPr" {
					attrs = []xml.Name{{Local: "txBox"}}
				}
				if requireOnlyNativeAttrs(metadata, attrs...) != nil || requireOnlyNativeChildren(metadata) != nil {
					return nil, nil, fmt.Errorf("pptxpatch: unmodeled inherited visibility/nonvisual metadata")
				}
			}
		}
		if requireOnlyNativeAttrs(shape) != nil || requireOnlyNativeChildren(shape, xml.Name{Space: dialect.presentation, Local: "nvSpPr"}, xml.Name{Space: dialect.presentation, Local: "spPr"}, xml.Name{Space: dialect.presentation, Local: "txBody"}) != nil {
			return nil, nil, fmt.Errorf("pptxpatch: unmodeled inherited placeholder markup")
		}
		properties, err := nativeSingleton(shape, dialect.presentation, "spPr", true)
		if err != nil {
			return nil, nil, err
		}
		if requireOnlyNativeAttrs(properties) != nil || requireOnlyNativeChildren(properties, xml.Name{Space: dialect.drawing, Local: "xfrm"}) != nil {
			return nil, nil, fmt.Errorf("pptxpatch: inherited placeholder shape paint is outside exact text subset")
		}
		localTransform, err := nativeSingleton(properties, dialect.drawing, "xfrm", false)
		if err != nil {
			return nil, nil, err
		}
		if localTransform != nil {
			gaps := nativeShapeGapSet{}
			if _, err := validateNativeAutoShapeTransform(localTransform, dialect, &gaps); err != nil {
				return nil, nil, err
			}
			if gaps.refused() || requireOnlyNativeAttrs(localTransform) != nil {
				return nil, nil, fmt.Errorf("pptxpatch: inherited placeholder transform is outside exact subset")
			}
			transform = localTransform
		}
		body, err := nativeSingleton(shape, dialect.presentation, "txBody", false)
		if err != nil {
			return nil, nil, err
		}
		if body == nil {
			continue
		}
		if requireOnlyNativeAttrs(body) != nil || requireOnlyNativeChildren(body, xml.Name{Space: dialect.drawing, Local: "bodyPr"}, xml.Name{Space: dialect.drawing, Local: "lstStyle"}, xml.Name{Space: dialect.drawing, Local: "p"}) != nil {
			return nil, nil, fmt.Errorf("pptxpatch: unmodeled inherited text body")
		}
		if shape != node {
			for _, paragraph := range body.Children {
				if paragraph.Name == (xml.Name{Space: dialect.drawing, Local: "p"}) {
					if err := validateNativePlaceholderPrompt(paragraph, dialect, extractor.theme); err != nil {
						return nil, nil, err
					}
				}
			}
		}
		localBody, err := nativeSingleton(body, dialect.drawing, "bodyPr", false)
		if err != nil {
			return nil, nil, err
		}
		if localBody != nil {
			if _, err := extractNativeTextBodyLayout(body, dialect); err != nil {
				return nil, nil, err
			}
			bodyProperties = mergeNativeStyleNodes(bodyProperties, localBody, dialect)
			bodyProperties.Name = xml.Name{Space: dialect.drawing, Local: "bodyPr"}
		}
		localList, err := nativeSingleton(body, dialect.drawing, "lstStyle", false)
		if err != nil {
			return nil, nil, err
		}
		list, err = mergeNativeListStyles(list, localList, dialect, extractor.theme)
		if err != nil {
			return nil, nil, err
		}
		if shape != node {
			list, err = mergeNativePlaceholderPromptStyles(list, body, dialect, extractor.theme)
			if err != nil {
				return nil, nil, err
			}
		}
	}
	if transform == nil || bodyProperties == nil {
		return nil, nil, fmt.Errorf("pptxpatch: placeholder geometry/body metadata remains unresolved")
	}
	localBody, err := nativeSingleton(node, dialect.presentation, "txBody", true)
	if err != nil {
		return nil, nil, err
	}
	resolvedBody := *localBody
	resolvedBody.Children = []*nativeXMLNode{bodyProperties, list}
	for _, child := range localBody.Children {
		if child.Name != (xml.Name{Space: dialect.drawing, Local: "bodyPr"}) && child.Name != (xml.Name{Space: dialect.drawing, Local: "lstStyle"}) {
			resolvedBody.Children = append(resolvedBody.Children, child)
		}
	}
	result := *node
	result.Children = make([]*nativeXMLNode, 0, len(node.Children))
	for _, child := range node.Children {
		switch child.Name {
		case xml.Name{Space: dialect.presentation, Local: "spPr"}:
			projected := *child
			projected.Children = []*nativeXMLNode{transform}
			result.Children = append(result.Children, &projected)
		case xml.Name{Space: dialect.presentation, Local: "txBody"}:
			result.Children = append(result.Children, &resolvedBody)
		case xml.Name{Space: dialect.presentation, Local: "nvSpPr"}:
			projected := *child
			projected.Children = make([]*nativeXMLNode, 0, len(child.Children))
			for _, properties := range child.Children {
				if properties.Name == (xml.Name{Space: dialect.presentation, Local: "nvPr"}) {
					projected.Children = append(projected.Children, &nativeXMLNode{Name: properties.Name})
					continue
				}
				if properties.Name == (xml.Name{Space: dialect.presentation, Local: "cNvSpPr"}) {
					copy := *properties
					copy.Attrs = append([]xml.Attr(nil), properties.Attrs...)
					found := false
					for i := range copy.Attrs {
						if copy.Attrs[i].Name == (xml.Name{Local: "txBox"}) {
							copy.Attrs[i].Value = "1"
							found = true
						}
					}
					if !found {
						copy.Attrs = append(copy.Attrs, xml.Attr{Name: xml.Name{Local: "txBox"}, Value: "1"})
					}
					projected.Children = append(projected.Children, &copy)
					continue
				}
				projected.Children = append(projected.Children, properties)
			}
			result.Children = append(result.Children, &projected)
		default:
			result.Children = append(result.Children, child)
		}
	}
	kind := NativePlaceholderType(layoutIdentity.kind)
	return &result, &kind, nil
}
