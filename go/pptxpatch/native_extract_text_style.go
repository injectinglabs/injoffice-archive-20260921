package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
	"unicode"
	"unicode/utf8"
)

// Materialize the bounded DrawingML style cascade into an owned projection.
// Raw source nodes/offsets remain untouched: writes still target the original
// element subtree, not this rendering-only property view.
func resolveNativeLocalTextStyles(body *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*nativeXMLNode, error) {
	if body == nil {
		return nil, nil
	}
	list, err := nativeSingleton(body, dialect.drawing, "lstStyle", true)
	if err != nil {
		return nil, err
	}
	if requireOnlyNativeAttrs(list) != nil || !onlyNativeXMLSpace(list.Text) {
		return nil, unsupportedNativeTextContent("list style has unmodeled attributes or text")
	}
	levels := map[string]*nativeXMLNode{}
	for _, child := range list.Children {
		if child.Name.Space != dialect.drawing || (child.Name.Local != "defPPr" && (len(child.Name.Local) != 7 || child.Name.Local[:3] != "lvl" || child.Name.Local[3] < '1' || child.Name.Local[3] > '9' || child.Name.Local[4:] != "pPr")) {
			return nil, unsupportedNativeTextContent("list style contains an unmodeled level")
		}
		if levels[child.Name.Local] != nil {
			return nil, fmt.Errorf("pptxpatch: duplicate list style level")
		}
		if err := validateNativeTextStyleProperties(child, dialect, true, theme); err != nil {
			return nil, err
		}
		levels[child.Name.Local] = child
	}
	result := *body
	result.Children = make([]*nativeXMLNode, 0, len(body.Children))
	for _, child := range body.Children {
		if child == list {
			result.Children = append(result.Children, &nativeXMLNode{Name: list.Name})
			continue
		}
		if child.Name != (xml.Name{Space: dialect.drawing, Local: "p"}) {
			result.Children = append(result.Children, child)
			continue
		}
		local, err := nativeSingleton(child, dialect.drawing, "pPr", false)
		if err != nil {
			return nil, err
		}
		if local != nil {
			if err := validateNativeTextStyleProperties(local, dialect, true, theme); err != nil {
				return nil, err
			}
		}
		level := int64(0)
		if local != nil {
			if value, ok := exactNativeAttr(local, "", "lvl"); ok {
				level, err = parseCanonicalNativeInt(value, 0, 8)
				if err != nil {
					return nil, err
				}
			}
		}
		properties := mergeNativeStyleNodes(levels["defPPr"], levels[fmt.Sprintf("lvl%dpPr", level+1)], dialect)
		properties = mergeNativeStyleNodes(properties, local, dialect)
		properties.Name = xml.Name{Space: dialect.drawing, Local: "pPr"}
		if _, ok := exactNativeAttr(properties, "", "lvl"); !ok {
			properties.Attrs = append(properties.Attrs, xml.Attr{Name: xml.Name{Local: "lvl"}, Value: strconv.FormatInt(level, 10)})
		}
		defaults, err := nativeSingleton(properties, dialect.drawing, "defRPr", false)
		if err != nil {
			return nil, err
		}
		filtered := make([]*nativeXMLNode, 0, len(properties.Children))
		for _, p := range properties.Children {
			if p != defaults {
				filtered = append(filtered, p)
			}
		}
		properties.Children = filtered
		paragraph := *child
		paragraph.Children = []*nativeXMLNode{properties}
		for _, run := range child.Children {
			if run == local {
				continue
			}
			if run.Name != (xml.Name{Space: dialect.drawing, Local: "r"}) {
				paragraph.Children = append(paragraph.Children, run)
				continue
			}
			localRun, err := nativeSingleton(run, dialect.drawing, "rPr", false)
			if err != nil {
				return nil, err
			}
			if localRun != nil {
				if err := validateNativeTextStyleProperties(localRun, dialect, false, theme); err != nil {
					return nil, err
				}
			}
			merged := mergeNativeStyleNodes(defaults, localRun, dialect)
			merged.Attrs = nativeTextPaintAttrs(merged.Attrs)
			merged.Name = xml.Name{Space: dialect.drawing, Local: "rPr"}
			projected := *run
			projected.Children = []*nativeXMLNode{merged}
			for _, item := range run.Children {
				if item != localRun {
					projected.Children = append(projected.Children, item)
				}
			}
			paragraph.Children = append(paragraph.Children, &projected)
		}
		result.Children = append(result.Children, &paragraph)
	}
	return &result, nil
}

func validateNativeTextStyleProperties(node *nativeXMLNode, dialect nativeExtractDialect, paragraph bool, theme nativeResolvedTheme) error {
	attrs := []xml.Name{{Local: "b"}, {Local: "i"}, {Local: "sz"}, {Local: "dirty"}, {Local: "smtClean"}}
	names := []string{"latin", "ea", "cs", "solidFill"}
	if paragraph {
		attrs = []xml.Name{{Local: "algn"}, {Local: "lvl"}, {Local: "marL"}, {Local: "indent"}}
		names = []string{"buNone", "buChar", "defRPr"}
	}
	if err := requireOnlyNativeAttrs(node, attrs...); err != nil {
		return unsupportedNativeTextContent("unmodeled inherited text property: " + err.Error())
	}
	// Validate each source before precedence can hide a malformed value.
	for _, attr := range node.Attrs {
		var err error
		switch attr.Name.Local {
		case "b", "i", "dirty", "smtClean":
			_, err = nativeBool(attr.Value)
		case "sz":
			_, err = parseCanonicalNativeInt(attr.Value, 1, 400000)
		case "lvl":
			_, err = parseCanonicalNativeInt(attr.Value, 0, 8)
		case "marL":
			_, err = parseCanonicalNativeInt(attr.Value, 0, 51206400)
		case "indent":
			_, err = parseCanonicalNativeInt(attr.Value, -51206400, 51206400)
		case "algn":
			_, err = nativeTextAlign(attr.Value)
		}
		if err != nil {
			return err
		}
	}
	allowed := make([]xml.Name, 0, len(names))
	for _, name := range names {
		allowed = append(allowed, xml.Name{Space: dialect.drawing, Local: name})
	}
	if err := requireOnlyNativeChildren(node, allowed...); err != nil {
		return unsupportedNativeTextContent("unmodeled inherited text property child: " + err.Error())
	}
	for _, name := range names {
		child, err := nativeSingleton(node, dialect.drawing, name, false)
		if err != nil {
			return err
		}
		if child == nil {
			continue
		}
		switch name {
		case "defRPr":
			err = validateNativeTextStyleProperties(child, dialect, false, theme)
		case "solidFill":
			_, err = exactNativeSolidColor(child, dialect, theme)
		case "latin", "ea", "cs":
			family, ok := exactNativeAttr(child, "", "typeface")
			if !ok || family == "" || requireOnlyNativeAttrs(child, xml.Name{Local: "typeface"}) != nil || requireOnlyNativeChildren(child) != nil {
				return unsupportedNativeTextContent("invalid inherited typeface metadata")
			}
			_, err = theme.resolveTypeface(family)
		case "buNone":
			err = requireEmptyNativeElement(child)
		case "buChar":
			marker, ok := exactNativeAttr(child, "", "char")
			if !ok || !utf8.ValidString(marker) || utf8.RuneCountInString(marker) != 1 || strings.IndexFunc(marker, unicode.IsControl) >= 0 || requireOnlyNativeAttrs(child, xml.Name{Local: "char"}) != nil || requireOnlyNativeChildren(child) != nil {
				return unsupportedNativeTextContent("invalid inherited bullet metadata")
			}
		}
		if err != nil {
			return unsupportedNativeTextContent("invalid inherited text property: " + err.Error())
		}
	}
	if paragraph {
		none, _ := nativeSingleton(node, dialect.drawing, "buNone", false)
		marker, _ := nativeSingleton(node, dialect.drawing, "buChar", false)
		if none != nil && marker != nil {
			return fmt.Errorf("pptxpatch: conflicting inherited bullet properties")
		}
	}
	return nil
}

// These two Boolean flags track spelling/smart-tag checking, not glyph layout.
// They are validated before cascade resolution and omitted only from the owned
// paint projection. Language, kumimoji, and all other attributes remain strict.
func nativeTextPaintAttrs(attrs []xml.Attr) []xml.Attr {
	result := make([]xml.Attr, 0, len(attrs))
	for _, attr := range attrs {
		if attr.Name.Space == "" && (attr.Name.Local == "dirty" || attr.Name.Local == "smtClean") {
			continue
		}
		result = append(result, attr)
	}
	return result
}

func nativeHasTextCheckingMetadata(node *nativeXMLNode, dialect nativeExtractDialect) bool {
	if node == nil {
		return false
	}
	if node.Name.Space == dialect.drawing && (node.Name.Local == "rPr" || node.Name.Local == "defRPr") {
		for _, attr := range node.Attrs {
			if attr.Name.Space == "" && (attr.Name.Local == "dirty" || attr.Name.Local == "smtClean") {
				return true
			}
		}
	}
	for _, child := range node.Children {
		if nativeHasTextCheckingMetadata(child, dialect) {
			return true
		}
	}
	return false
}

func nativePreserveTextCheckingMetadata(element *NativeElement, node *nativeXMLNode, dialect nativeExtractDialect) {
	if element.Compatibility.Status == NativeCompatibilityStatusRefused || !nativeHasTextCheckingMetadata(node, dialect) {
		return
	}
	element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning, Code: "pptx.text-checking-metadata-preserved",
		Message: "spelling and smart-tag check flags do not affect static text paint; source remains preserve-only because replacement does not round-trip these flags",
	})
}

func mergeNativeStyleNodes(base, override *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	merged := &nativeXMLNode{}
	if base != nil {
		merged.Attrs = append(merged.Attrs, base.Attrs...)
		merged.Children = append(merged.Children, base.Children...)
	}
	if override == nil {
		return merged
	}
	for _, attr := range override.Attrs {
		found := false
		for i := range merged.Attrs {
			if merged.Attrs[i].Name == attr.Name {
				merged.Attrs[i] = attr
				found = true
				break
			}
		}
		if !found {
			merged.Attrs = append(merged.Attrs, attr)
		}
	}
	for _, child := range override.Children {
		found := false
		for i, previous := range merged.Children {
			bulletReplacement := child.Name.Space == dialect.drawing && previous.Name.Space == dialect.drawing && (child.Name.Local == "buNone" || child.Name.Local == "buChar") && (previous.Name.Local == "buNone" || previous.Name.Local == "buChar")
			if previous.Name == child.Name || bulletReplacement {
				if child.Name == (xml.Name{Space: dialect.drawing, Local: "defRPr"}) {
					nested := mergeNativeStyleNodes(previous, child, dialect)
					nested.Name = child.Name
					merged.Children[i] = nested
				} else {
					merged.Children[i] = child
				}
				found = true
				break
			}
		}
		if !found {
			merged.Children = append(merged.Children, child)
		}
	}
	return merged
}
