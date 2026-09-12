package docxpatch

import (
	"encoding/xml"
	"strings"
)

// Font matching descriptors do not supply glyphs or metrics. Keep them in the
// source and in diagnostics; consumers may proceed only with exact font faces.
// Active symbol/legacy descriptors and unknown metadata remain unsupported.
func nativeQualifiedFontDescriptor(node, font *nativeXMLNode, ns string) bool {
	if node.Name.Space != ns || !nativeExactContainer(font, xml.Name{Space: ns, Local: "name"}) || len(directNativeChildren(font, ns, node.Name.Local)) != 1 {
		return false
	}
	attrs := []string{"val"}
	if node.Name.Local == "sig" {
		attrs = []string{"usb0", "usb1", "usb2", "usb3", "csb0", "csb1"}
	}
	allowed := make([]xml.Name, len(attrs))
	for i, name := range attrs {
		allowed[i] = xml.Name{Space: ns, Local: name}
	}
	if !nativeExactLeaf(node, allowed...) {
		return false
	}
	values := make([]string, len(attrs))
	for i, name := range attrs {
		value, ok := nativeAttr(node, ns, name)
		if !ok {
			return false
		}
		values[i] = value
	}
	switch node.Name.Local {
	case "panose1":
		return nativeFontDescriptorHex(values[0], 20)
	case "charset":
		return values[0] == "00" || strings.EqualFold(values[0], "EE")
	case "family":
		switch values[0] {
		case "auto", "decorative", "modern", "roman", "script", "swiss":
			return true
		}
	case "pitch":
		switch values[0] {
		case "default", "fixed", "variable":
			return true
		}
	case "sig":
		for _, value := range values {
			if !nativeFontDescriptorHex(value, 8) {
				return false
			}
		}
		return true
	}
	return false
}

// Ignorable declares an extension namespace; it does not authorize ignoring any
// child content. All foreign children still receive the existing diagnostics.
func nativeQualifiedFontTableOwner(root *nativeXMLNode) bool {
	const mc = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	if !nativeExactContainer(root, xml.Name{Space: mc, Local: "Ignorable"}) {
		return false
	}
	value, present := nativeAttr(root, mc, "Ignorable")
	if !present {
		return true
	}
	if value != "w14" {
		return false
	}
	for _, attr := range root.Attrs {
		if attr.Name == (xml.Name{Space: "xmlns", Local: "w14"}) {
			return attr.Value == "http://schemas.microsoft.com/office/word/2010/wordml"
		}
	}
	return false
}

type nativeUnusedFontDescriptor struct {
	name  string
	alias *string
	part  string
	node  *nativeXMLNode
}

func nativePotentiallyUnusedFontDescriptor(node, owner *nativeXMLNode, ns string) bool {
	if node.Name.Space != ns || !nativeExactContainer(owner, xml.Name{Space: ns, Local: "name"}) || len(directNativeChildren(owner, ns, node.Name.Local)) != 1 || !nativeExactLeaf(node, xml.Name{Space: ns, Local: "val"}) {
		return false
	}
	if node.Name.Local == "charset" {
		value, _ := nativeAttr(node, ns, "val")
		return value == "02"
	}
	if node.Name.Local == "notTrueType" {
		_, valid := nativeOnOff(node, ns)
		return valid
	}
	return false
}

func (resolver *nativeLayoutResolver) resolveUnusedFontDescriptors(layout *NativeResolvedLayoutInputV1) {
	used := map[string]bool{}
	uncertain := false
	add := func(properties NativeResolvedRunPropertiesV1) {
		if properties.FontFamily == nil {
			uncertain = true
		} else {
			used[strings.ToLower(*properties.FontFamily)] = true
		}
	}
	for _, run := range layout.Runs {
		add(run.Properties)
	}
	for _, paragraph := range layout.Paragraphs {
		add(paragraph.ParagraphMarkProperties)
		if paragraph.Numbering != nil {
			add(paragraph.Numbering.Marker)
		}
	}
	for _, candidate := range resolver.unusedFontDescriptors {
		if uncertain || used[strings.ToLower(candidate.name)] || candidate.alias != nil && used[strings.ToLower(*candidate.alias)] {
			resolver.addDiagnostic("UNMODELED_FONT_METADATA", resolver.doc.DocumentID, candidate.part, candidate.node, "Active or uncertain symbol/legacy font metadata remains unsupported")
		} else {
			resolver.addDiagnostic("FONT_MATCHING_METADATA_PRESERVED", resolver.doc.DocumentID, candidate.part, candidate.node, "Validated descriptor belongs to an unused font-table face; source is preserved and exact supplied font faces remain required")
		}
	}
}

func nativeFontDescriptorHex(value string, length int) bool {
	if len(value) != length {
		return false
	}
	for _, c := range value {
		if !(c >= '0' && c <= '9' || c >= 'a' && c <= 'f' || c >= 'A' && c <= 'F') {
			return false
		}
	}
	return true
}
