package docxpatch

import "encoding/xml"

// Font matching descriptors do not supply glyphs or metrics. Keep them in the
// source and in diagnostics; consumers may proceed only with exact font faces.
// Symbol/legacy charsets and unknown descriptors remain unsupported.
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
		return values[0] == "00"
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
