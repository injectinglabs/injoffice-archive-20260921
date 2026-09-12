package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// A default numbering style with only UI metadata contributes no list or text
// formatting. Do not infer this for basedOn/link/numPr/rPr or extension content.
func nativeEmptyDefaultNumberingStyle(node, root *nativeXMLNode, ns string) bool {
	// Use the same bounded root/ignorable declaration policy as font tables.
	if !nativeQualifiedFontTableOwner(root) || !nativeExactContainer(node, xml.Name{Space: ns, Local: "type"}, xml.Name{Space: ns, Local: "styleId"}, xml.Name{Space: ns, Local: "default"}) {
		return false
	}
	kind, _ := nativeAttr(node, ns, "type")
	id, _ := nativeAttr(node, ns, "styleId")
	value, present := nativeAttr(node, ns, "default")
	enabled, valid := nativeLexicalOnOff(value)
	if kind != "numbering" || !nativeIDPattern.MatchString(id) || !present || !valid || !enabled {
		return false
	}
	count := 0
	for _, style := range directNativeChildren(root, ns, "style") {
		other, _ := nativeAttr(style, ns, "styleId")
		if other == id {
			count++
		}
	}
	if count != 1 {
		return false
	}
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != ns || seen[child.Name.Local] || !nativeExactLeaf(child, xml.Name{Space: ns, Local: "val"}) {
			return false
		}
		seen[child.Name.Local] = true
		switch child.Name.Local {
		case "name":
			value, ok := nativeAttr(child, ns, "val")
			if !ok || !nativeBoundedResolvedString(value, 256) {
				return false
			}
		case "uiPriority":
			raw, ok := nativeAttr(child, ns, "val")
			value, err := strconv.ParseUint(raw, 10, 7)
			if !ok || err != nil || value > 99 || strconv.FormatUint(value, 10) != raw {
				return false
			}
		case "semiHidden", "unhideWhenUsed":
			if _, ok := nativeOnOff(child, ns); !ok {
				return false
			}
		default:
			return false
		}
	}
	return true
}
