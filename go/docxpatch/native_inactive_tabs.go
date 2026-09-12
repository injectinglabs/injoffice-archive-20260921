package docxpatch

import (
	"encoding/xml"
	"strings"
)

// Non-bar stops only affect tab positioning. Keep all source stops intact;
// qualify their inactivity only for proven plain text, never inferred fields.
func nativeExactInactiveTabCandidates(node *nativeXMLNode, ns string) bool {
	if owner := node.parent; owner != nil && owner.parent != nil && owner.parent.Name == (xml.Name{Space: ns, Local: "style"}) {
		if !nativeExactContainer(owner.parent, xml.Name{Space: ns, Local: "type"}, xml.Name{Space: ns, Local: "styleId"}, xml.Name{Space: ns, Local: "default"}, xml.Name{Space: ns, Local: "customStyle"}) {
			return false
		}
	}
	if !nativeExactContainer(node) || len(node.Children) == 0 || len(node.Children) > 64 {
		return false
	}
	positions := map[int64]bool{}
	for _, tab := range node.Children {
		if tab.Name != (xml.Name{Space: ns, Local: "tab"}) || !nativeExactLeaf(tab, xml.Name{Space: ns, Local: "val"}, xml.Name{Space: ns, Local: "pos"}, xml.Name{Space: ns, Local: "leader"}) {
			return false
		}
		kind, ok := nativeAttr(tab, ns, "val")
		if !ok || (kind != "left" && kind != "center" && kind != "right" && kind != "decimal" && kind != "clear") {
			return false
		}
		position, ok := nativeNonnegativeInt64Attr(tab, ns, "pos")
		if !ok || position > nativeMaxTwipsForMilliPoints || positions[position] {
			return false
		}
		positions[position] = true
		if leader, present := nativeAttr(tab, ns, "leader"); present && leader != "none" {
			return false
		}
	}
	return true
}

func nativePlainParagraphWithoutTabs(node *nativeXMLNode, ns string) bool {
	if node == nil || node.Name != (xml.Name{Space: ns, Local: "p"}) || !nativeExactRevisionContainer(node, ns, "rsidR", "rsidRDefault", "rsidP", "rsidRPr") {
		return false
	}
	for _, child := range node.Children {
		if child.Name == (xml.Name{Space: ns, Local: "pPr"}) {
			continue
		}
		if child.Name != (xml.Name{Space: ns, Local: "r"}) || !nativeExactRevisionContainer(child, ns, "rsidR", "rsidRPr", "rsidDel") {
			return false
		}
		for _, content := range child.Children {
			if content.Name == (xml.Name{Space: ns, Local: "rPr"}) {
				continue
			}
			if content.Name != (xml.Name{Space: ns, Local: "t"}) || len(content.Children) > 0 || strings.Contains(content.Text, "\t") {
				return false
			}
			for _, attr := range content.Attrs {
				if !nativeSettingsNamespaceDeclaration(attr) && (attr.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || (attr.Value != "preserve" && attr.Value != "default")) {
					return false
				}
			}
		}
	}
	return true
}
