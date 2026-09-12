package docxpatch

import (
	"encoding/xml"
	"strings"
)

// Non-bar stops only affect tab positioning. Keep all source stops intact;
// qualify their inactivity only for proven consumers, never inferred fields.
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

// Only extraction-qualified decimal page fields can extend the plain-text
// inactivity proof. The resolver owns both this native paragraph and its raw
// indexed XML; caller-authored field claims never enter this source join.
// PAGE/NUMPAGES render generated digits, not their cached result text.
func (r *nativeLayoutResolver) nativeParagraphWithoutTabConsumers(node *nativeXMLNode, paragraph *NativeParagraphV1) bool {
	if node == nil || paragraph == nil {
		return false
	}
	fields := map[*nativeXMLNode]string{}
	for _, run := range paragraph.Runs {
		if run.PageField != "PAGE" && run.PageField != "NUMPAGES" {
			continue
		}
		if run.Kind != "text" || run.Text == nil || *run.Text != "" || run.Anchor.PartName != paragraph.Anchor.PartName {
			return false
		}
		owner := r.nodeForAnchor(run.Anchor)
		for owner != nil && owner.Name != (xml.Name{Space: r.wordNS, Local: "r"}) {
			owner = owner.parent
		}
		if owner == nil || fields[owner] != "" {
			return false
		}
		if owner.parent != node && (owner.parent == nil || owner.parent.Name != (xml.Name{Space: r.wordNS, Local: "fldSimple"}) || owner.parent.parent != node) {
			return false
		}
		fields[owner] = run.PageField
	}
	if len(fields) == 0 {
		return nativePlainParagraphWithoutTabs(node, r.wordNS)
	}
	projection := *node
	projection.Children = nil
	for i := 0; i < len(node.Children); i++ {
		child := node.Children[i]
		if child.Name == (xml.Name{Space: r.wordNS, Local: "fldSimple"}) && len(child.Children) == 1 && fields[child.Children[0]] != "" {
			// This same exact owner/result already passed extractParagraphRuns.
			delete(fields, child.Children[0])
			continue
		}
		if child.Name == (xml.Name{Space: r.wordNS, Local: "r"}) && hasNativeFieldBegin(child, r.wordNS) && i+4 < len(node.Children) && fields[node.Children[i+3]] != "" {
			for _, member := range node.Children[i : i+5] {
				if member.parent != node || member.Name != (xml.Name{Space: r.wordNS, Local: "r"}) {
					return false
				}
			}
			// extractFlatPageField emits a PageField only for this exact five-run
			// sequence and preserves the result anchor inside its fourth run.
			delete(fields, node.Children[i+3])
			i += 4
			continue
		}
		projection.Children = append(projection.Children, child)
	}
	return len(fields) == 0 && nativePlainParagraphWithoutTabs(&projection, r.wordNS)
}
