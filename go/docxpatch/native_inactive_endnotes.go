package docxpatch

import (
	"encoding/xml"
	"strings"
)

// This proof is deliberately package-wide: hidden, unreferenced, foreign and
// alternate content cannot become an implicit license to ignore note placement.
// It changes no source bytes and confers no section mutation capability.
func (e *nativeExtractor) proveNoContentEndnotes() bool {
	if e.noEndnotesChecked {
		return e.noEndnotesProven
	}
	e.noEndnotesChecked = true
	relatedPart := ""
	for source, rels := range e.pkg.rels {
		for _, rel := range rels {
			if !strings.HasSuffix(rel.Type, "/endnotes") {
				continue
			}
			if source != e.mainPart || rel.Type != strings.TrimSuffix(e.relNS, "/")+"/endnotes" || rel.External || relatedPart != "" {
				return false
			}
			relatedPart = rel.PartName
		}
	}
	budget, foundPart := 8*1024*1024, ""
	for part, data := range e.pkg.files {
		ct := e.pkg.contentTypes[part]
		if !strings.HasSuffix(strings.ToLower(part), ".xml") && !strings.HasSuffix(ct, "+xml") && ct != "application/xml" && ct != "text/xml" {
			continue
		}
		budget -= len(data)
		if budget < 0 {
			return false
		}
		root, err := parseNativeXML(part, data)
		if err != nil {
			return false
		}
		if root.Name.Local == "endnotes" || ct == "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml" {
			if part != relatedPart || foundPart != "" || root.Name != (xml.Name{Space: e.wordNS, Local: "endnotes"}) || !nativeExactContainer(root, xml.Name{Space: "http://schemas.openxmlformats.org/markup-compatibility/2006", Local: "Ignorable"}) {
				return false
			}
			foundPart = part
			seen := map[string]bool{}
			for _, note := range root.Children {
				id, _ := nativeAttr(note, e.wordNS, "id")
				kind, _ := nativeAttr(note, e.wordNS, "type")
				role := "separator"
				if kind == "continuationSeparator" {
					role = "continuation-separator"
				}
				if note.Name != (xml.Name{Space: e.wordNS, Local: "endnote"}) || seen[id] || !nativeExactContainer(note, xml.Name{Space: e.wordNS, Local: "id"}, xml.Name{Space: e.wordNS, Local: "type"}) || !(id == "-1" && kind == "separator" || id == "0" && kind == "continuationSeparator") {
					return false
				}
				seen[id] = true
				if len(note.Children) != 1 || !nativeExactRevisionContainer(note.Children[0], e.wordNS, "rsidR", "rsidRDefault", "rsidP", "rsidRPr") {
					return false
				}
				// Own temporary node headers, never mutate the parsed source tree.
				paragraph := *note.Children[0]
				paragraph.Attrs = nil
				for _, property := range directNativeChildren(&paragraph, e.wordNS, "pPr") {
					if !nativeExactContainer(property) {
						return false
					}
				}
				copyNote := *note
				copyNote.Children = []*nativeXMLNode{&paragraph}
				if !nativeExactNoteSentinel(&copyNote, e.wordNS, role) {
					return false
				}
			}
			continue
		}
		var inspect func(*nativeXMLNode) bool
		inspect = func(n *nativeXMLNode) bool {
			switch n.Name.Local {
			case "endnote":
				// Settings may select the two reserved separator IDs. These
				// exact empty leaves are not content-note stories/references.
				id, _ := nativeAttr(n, e.wordNS, "id")
				if n.Name.Space != e.wordNS || (id != "-1" && id != "0") || !nativeExactLeaf(n, xml.Name{Space: e.wordNS, Local: "id"}) || n.parent == nil || n.parent.Name != (xml.Name{Space: e.wordNS, Local: "endnotePr"}) || n.parent.parent != root || root.Name != (xml.Name{Space: e.wordNS, Local: "settings"}) || ct != "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml" {
					return false
				}
				matches := 0
				for _, rel := range e.pkg.rels[e.mainPart] {
					if rel.Type == strings.TrimSuffix(e.relNS, "/")+"/settings" && !rel.External && rel.PartName == part {
						matches++
					}
				}
				if matches != 1 {
					return false
				}
				count := 0
				for _, sibling := range n.parent.Children {
					if sibling.Name.Local == "endnote" {
						other, _ := nativeAttr(sibling, e.wordNS, "id")
						if other == id {
							count++
						}
					}
				}
				return count == 1
			case "endnotes", "endnoteReference", "endnoteRef":
				return false
			}
			for _, c := range n.Children {
				if !inspect(c) {
					return false
				}
			}
			return true
		}
		if !inspect(root) {
			return false
		}
	}
	e.noEndnotesProven = foundPart == relatedPart
	return e.noEndnotesProven
}
