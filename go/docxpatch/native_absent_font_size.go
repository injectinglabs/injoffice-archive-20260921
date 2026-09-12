package docxpatch

import "encoding/xml"

// Source absence is not a size recommendation. A read-only consumer may choose
// its own declared default under ECMA-376 17.3.2.38; strict layout stays absent.
type NativeDocxAbsentFontSizeV1 struct {
	ScopeKind     string `json:"scope_kind"`
	ScopeID       string `json:"scope_id"`
	PartName      string `json:"part_name"`
	Path          string `json:"path"`
	PackageSHA256 string `json:"package_sha256"`
}

func nativeAbsentFontSizes(data []byte) ([]NativeDocxAbsentFontSizeV1, error) {
	r, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, err
	}
	layout, err := r.resolve()
	if err != nil {
		return nil, err
	}
	if !r.absentDefaultSize() {
		return nil, nil
	}
	paragraphs := map[string]NativeResolvedParagraphV1{}
	runs := map[string]NativeResolvedRunV1{}
	for _, p := range layout.Paragraphs {
		paragraphs[p.ParagraphID] = p
	}
	for _, run := range layout.Runs {
		runs[run.RunID] = run
	}
	var facts []NativeDocxAbsentFontSizeV1
	add := func(kind, id string, anchor NativeSourceAnchorV1) {
		facts = append(facts, NativeDocxAbsentFontSizeV1{ScopeKind: kind, ScopeID: id, PartName: anchor.PartName, Path: anchor.Path, PackageSHA256: r.doc.Source.PackageSHA256})
	}
	stories := append([]NativeStoryV1{r.doc.Body}, r.doc.Headers...)
	stories = append(stories, r.doc.Footers...)
	for _, note := range r.doc.Notes {
		if note.Anchor == nil || note.NativeStoryID == nil || !(note.NoteRole == "separator" && *note.NativeStoryID == "-1" || note.NoteRole == "continuation-separator" && *note.NativeStoryID == "0") {
			continue
		}
		node := r.nodeForAnchor(*note.Anchor)
		if node == nil || !nativeExactNoteSentinel(node, r.wordNS, note.NoteRole) {
			continue
		}
		clean := true
		for _, diagnostic := range r.doc.Unsupported {
			if diagnostic.ScopeID == note.ID || diagnostic.Anchor != nil && diagnostic.Anchor.PartName == note.PartName {
				clean = false
			}
		}
		if clean {
			stories = append(stories, note)
		}
	}
	for _, story := range stories {
		for _, block := range story.Blocks {
			// Table cascade exceptions are deliberately outside this first policy.
			p := block.Paragraph
			if p == nil {
				continue
			}
			resolved, ok := paragraphs[p.ID]
			if !ok || resolved.Numbering != nil || !r.absentStyleSize(resolved.AppliedStyles, "paragraph") {
				continue
			}
			node := r.nodeForAnchor(p.Anchor)
			if node == nil || node.Name != (xml.Name{Space: r.wordNS, Local: "p"}) {
				continue
			}
			pprs := directNativeChildren(node, r.wordNS, "pPr")
			if len(pprs) > 1 {
				continue
			}
			var ppr *nativeXMLNode
			if len(pprs) == 1 {
				ppr = pprs[0]
			}
			if !r.exactSizeStyleReference(ppr, "pStyle", resolved.StyleID) {
				continue
			}
			if resolved.ParagraphMarkProperties.FontSizeHalfPoint == nil && r.absentOwnerRunSize(ppr) {
				add("paragraph-mark", p.ID, p.Anchor)
			}
			for _, run := range p.Runs {
				rr, ok := runs[run.ID]
				if !ok || rr.Properties.FontSizeHalfPoint != nil || !r.absentStyleSize(rr.AppliedCharacterStyles, "character") {
					continue
				}
				owner := r.nodeForAnchor(run.Anchor)
				for owner != nil && owner.Name != (xml.Name{Space: r.wordNS, Local: "r"}) {
					owner = owner.parent
				}
				if owner == nil || !r.absentOwnerRunSize(owner) {
					continue
				}
				rpr := firstDirectNativeChild(owner, r.wordNS, "rPr")
				if !r.exactSizeStyleReference(rpr, "rStyle", rr.CharacterStyle) {
					continue
				}
				add("run", run.ID, run.Anchor)
			}
		}
	}
	if len(facts) > 1000 {
		return nil, nil
	}
	return facts, nil
}

func (r *nativeLayoutResolver) absentDefaultSize() bool {
	if r.parts.StylesPart == nil {
		return false
	}
	root, err := parseNativeXML(*r.parts.StylesPart, r.pkg.files[*r.parts.StylesPart])
	if err != nil || root.Name != (xml.Name{Space: r.wordNS, Local: "styles"}) || !nativeExactContainer(root) {
		return false
	}
	defaults := directNativeChildren(root, r.wordNS, "docDefaults")
	if len(defaults) > 1 {
		return false
	}
	if len(defaults) == 0 {
		return true
	}
	if !nativeExactContainer(defaults[0]) {
		return false
	}
	for _, child := range defaults[0].Children {
		if child.Name.Space != r.wordNS || child.Name.Local != "rPrDefault" && child.Name.Local != "pPrDefault" {
			return false
		}
		if child.Name.Local == "pPrDefault" {
			for _, ppr := range directNativeChildren(child, r.wordNS, "pPr") {
				if len(directNativeChildren(ppr, r.wordNS, "rPr")) > 0 {
					return false
				}
			}
		}
	}
	rprs := directNativeChildren(defaults[0], r.wordNS, "rPrDefault")
	if len(rprs) > 1 {
		return false
	}
	if len(rprs) == 0 {
		return true
	}
	if !nativeExactContainer(rprs[0]) {
		return false
	}
	for _, child := range rprs[0].Children {
		if child.Name != (xml.Name{Space: r.wordNS, Local: "rPr"}) {
			return false
		}
	}
	return r.absentOwnerRunSize(rprs[0])
}

func (r *nativeLayoutResolver) absentStyleSize(ids []string, kind string) bool {
	// A complete explicit chain is required, never a missing/default style guess.
	if len(ids) == 0 {
		return false
	}
	previous := ""
	for _, id := range ids {
		s := r.styles[kind+"\x00"+id]
		if s == nil || s.kind != kind || s.basedOn != previous || !r.absentOwnerRunSize(s.node) {
			return false
		}
		for _, ppr := range directNativeChildren(s.node, r.wordNS, "pPr") {
			if len(directNativeChildren(ppr, r.wordNS, "rPr")) > 0 {
				return false
			}
		}
		based := directNativeChildren(s.node, r.wordNS, "basedOn")
		if len(based) > 1 {
			return false
		}
		if len(based) == 1 {
			value, ok := nativeAttr(based[0], r.wordNS, "val")
			if !ok || !nativeIDPattern.MatchString(value) || value != previous || !nativeExactLeaf(based[0], xml.Name{Space: r.wordNS, Local: "val"}) {
				return false
			}
		}
		previous = id
	}
	return true
}

func (r *nativeLayoutResolver) exactSizeStyleReference(owner *nativeXMLNode, local string, resolved *string) bool {
	if owner == nil {
		return true
	}
	nodes := directNativeChildren(owner, r.wordNS, local)
	if len(nodes) > 1 {
		return false
	}
	if len(nodes) == 0 {
		return true
	}
	value, ok := nativeAttr(nodes[0], r.wordNS, "val")
	return ok && resolved != nil && value == *resolved && nativeIDPattern.MatchString(value) && nativeExactLeaf(nodes[0], xml.Name{Space: r.wordNS, Local: "val"})
}

func (r *nativeLayoutResolver) absentOwnerRunSize(owner *nativeXMLNode) bool {
	if owner == nil {
		return true
	}
	nodes := directNativeChildren(owner, r.wordNS, "rPr")
	if len(nodes) > 1 {
		return false
	}
	if len(nodes) == 0 {
		return true
	}
	node := nodes[0]
	if !nativeExactContainer(node) {
		return false
	}
	for _, child := range node.Children {
		if child.Name.Space != r.wordNS {
			return false
		}
		switch child.Name.Local {
		case "rStyle", "rFonts", "b", "i", "rtl", "vanish", "bCs", "iCs", "u", "color", "highlight", "lang", "vertAlign", "kern", "noProof":
			// Their independently retained diagnostics still qualify or refuse
			// actual rendering. No unknown/complex-size override is bypassed.
		default:
			return false
		}
	}
	return true
}
