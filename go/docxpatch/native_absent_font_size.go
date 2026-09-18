package docxpatch

import "encoding/xml"

// NativeDocxAbsentDefaultSizeShapeV1 names which source shape proved the
// package states no default run size. The two shapes are not interchangeable:
// Microsoft Word 16.112 resolves them to different sizes, so a consumer that
// declares one host default per shape needs the shape as a source fact.
const (
	// NativeDocxAbsentDocumentDefaultsV1 is a package carrying no w:docDefaults
	// record at all, whether because it has no styles part or because its
	// styles part declares none (ECMA-376 17.7.2 makes both optional).
	NativeDocxAbsentDocumentDefaultsV1 = "absent-document-defaults"
	// NativeDocxSizelessDocumentDefaultsV1 is a package whose w:docDefaults
	// record exists and states no w:sz.
	NativeDocxSizelessDocumentDefaultsV1 = "sizeless-document-defaults"
)

// NativeDocxHostDefaultSizeHalfPointsV1 is the single definition of the
// read-only host default size per proven source shape. Both values were read
// directly out of the Tf operators of Microsoft Word 16.112's own PDF exports
// of corpus packages that state no size (Word writes text on a 1/300 in grid,
// so 50 units == 12 pt and 42 units == 10 pt): a package carrying no
// w:docDefaults record is laid out at 12 pt, and one whose w:docDefaults states
// no w:sz at 10 pt. Nothing else may spell these numbers.
func NativeDocxHostDefaultSizeHalfPointsV1(shape string) (int, bool) {
	switch shape {
	case NativeDocxAbsentDocumentDefaultsV1:
		return 24, true
	case NativeDocxSizelessDocumentDefaultsV1:
		return 20, true
	}
	return 0, false
}

// Source absence is not a size recommendation. A read-only consumer may choose
// its own declared default under ECMA-376 17.3.2.38; strict layout stays absent.
type NativeDocxAbsentFontSizeV1 struct {
	ScopeKind     string `json:"scope_kind"`
	ScopeID       string `json:"scope_id"`
	PartName      string `json:"part_name"`
	Path          string `json:"path"`
	PackageSHA256 string `json:"package_sha256"`
}

func nativeAbsentFontSizes(data []byte) ([]NativeDocxAbsentFontSizeV1, string, error) {
	r, err := newNativeLayoutResolver(data, NativeExtractionOptions{})
	if err != nil {
		return nil, "", err
	}
	layout, err := r.resolve()
	if err != nil {
		return nil, "", err
	}
	shape := r.absentDefaultSizeShape()
	if shape == "" {
		return nil, "", nil
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
		// The reserved role is attested by w:type in the note part, not by the id
		// literal: Word writes -1/0 where LibreOffice writes 0/1 for the same pair.
		if note.Anchor == nil || note.NativeStoryID == nil || note.NoteRole == "" || note.NoteRole == "content" {
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
	consider := func(p *NativeParagraphV1) {
		resolved, ok := paragraphs[p.ID]
		if !ok || !r.absentStyleSize(resolved.AppliedStyles, "paragraph") {
			return
		}
		node := r.nodeForAnchor(p.Anchor)
		if node == nil || node.Name != (xml.Name{Space: r.wordNS, Local: "p"}) {
			return
		}
		pprs := directNativeChildren(node, r.wordNS, "pPr")
		if len(pprs) > 1 {
			return
		}
		var ppr *nativeXMLNode
		if len(pprs) == 1 {
			ppr = pprs[0]
		}
		if !r.exactSizeStyleReference(ppr, "pStyle", resolved.StyleID) {
			return
		}
		if resolved.ParagraphMarkProperties.FontSizeHalfPoint == nil && r.absentOwnerRunSize(ppr) {
			add("paragraph-mark", p.ID, p.Anchor)
		}
		// A numbered paragraph's marker is a run of its own and needs its own
		// size. Its run layer is exactly three layers -- the paragraph style
		// cascade, the numbering level's own w:rPr and the paragraph's direct
		// mark rPr -- and the first and third are the same two the paragraph-mark
		// arm above already proved size-free, so only the level is read here.
		if resolved.Numbering != nil && resolved.Numbering.Marker.FontSizeHalfPoint == nil && r.absentOwnerRunSize(ppr) && r.absentNumberingLevelSize(resolved.Numbering) {
			add("numbering-marker", p.ID, p.Anchor)
		}
		for _, run := range p.Runs {
			rr, ok := runs[run.ID]
			if !ok || rr.Properties.FontSizeHalfPoint != nil || !r.absentRunCharacterStyleSize(rr) {
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
	for _, story := range stories {
		for _, block := range story.Blocks {
			if block.Paragraph != nil {
				consider(block.Paragraph)
				continue
			}
			// Cell paragraphs qualify only when the whole table-style chain is
			// also size-free, so no table cascade exception can apply.
			if block.Table == nil || !r.absentTableStyleSize(block.Table.TableStyleID) {
				continue
			}
			for _, row := range block.Table.Rows {
				for _, cell := range row.Cells {
					for index := range cell.Paragraphs {
						consider(&cell.Paragraphs[index])
					}
				}
			}
		}
	}
	if len(facts) > 1000 || len(facts) == 0 {
		return nil, "", nil
	}
	return facts, shape, nil
}

// absentRunCharacterStyleSize proves the character layer of one run adds no run
// size. An empty applied chain is read as complete only when the run resolved no
// character style reference at all: resolution applies the package's declared
// default character style to every run that states no w:rStyle, so a nil
// resolved style is itself the proof that no default was declared and none was
// dropped. A named style that resolved to an empty chain is an unresolved
// reference and still refuses.
func (r *nativeLayoutResolver) absentRunCharacterStyleSize(run NativeResolvedRunV1) bool {
	if len(run.AppliedCharacterStyles) == 0 {
		return run.CharacterStyle == nil
	}
	return r.absentStyleSize(run.AppliedCharacterStyles, "character")
}

// absentNumberingLevelSize proves the numbering level that supplies a marker
// states no run size of its own. The level is re-read from the numbering part
// by the ids the resolved marker carries rather than trusted from the resolved
// value, so the source shape is the fact and a resolver that simply failed to
// read a size cannot pass for an omission.
func (r *nativeLayoutResolver) absentNumberingLevelSize(numbering *NativeResolvedNumberingV1) bool {
	instance := r.nums[numbering.NumID]
	if instance == nil || instance.abstractID != numbering.AbstractNumID {
		return false
	}
	abstract := r.abstractNums[instance.abstractID]
	if abstract == nil {
		return false
	}
	level := r.effectiveNumberingLevel(instance, abstract, numbering.Level)
	if level == nil || level.node == nil || level.node.Name != (xml.Name{Space: r.wordNS, Local: "lvl"}) {
		return false
	}
	// A w:rStyle inside the level's run properties names a character style whose
	// own size the marker cascade never reads, so its presence refuses rather
	// than reporting an omission the style may fill. absentOwnerRunSize tolerates
	// it for the run and paragraph-mark scopes, which resolve it separately.
	for _, rpr := range directNativeChildren(level.node, r.wordNS, "rPr") {
		if len(directNativeChildren(rpr, r.wordNS, "rStyle")) > 0 {
			return false
		}
	}
	return r.absentOwnerRunSize(level.node)
}

func (r *nativeLayoutResolver) absentDefaultSizeShape() string {
	// A package with no styles part carries no w:docDefaults at all, which is
	// strictly stronger evidence of an absent default size than a styles part
	// whose docDefaults happen to state none. ECMA-376 17.7.2 makes the part
	// optional, so its absence is a source fact, not an unread default.
	if r.parts.StylesPart == nil {
		if len(r.styles) == 0 {
			return NativeDocxAbsentDocumentDefaultsV1
		}
		return ""
	}
	root, err := parseNativeXML(*r.parts.StylesPart, r.pkg.files[*r.parts.StylesPart])
	// mc:Ignorable only declares ignorable namespaces, exactly as the main
	// part extractor already accepts on part roots; it carries no size.
	if err != nil || root.Name != (xml.Name{Space: r.wordNS, Local: "styles"}) || !nativeExactContainer(root, xml.Name{Space: nativeMCNamespace, Local: "Ignorable"}) {
		return ""
	}
	defaults := directNativeChildren(root, r.wordNS, "docDefaults")
	if len(defaults) > 1 {
		return ""
	}
	if len(defaults) == 0 {
		return NativeDocxAbsentDocumentDefaultsV1
	}
	if !nativeExactContainer(defaults[0]) {
		return ""
	}
	for _, child := range defaults[0].Children {
		if child.Name.Space != r.wordNS || child.Name.Local != "rPrDefault" && child.Name.Local != "pPrDefault" {
			return ""
		}
		if child.Name.Local == "pPrDefault" {
			for _, ppr := range directNativeChildren(child, r.wordNS, "pPr") {
				if len(directNativeChildren(ppr, r.wordNS, "rPr")) > 0 {
					return ""
				}
			}
		}
	}
	rprs := directNativeChildren(defaults[0], r.wordNS, "rPrDefault")
	if len(rprs) > 1 {
		return ""
	}
	if len(rprs) == 0 {
		return NativeDocxSizelessDocumentDefaultsV1
	}
	if !nativeExactContainer(rprs[0]) {
		return ""
	}
	for _, child := range rprs[0].Children {
		if child.Name != (xml.Name{Space: r.wordNS, Local: "rPr"}) {
			return ""
		}
	}
	if !r.absentOwnerRunSize(rprs[0]) {
		return ""
	}
	return NativeDocxSizelessDocumentDefaultsV1
}

func (r *nativeLayoutResolver) absentStyleSize(ids []string, kind string) bool {
	// A complete explicit chain is required, never a missing/default style guess.
	// The one exception is a package with no styles part: there is no default
	// style to guess and no definition that could have been missed, so the
	// empty chain is the complete chain rather than an unresolved reference.
	if len(ids) == 0 {
		return r.parts.StylesPart == nil && len(r.styles) == 0
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
		case "rStyle", "rFonts", "b", "i", "cs", "rtl", "vanish", "bCs", "iCs", "u", "color", "highlight", "lang", "vertAlign", "kern", "noProof":
			// Their independently retained diagnostics still qualify or refuse
			// actual rendering. No unknown/complex-size override is bypassed.
		default:
			return false
		}
	}
	return true
}

// absentTableStyleSize proves the table-style cascade adds no run size. A
// table without a style reference uses the declared default table style, so
// that chain is checked rather than assumed empty; conditional table-style
// regions are refused because their run layers are not evaluated here.
func (r *nativeLayoutResolver) absentTableStyleSize(styleID *string) bool {
	id := ""
	if styleID != nil {
		id = *styleID
	} else {
		for _, definition := range r.styles {
			if definition == nil || definition.kind != "table" || definition.node == nil {
				continue
			}
			value, present := nativeAttr(definition.node, r.wordNS, "default")
			if !present {
				continue
			}
			isDefault, valid := nativeLexicalOnOff(value)
			if !valid {
				return false
			}
			if isDefault {
				if id != "" {
					return false
				}
				id = definition.id
			}
		}
		if id == "" {
			return true
		}
	}
	ids := []string{}
	seen := map[string]bool{}
	for current := id; current != ""; {
		if seen[current] || len(ids) >= NativeDOCXMaxDepth {
			return false
		}
		seen[current] = true
		definition := r.styles["table\x00"+current]
		if definition == nil {
			return false
		}
		ids = append(ids, current)
		current = definition.basedOn
	}
	for left, right := 0, len(ids)-1; left < right; left, right = left+1, right-1 {
		ids[left], ids[right] = ids[right], ids[left]
	}
	if !r.absentStyleSize(ids, "table") {
		return false
	}
	for _, id := range ids {
		if len(directNativeChildren(r.styles["table\x00"+id].node, r.wordNS, "tblStylePr")) > 0 {
			return false
		}
	}
	return true
}
