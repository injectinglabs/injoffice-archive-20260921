// D11 breadth: footnotes and endnotes. Same Patch{Replace/Add} machinery
// imagewrite.go introduced (patch.go), applied to a different OOXML shape:
// a footnote/endnote reference is a RUN APPENDED to an existing paragraph's
// existing runs (never a whole new paragraph the way an image is — a
// footnote marker sits inline, right after the word it's attached to),
// plus a note DEFINITION appended into word/footnotes.xml or
// word/endnotes.xml (created fresh if the document has neither yet).
//
// v1 anchor point: the reference is appended at the END of the target
// paragraph's existing content, not at an arbitrary character offset within
// it — mid-paragraph caret-position addressing would need the same
// character-offset scheme docxpatch's paragraph-level model deliberately
// doesn't have (Extract/Apply work at whole-paragraph granularity
// everywhere else too). Stated plainly rather than left silent.
package docxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"path"
	"strconv"
	"strings"
	"unicode/utf8"
)

// NoteKind selects footnotes or endnotes — identical structure, different
// part name / element names / relationship type / content type.
type NoteKind string

const (
	Footnote NoteKind = "footnote"
	Endnote  NoteKind = "endnote"
)

type noteShape struct {
	part        string // word/footnotes.xml or word/endnotes.xml
	rootEl      string // w:footnotes / w:endnotes
	noteEl      string // w:footnote / w:endnote
	refEl       string // w:footnoteReference / w:endnoteReference
	markEl      string // w:footnoteRef / w:endnoteRef
	styleID     string
	textStyleID string
	relType     string
	contentType string
}

func shapeFor(kind NoteKind) (noteShape, error) {
	switch kind {
	case Footnote:
		return noteShape{
			part: "word/footnotes.xml", rootEl: "w:footnotes", noteEl: "w:footnote",
			refEl: "w:footnoteReference", markEl: "w:footnoteRef", styleID: "FootnoteReference",
			textStyleID: "FootnoteText",
			relType:     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
			contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml",
		}, nil
	case Endnote:
		return noteShape{
			part: "word/endnotes.xml", rootEl: "w:endnotes", noteEl: "w:endnote",
			refEl: "w:endnoteReference", markEl: "w:endnoteRef", styleID: "EndnoteReference",
			textStyleID: "EndnoteText",
			relType:     "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes",
			contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml",
		}, nil
	default:
		return noteShape{}, fmt.Errorf("docxpatch: unknown note kind %q", kind)
	}
}

// noteXMLNS is declared locally on the root element — never trust what the
// source document declared elsewhere, same defensive posture as
// imageParagraphXML.
func freshNotesPart(shape noteShape, wordNS string) string {
	noteXMLNS := `xmlns:w="` + wordNS + `"`
	// Reserved ids -1 (separator) and 0 (continuation separator) are exact
	// instruction-only package sentinels. They carry no visible fallback text.
	sepType := map[bool]string{true: "separator", false: "continuationSeparator"}
	var b strings.Builder
	fmt.Fprintf(&b, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><%s %s>`, shape.rootEl, noteXMLNS)
	for _, id := range []int{-1, 0} {
		fmt.Fprintf(&b, `<%s w:type=%q w:id="%d"><w:p><w:r>`, shape.noteEl, sepType[id == -1], id)
		if id == -1 {
			b.WriteString(`<w:separator/>`)
		} else {
			b.WriteString(`<w:continuationSeparator/>`)
		}
		b.WriteString(`</w:r></w:p></` + shape.noteEl + `>`)
	}
	b.WriteString(`</` + shape.rootEl + `>`)
	return b.String()
}

func noteBodyXML(shape noteShape, id int, text string) string {
	var b strings.Builder
	fmt.Fprintf(&b, `<%s xmlns:w=%q w:id="%d"><w:p><w:pPr><w:pStyle w:val=%q/></w:pPr><w:r><w:rPr><w:rStyle w:val=%q/></w:rPr>`, shape.noteEl, shapeWordNS(shape), id, shape.textStyleID, shape.styleID)
	fmt.Fprintf(&b, `<%s/>`, shape.markEl)
	b.WriteString(`</w:r><w:r><w:t xml:space="preserve"> `)
	b.WriteString(xmlEscape(text))
	b.WriteString(`</w:t></w:r></w:p></` + shape.noteEl + `>`)
	return b.String()
}

func noteRefRunXML(shape noteShape, id int) string {
	return fmt.Sprintf(`<w:r xmlns:w=%q><w:rPr><w:rStyle w:val=%q/></w:rPr><%s w:id="%d"/></w:r>`, shapeWordNS(shape), shape.styleID, shape.refEl, id)
}

func shapeWordNS(shape noteShape) string {
	if strings.HasPrefix(shape.relType, relBaseStrict) {
		return wordMLStrict
	}
	return wordMLTransitional
}

func nativeOpeningQName(data []byte, node *nativeXMLNode) (string, error) {
	if node.Start < 0 || node.Start >= int64(len(data)) || data[node.Start] != '<' {
		return "", fmt.Errorf("invalid XML element boundary")
	}
	start := int(node.Start) + 1
	end := start
	for end < len(data) && data[end] != ' ' && data[end] != '\t' && data[end] != '\r' && data[end] != '\n' && data[end] != '>' && data[end] != '/' {
		end++
	}
	if end == start {
		return "", fmt.Errorf("empty XML element name")
	}
	return string(data[start:end]), nil
}

func appendNativeXMLChild(data []byte, root *nativeXMLNode, child string) ([]byte, error) {
	if root.Start < 0 || root.End <= root.Start || root.End > int64(len(data)) {
		return nil, fmt.Errorf("invalid XML root boundary")
	}
	raw := data[root.Start:root.End]
	if bytes.HasSuffix(raw, []byte("/>")) {
		qname, err := nativeOpeningQName(data, root)
		if err != nil {
			return nil, err
		}
		out := make([]byte, 0, len(data)+len(child)+len(qname)+1)
		out = append(out, data[:root.End-2]...)
		out = append(out, '>')
		out = append(out, child...)
		out = append(out, "</"+qname+">"...)
		out = append(out, data[root.End:]...)
		return out, nil
	}
	closeOffset := bytes.LastIndex(raw, []byte("</"))
	if closeOffset < 0 {
		return nil, fmt.Errorf("XML root has no closing element")
	}
	insertAt := int(root.Start) + closeOffset
	out := make([]byte, 0, len(data)+len(child))
	out = append(out, data[:insertAt]...)
	out = append(out, child...)
	out = append(out, data[insertAt:]...)
	return out, nil
}

func ensureAllNoteStyles(styles []byte, partName, wordNS string) ([]byte, error) {
	root, err := parseNativeXML(partName, styles)
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: wordNS, Local: "styles"}) {
		return nil, fmt.Errorf("docxpatch: insert note: styles part %q has spoofed or wrong-dialect root", partName)
	}
	if err := rejectNativeNamespaceSpoofing(root, wordNS); err != nil {
		return nil, fmt.Errorf("docxpatch: insert note: styles part %q: %w", partName, err)
	}
	wanted := map[string]string{
		"FootnoteText": "paragraph", "EndnoteText": "paragraph",
		"FootnoteReference": "character", "EndnoteReference": "character",
	}
	seen := map[string]bool{}
	for _, style := range directNativeChildren(root, wordNS, "style") {
		id, hasID := nativeAttr(style, wordNS, "styleId")
		kind, hasKind := nativeAttr(style, wordNS, "type")
		wantKind, targeted := wanted[id]
		if !targeted {
			continue
		}
		if seen[id] || !hasID || !hasKind || kind != wantKind {
			return nil, fmt.Errorf("docxpatch: insert note: style %q is duplicate or has type %q; expected %q", id, kind, wantKind)
		}
		seen[id] = true
		if kind == "character" {
			rPr := directNativeChildren(style, wordNS, "rPr")
			if len(rPr) != 1 {
				return nil, fmt.Errorf("docxpatch: insert note: style %q must have one exact run-properties container", id)
			}
			align := directNativeChildren(rPr[0], wordNS, "vertAlign")
			value, ok := "", false
			if len(align) == 1 && nativeExactLeaf(align[0], xml.Name{Space: wordNS, Local: "val"}) {
				value, ok = nativeAttr(align[0], wordNS, "val")
			}
			if !ok || value != "superscript" {
				return nil, fmt.Errorf("docxpatch: insert note: style %q is not an exact superscript reference style", id)
			}
		}
	}
	result := append([]byte(nil), styles...)
	for _, definition := range []struct{ id, kind, name string }{
		{"FootnoteText", "paragraph", "footnote text"}, {"EndnoteText", "paragraph", "endnote text"},
		{"FootnoteReference", "character", "footnote reference"}, {"EndnoteReference", "character", "endnote reference"},
	} {
		if seen[definition.id] {
			continue
		}
		extra := ""
		if definition.kind == "character" {
			extra = `<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`
		}
		entry := fmt.Sprintf(`<w:style xmlns:w=%q w:type=%q w:styleId=%q><w:name w:val=%q/>%s</w:style>`, wordNS, definition.kind, definition.id, definition.name, extra)
		root, err = parseNativeXML(partName, result)
		if err != nil {
			return nil, err
		}
		result, err = appendNativeXMLChild(result, root, entry)
		if err != nil {
			return nil, err
		}
	}
	return result, nil
}

// InsertNote appends a footnote/endnote reference to the END of paragraph
// afterIndex's existing content, and appends the note's text as a new entry
// in word/footnotes.xml or word/endnotes.xml (created fresh, with the
// conventional separator/continuationSeparator boilerplate, if the document
// has none yet). Returns the new docx bytes and the assigned note id.
func InsertNote(docx []byte, kind NoteKind, afterIndex int, text string) ([]byte, int, error) {
	shape, err := shapeFor(kind)
	if err != nil {
		return nil, 0, err
	}
	if strings.TrimSpace(text) == "" {
		return nil, 0, fmt.Errorf("docxpatch: empty %s text", kind)
	}
	if !utf8.ValidString(text) || !nativeMutationXMLTextValid(text) || nativeMutationUTF16CodeUnits(text) > NativeDOCXMaxTextLength {
		return nil, 0, fmt.Errorf("docxpatch: %s text must be valid XML 1.0 text of at most %d UTF-16 code units", kind, NativeDOCXMaxTextLength)
	}
	pkg, err := openNativeDOCXPackage(docx)
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s source validation: %w", kind, err)
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: signed packages cannot be changed atomically", kind)
	}
	mainPart, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, 0, err
	}
	wordNS, relBase := wordMLTransitional, relBaseTransitional
	if strict {
		wordNS, relBase = wordMLStrict, relBaseStrict
	}
	shape.relType = relBase + string(kind) + "s"
	mainXML := pkg.files[mainPart]
	mainRoot, err := parseNativeXML(mainPart, mainXML)
	if err != nil {
		return nil, 0, err
	}
	if mainRoot.Name != (xml.Name{Space: wordNS, Local: "document"}) {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: main part has spoofed or wrong-dialect root", kind)
	}
	if err := rejectNativeNamespaceSpoofing(mainRoot, wordNS); err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: %w", kind, err)
	}
	bodyNode := firstDirectNativeChild(mainRoot, wordNS, "body")
	if bodyNode == nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: main document has no body", kind)
	}
	paragraphs := nativeDescendants(bodyNode, wordNS, "p")
	if afterIndex < 0 || afterIndex >= len(paragraphs) {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: paragraph %d does not exist (document has %d)", kind, afterIndex, len(paragraphs))
	}

	noteRel, err := nativeWriterSingletonRelationship(pkg, mainPart, shape.relType, relBaseStrict+string(kind)+"s", relBaseTransitional+string(kind)+"s")
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: %w", kind, err)
	}
	hasNotesPart := noteRel != nil
	if hasNotesPart {
		shape.part = noteRel.PartName
	} else {
		mainDir := path.Dir(mainPart)
		if mainDir == "." {
			mainDir = ""
		}
		shape.part = path.Join(mainDir, string(kind)+"s.xml")
		if key, keyErr := nativeDecodedPartKey(shape.part); keyErr != nil {
			return nil, 0, keyErr
		} else if collision := pkg.partByKey[key]; collision != "" {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: unbound canonical note part %q already exists", kind, collision)
		}
	}
	if hasNotesPart && !nativeASCIIEqual(pkg.contentTypes[shape.part], shape.contentType) {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: note part %q has content type %q", kind, shape.part, pkg.contentTypes[shape.part])
	}
	var notesXML []byte
	if hasNotesPart {
		notesXML = pkg.files[shape.part]
	} else {
		notesXML = []byte(freshNotesPart(shape, wordNS))
	}
	notesRoot, nextID, err := inspectNativeNotePart(notesXML, shape, wordNS)
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: %w", kind, err)
	}
	notesXML, err = appendNativeXMLChild(notesXML, notesRoot, noteBodyXML(shape, nextID, text))
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: append note: %w", kind, err)
	}
	newMain, err := appendNativeXMLChild(mainXML, paragraphs[afterIndex], noteRefRunXML(shape, nextID))
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: append reference: %w", kind, err)
	}

	relsPart := pkg.relsPart[mainPart]
	relsXML := pkg.files[relsPart]
	if relsPart == "" {
		relsPart = nativeRelationshipPartName(mainPart)
		if key, keyErr := nativeDecodedPartKey(relsPart); keyErr != nil {
			return nil, 0, keyErr
		} else if collision := pkg.partByKey[key]; collision != "" {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: relationship-part collision at %q", kind, collision)
		}
		relsXML = []byte(emptyRelsXML)
	}
	newRels := append([]byte(nil), relsXML...)
	newContentTypes := append([]byte(nil), pkg.files[contentTypes]...)
	allocatedIDs := append([]nativeRelationship(nil), pkg.rels[mainPart]...)
	if !hasNotesPart {
		id := nextNativeRelationshipID(allocatedIDs)
		newRels, err = appendNativeRelationship(newRels, id, shape.relType, canonicalNativeRelationshipTarget(mainPart, shape.part))
		if err != nil {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: note relationship: %w", kind, err)
		}
		allocatedIDs = append(allocatedIDs, nativeRelationship{ID: id})
		newContentTypes, err = appendNativeContentTypeOverride(newContentTypes, shape.part, shape.contentType)
		if err != nil {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: note content type: %w", kind, err)
		}
	}

	stylesType := relBase + "styles"
	stylesRel, err := nativeWriterSingletonRelationship(pkg, mainPart, stylesType, relBaseStrict+"styles", relBaseTransitional+"styles")
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s: %w", kind, err)
	}
	stylesPart := ""
	var stylesXML []byte
	hasStyles := stylesRel != nil
	if hasStyles {
		stylesPart = stylesRel.PartName
		if !nativeASCIIEqual(pkg.contentTypes[stylesPart], "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml") {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: styles part %q has wrong content type", kind, stylesPart)
		}
		stylesXML = pkg.files[stylesPart]
	} else {
		mainDir := path.Dir(mainPart)
		if mainDir == "." {
			mainDir = ""
		}
		stylesPart = path.Join(mainDir, "styles.xml")
		if key, keyErr := nativeDecodedPartKey(stylesPart); keyErr != nil {
			return nil, 0, keyErr
		} else if collision := pkg.partByKey[key]; collision != "" {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: unbound canonical styles part %q already exists", kind, collision)
		}
		stylesXML = []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="` + wordNS + `"></w:styles>`)
		id := nextNativeRelationshipID(allocatedIDs)
		newRels, err = appendNativeRelationship(newRels, id, stylesType, canonicalNativeRelationshipTarget(mainPart, stylesPart))
		if err != nil {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: styles relationship: %w", kind, err)
		}
		allocatedIDs = append(allocatedIDs, nativeRelationship{ID: id})
		newContentTypes, err = appendNativeContentTypeOverride(newContentTypes, stylesPart, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml")
		if err != nil {
			return nil, 0, fmt.Errorf("docxpatch: insert %s: styles content type: %w", kind, err)
		}
	}
	newStyles, err := ensureAllNoteStyles(stylesXML, stylesPart, wordNS)
	if err != nil {
		return nil, 0, err
	}

	patch := Patch{Replace: map[string][]byte{mainPart: newMain, contentTypes: newContentTypes}, Add: map[string][]byte{}, Delete: map[string]bool{}}
	if hasNotesPart {
		patch.Replace[shape.part] = notesXML
	} else {
		patch.Add[shape.part] = notesXML
	}
	if hasStyles {
		if !bytes.Equal(newStyles, stylesXML) {
			patch.Replace[stylesPart] = newStyles
		}
	} else {
		patch.Add[stylesPart] = newStyles
	}
	if _, existed := pkg.files[relsPart]; existed {
		if !bytes.Equal(newRels, relsXML) {
			patch.Replace[relsPart] = newRels
		}
	} else {
		patch.Add[relsPart] = newRels
	}
	out, err := ApplyPatch(docx, patch)
	if err != nil {
		return nil, 0, err
	}
	verified, err := openNativeDOCXPackage(out)
	if err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s post-write package validation: %w", kind, err)
	}
	verifiedRel, err := nativeWriterSingletonRelationship(verified, mainPart, shape.relType, relBaseStrict+string(kind)+"s", relBaseTransitional+string(kind)+"s")
	if err != nil || verifiedRel == nil || verifiedRel.PartName != shape.part {
		return nil, 0, fmt.Errorf("docxpatch: insert %s post-write relationship validation failed", kind)
	}
	if _, _, err := inspectNativeNotePart(verified.files[shape.part], shape, wordNS); err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s post-write note validation: %w", kind, err)
	}
	if _, err := ensureAllNoteStyles(verified.files[stylesPart], stylesPart, wordNS); err != nil {
		return nil, 0, fmt.Errorf("docxpatch: insert %s post-write style validation: %w", kind, err)
	}
	return out, nextID, nil
}

func nativeRelationshipPartName(owner string) string {
	dir, base := path.Dir(owner), path.Base(owner)
	if dir == "." {
		return path.Join("_rels", base+".rels")
	}
	return path.Join(dir, "_rels", base+".rels")
}

func canonicalNativeRelationshipTarget(owner, target string) string {
	ownerDir := path.Dir(owner)
	if ownerDir == "." {
		ownerDir = ""
	}
	from := []string{}
	if ownerDir != "" {
		from = strings.Split(ownerDir, "/")
	}
	to := strings.Split(target, "/")
	common := 0
	for common < len(from) && common < len(to)-1 && from[common] == to[common] {
		common++
	}
	parts := make([]string, 0, len(from)-common+len(to)-common)
	for index := common; index < len(from); index++ {
		parts = append(parts, "..")
	}
	parts = append(parts, to[common:]...)
	return strings.Join(parts, "/")
}

func nativeWriterSingletonRelationship(pkg *nativePackage, owner, want string, dialectTypes ...string) (*nativeRelationship, error) {
	var match *nativeRelationship
	for index := range pkg.rels[owner] {
		rel := &pkg.rels[owner][index]
		isDialectType := false
		for _, candidate := range dialectTypes {
			isDialectType = isDialectType || rel.Type == candidate
		}
		if isDialectType && rel.Type != want {
			return nil, fmt.Errorf("relationship %q uses the wrong Strict/Transitional type %q", rel.ID, rel.Type)
		}
		if rel.Type != want {
			continue
		}
		if match != nil {
			return nil, fmt.Errorf("multiple %q relationships", want)
		}
		if rel.External || rel.PartName == "" {
			return nil, fmt.Errorf("relationship %q must be internal", rel.ID)
		}
		match = rel
	}
	if match == nil {
		return nil, nil
	}
	relsPart := pkg.relsPart[owner]
	root, err := parseNativeXML(relsPart, pkg.files[relsPart])
	if err != nil {
		return nil, err
	}
	found := 0
	for _, node := range root.Children {
		id, _ := nativeUnqualifiedAttr(node, "Id")
		if id != match.ID {
			continue
		}
		found++
		typeURI, okType := nativeUnqualifiedAttr(node, "Type")
		target, okTarget := nativeUnqualifiedAttr(node, "Target")
		_, hasMode := nativeUnqualifiedAttr(node, "TargetMode")
		if !okType || typeURI != want || !okTarget || hasMode || target != canonicalNativeRelationshipTarget(owner, match.PartName) {
			return nil, fmt.Errorf("relationship %q must have exact Type, absent TargetMode, and canonical Target", match.ID)
		}
	}
	if found != 1 {
		return nil, fmt.Errorf("relationship %q is not an exact singleton", match.ID)
	}
	return match, nil
}

func nextNativeRelationshipID(relationships []nativeRelationship) string {
	seen := make(map[string]bool, len(relationships))
	for _, rel := range relationships {
		seen[rel.ID] = true
	}
	for index := 1; ; index++ {
		candidate := "rId" + strconv.Itoa(index)
		if !seen[candidate] {
			return candidate
		}
	}
}

func appendNativeRelationship(data []byte, id, relType, target string) ([]byte, error) {
	root, err := parseNativeXML("relationships", data)
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: opcRelationshipsNS, Local: "Relationships"}) || !nativeExactContainer(root) {
		return nil, fmt.Errorf("relationship part has invalid root")
	}
	for _, node := range root.Children {
		existing, _ := nativeUnqualifiedAttr(node, "Id")
		if existing == id {
			return nil, fmt.Errorf("duplicate relationship id %q", id)
		}
	}
	entry := fmt.Sprintf(`<Relationship xmlns=%q Id=%q Type=%q Target=%q/>`, opcRelationshipsNS, id, relType, target)
	return appendNativeXMLChild(data, root, entry)
}

func appendNativeContentTypeOverride(data []byte, partName, contentType string) ([]byte, error) {
	root, err := parseNativeXML(contentTypes, data)
	if err != nil {
		return nil, err
	}
	if root.Name != (xml.Name{Space: opcContentTypesNS, Local: "Types"}) || !nativeExactContainer(root) {
		return nil, fmt.Errorf("content-types part has invalid root")
	}
	wantKey, err := nativeDecodedPartKey(partName)
	if err != nil {
		return nil, err
	}
	for _, node := range root.Children {
		if node.Name != (xml.Name{Space: opcContentTypesNS, Local: "Override"}) {
			continue
		}
		existing, _ := nativeUnqualifiedAttr(node, "PartName")
		if existing == "" {
			continue
		}
		key, keyErr := nativeDecodedPartKey(strings.TrimPrefix(existing, "/"))
		if keyErr != nil {
			return nil, keyErr
		}
		if key == wantKey {
			return nil, fmt.Errorf("duplicate or spoofed content-type PartName for %q", partName)
		}
	}
	entry := fmt.Sprintf(`<Override xmlns=%q PartName=%q ContentType=%q/>`, opcContentTypesNS, "/"+partName, contentType)
	return appendNativeXMLChild(data, root, entry)
}

func inspectNativeNotePart(data []byte, shape noteShape, wordNS string) (*nativeXMLNode, int, error) {
	root, err := parseNativeXML(shape.part, data)
	if err != nil {
		return nil, 0, err
	}
	wantRoot := strings.TrimPrefix(shape.rootEl, "w:")
	wantNote := strings.TrimPrefix(shape.noteEl, "w:")
	if root.Name != (xml.Name{Space: wordNS, Local: wantRoot}) || !nativeExactContainer(root) {
		return nil, 0, fmt.Errorf("note part %q has spoofed, wrong-dialect, or non-exact root", shape.part)
	}
	if err := rejectNativeNamespaceSpoofing(root, wordNS); err != nil {
		return nil, 0, err
	}
	seen := map[string]bool{}
	maxID := 0
	for _, node := range root.Children {
		if node.Name != (xml.Name{Space: wordNS, Local: wantNote}) || !nativeExactContainer(node, xml.Name{Space: wordNS, Local: "id"}, xml.Name{Space: wordNS, Local: "type"}) {
			return nil, 0, fmt.Errorf("note part %q contains non-exact note markup at %s", shape.part, node.Path)
		}
		id, hasID := nativeAttr(node, wordNS, "id")
		noteType, hasType := nativeAttr(node, wordNS, "type")
		if !hasID || seen[id] {
			return nil, 0, fmt.Errorf("note part %q contains missing or duplicate note id %q", shape.part, id)
		}
		seen[id] = true
		switch id {
		case "-1":
			if !hasType || noteType != "separator" || !nativeExactNoteSentinel(node, wordNS, "separator") {
				return nil, 0, fmt.Errorf("note part %q has malformed -1 separator sentinel", shape.part)
			}
		case "0":
			if !hasType || noteType != "continuationSeparator" || !nativeExactNoteSentinel(node, wordNS, "continuation-separator") {
				return nil, 0, fmt.Errorf("note part %q has malformed 0 continuation sentinel", shape.part)
			}
		default:
			value, parseErr := strconv.ParseInt(id, 10, 32)
			if hasType || parseErr != nil || value <= 0 || strconv.FormatInt(value, 10) != id {
				return nil, 0, fmt.Errorf("note part %q has non-canonical positive note id %q", shape.part, id)
			}
			if int(value) > maxID {
				maxID = int(value)
			}
		}
	}
	if !seen["-1"] || !seen["0"] {
		return nil, 0, fmt.Errorf("note part %q must contain exact -1 and 0 sentinels", shape.part)
	}
	if maxID == int(^uint32(0)>>1) {
		return nil, 0, fmt.Errorf("note id space is exhausted")
	}
	return root, maxID + 1, nil
}

// overridePartWith is shared by the older chart/watermark writers. Its
// implementation is bounded and namespace-aware so a substring cannot spoof
// an existing OPC override.
func overridePartWith(contentTypesXML, partName, contentType string) (string, error) {
	root, err := parseNativeXML(contentTypes, []byte(contentTypesXML))
	if err != nil {
		return "", err
	}
	wantKey, err := nativeDecodedPartKey(strings.TrimPrefix(partName, "/"))
	if err != nil {
		return "", err
	}
	for _, node := range root.Children {
		if node.Name != (xml.Name{Space: opcContentTypesNS, Local: "Override"}) {
			continue
		}
		existing, _ := nativeUnqualifiedAttr(node, "PartName")
		key, keyErr := nativeDecodedPartKey(strings.TrimPrefix(existing, "/"))
		if keyErr != nil {
			return "", keyErr
		}
		if key != wantKey {
			continue
		}
		got, ok := nativeUnqualifiedAttr(node, "ContentType")
		if !ok || !nativeASCIIEqual(got, contentType) {
			return "", fmt.Errorf("docxpatch: content-type override for %q conflicts with %q", partName, got)
		}
		return contentTypesXML, nil
	}
	updated, err := appendNativeContentTypeOverride([]byte(contentTypesXML), strings.TrimPrefix(partName, "/"), contentType)
	return string(updated), err
}
