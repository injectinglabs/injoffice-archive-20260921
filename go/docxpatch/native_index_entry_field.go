package docxpatch

import (
	"encoding/xml"
	"strings"
)

// An XE field marks a word for inclusion in an index that some INDEX field
// builds elsewhere. ECMA-376 17.16.5.34 gives it no result: the field states
// the entry text and the switches that decide how the INDEX output spells it,
// and Word paints nothing where the field itself sits. The markup carries that
// out structurally too — an XE field is written begin/instruction/end with no
// w:fldChar separate, so the run sequence holds no result to paint and the
// document says so before anything is resolved.
//
// Microsoft Word 16.112.4's own PDF export of alphabeticalIndex_MultipleColumns
// and IndexFieldFlagF confirms it. Both documents tag words inside ordinary
// sentences, and the exported content stream runs those sentences together with
// no glyph and no advance at any tagged site: "You can use these galleries to
// insert tables, headers, footers" prints contiguously although XE fields sit
// between "galleries" and " to insert" and between " to insert" and " tables".
//
// So an XE field is authoring metadata in the sense #336 accepted w:name: it
// selects nothing that paints and cannot move a line or a page. The field's
// source is still preserved verbatim and the containing paragraph stays
// non-editable, exactly as a closed empty bookmark does.
//
// Every other instruction keeps refusing. PAGE, NUMPAGES, REF, PAGEREF,
// STYLEREF, SEQ, TOC and INDEX all produce a result that Word paints, and an
// INDEX field is the opposite case in the very same documents: it generates the
// index body, which is real painted content. Only a keyword with no result at
// all qualifies here.
const nativeIndexEntryFieldKeyword = "XE"

// nativeIndexEntryMaxInstruction bounds the concatenated instruction text a
// single marker field may carry, so a pathological run sequence cannot be
// walked without limit.
const nativeIndexEntryMaxInstruction = 4096

// nativeIndexEntryFieldSpan reads one XE marker field from the front of
// sequence and returns how many siblings the field spans. The span may contain
// the content-free range endpoints Word writes beside a field — proofing,
// bookmark, comment and permission markers — which keep their own treatment.
func nativeIndexEntryFieldSpan(sequence []*nativeXMLNode, wordNS string) (int, bool) {
	if len(sequence) < 2 || !nativeIndexEntryFieldRun(sequence[0], wordNS, "begin") {
		return 0, false
	}
	instruction := ""
	for index := 1; index < len(sequence); index++ {
		child := sequence[index]
		if child.Name != (xml.Name{Space: wordNS, Local: "r"}) {
			// The content-free endpoints Word writes beside a field keep their
			// own treatment; anything else inside the boundaries is unread.
			if nativeNonVisualRangeMarker(child, wordNS) {
				continue
			}
			return 0, false
		}
		// Only the matching end closes the field. A nested begin or a separate
		// would mean a result exists, or another field is inside this one, and
		// neither is this shape.
		if nativeIndexEntryFieldRun(child, wordNS, "end") {
			if !nativeIndexEntryInstruction(instruction) {
				return 0, false
			}
			return index + 1, true
		}
		text, ok := nativeIndexEntryFieldRunInstruction(child, wordNS)
		if !ok || len(instruction)+len(text) > nativeIndexEntryMaxInstruction {
			return 0, false
		}
		instruction += text
	}
	return 0, false
}

// nativeIndexEntryFieldContent returns the single element a field run states
// beside its optional run properties.
func nativeIndexEntryFieldContent(run *nativeXMLNode, wordNS string) *nativeXMLNode {
	if run.Name != (xml.Name{Space: wordNS, Local: "r"}) || !nativeExactRevisionContainer(run, wordNS, "rsidR", "rsidRPr", "rsidDel") {
		return nil
	}
	if len(directNativeChildren(run, wordNS, "rPr")) > 1 {
		return nil
	}
	var content *nativeXMLNode
	for _, child := range run.Children {
		if child.Name == (xml.Name{Space: wordNS, Local: "rPr"}) {
			continue
		}
		if content != nil {
			return nil
		}
		content = child
	}
	return content
}

// nativeIndexEntryFieldRun reports whether run states exactly one field
// boundary character of the named kind and nothing else.
func nativeIndexEntryFieldRun(run *nativeXMLNode, wordNS, kind string) bool {
	content := nativeIndexEntryFieldContent(run, wordNS)
	if content == nil || content.Name != (xml.Name{Space: wordNS, Local: "fldChar"}) {
		return false
	}
	actual, present := nativeAttr(content, wordNS, "fldCharType")
	return present && actual == kind && nativeExactLeaf(content, xml.Name{Space: wordNS, Local: "fldCharType"})
}

// nativeIndexEntryFieldRunInstruction returns the instruction text run states.
// An exact w:instrText is text only, with no attribute but the xml:space it
// needs to keep the spaces that separate the instruction's tokens.
func nativeIndexEntryFieldRunInstruction(run *nativeXMLNode, wordNS string) (string, bool) {
	content := nativeIndexEntryFieldContent(run, wordNS)
	if content == nil || content.Name != (xml.Name{Space: wordNS, Local: "instrText"}) || len(content.Children) != 0 {
		return "", false
	}
	for _, attr := range content.Attrs {
		if attr.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || (attr.Value != "preserve" && attr.Value != "default") {
			return "", false
		}
	}
	return content.Text, true
}

// nativeIndexEntryInstruction reports whether the concatenated instruction
// names the index-entry keyword. Its switches are not read: every XE switch
// (\b \f \i \r \t \y \z) decides how the INDEX output spells the entry, and
// none of them gives the XE field itself a result to paint.
func nativeIndexEntryInstruction(instruction string) bool {
	if len(instruction) > nativeIndexEntryMaxInstruction {
		return false
	}
	// The keyword is the whole first token, never the head of a longer one.
	tokens := strings.FieldsFunc(instruction, func(r rune) bool { return r == ' ' || r == '\t' || r == '\r' || r == '\n' })
	return len(tokens) > 0 && tokens[0] == nativeIndexEntryFieldKeyword
}
