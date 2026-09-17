package docxpatch

import "encoding/xml"

// Word records a dragged table row, or a table row moved between two tables, as
// a tracked move. Two kinds of markup carry it, and neither one is content.
//
// The range endpoints (w:moveFromRangeStart/End, w:moveToRangeStart/End, and
// the bookmark and permission endpoints that sit beside them) are empty
// elements that delimit a span. Word paints nothing for them. Between two
// w:tr elements they were reaching the generic "table content outside direct
// rows" refusal, which exists for markup that can carry visible content, such
// as a stray paragraph or a wrapper around rows. A delimiter carries none, so
// it is recorded under its own code and the rows around it stay paintable.
//
// The mark annotations (w:ins, w:del, w:moveFrom, w:moveTo inside a
// w:pPr/w:rPr) state which revision the paragraph mark belongs to. They are
// CT_TrackChange metadata, not formatting: nothing in them resolves to a font,
// a size, a colour or an advance. They were falling through to "this run
// property is preserved and not guessed", which says the mark's formatting is
// unresolved when in fact nothing about the mark's formatting is stated.
//
// Both keep an explicit diagnostic, because neither the revision-marked display
// Word shows by default nor the final display it shows with markup off is
// modelled here. Only the approximate tier, which already omits the w:ins and
// w:del run content these markers accompany, treats them as non-blocking.

// nativeTrackedMoveRangeMarkers are the content-free span endpoints that may
// appear between a table's rows.
var nativeTrackedMoveRangeMarkers = map[string]bool{
	"bookmarkStart": true, "bookmarkEnd": true,
	"moveFromRangeStart": true, "moveFromRangeEnd": true,
	"moveToRangeStart": true, "moveToRangeEnd": true,
	"commentRangeStart": true, "commentRangeEnd": true,
	"permStart": true, "permEnd": true,
	"proofErr": true,
}

// nativeTrackedMarkRevisions are the CT_TrackChange annotations a paragraph
// mark's run properties may carry.
var nativeTrackedMarkRevisions = map[string]bool{
	"ins": true, "del": true, "moveFrom": true, "moveTo": true,
}

// nativeNonVisualRangeMarker reports whether node is one of the content-free
// span endpoints above, stated exactly: a word-namespace empty element whose
// attributes are all in the word namespace. Anything nested inside it, or any
// foreign attribute on it, may carry meaning this layer has not read, so it is
// not one of these markers.
func nativeNonVisualRangeMarker(node *nativeXMLNode, wordNS string) bool {
	if node.Name.Space != wordNS || !nativeTrackedMoveRangeMarkers[node.Name.Local] {
		return false
	}
	if len(node.Children) != 0 || !nativeXMLWhitespaceOnly(node.Text) {
		return false
	}
	for _, attr := range node.Attrs {
		if nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		if attr.Name.Space != wordNS {
			return false
		}
	}
	return true
}

// nativeTrackedMarkRevision reports whether node is a paragraph-mark revision
// annotation stated exactly: a word-namespace empty element carrying only the
// CT_TrackChange attributes.
func nativeTrackedMarkRevision(node *nativeXMLNode, wordNS string) bool {
	if node.Name.Space != wordNS || !nativeTrackedMarkRevisions[node.Name.Local] {
		return false
	}
	return nativeExactLeaf(node,
		xml.Name{Space: wordNS, Local: "id"},
		xml.Name{Space: wordNS, Local: "author"},
		xml.Name{Space: wordNS, Local: "date"},
	)
}
