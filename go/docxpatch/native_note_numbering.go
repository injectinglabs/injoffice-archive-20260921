package docxpatch

import (
	"encoding/xml"
	"strings"
)

// Note numbering, ECMA-376 17.11.17/17.11.18. A section's w:footnotePr or
// w:endnotePr states the counter alphabet Word prints for that kind's anchors
// and labels; without it Word prints decimal footnote numbers and lowerRoman
// endnote numbers only when the package says so, so an absent w:numFmt keeps
// this tier's decimal counter.
//
// Only the counter alphabet is modeled here. The rest of CT_FtnProps and
// CT_EdnProps -- placement, restart, custom marks and the numbering start --
// still leaves the property unmodeled, because each of those moves where a
// note lands or which value it starts from, which this tier has no input for.
//
// The formatting itself is nativeFormatNumberingCounter, the same
// implementation the list-numbering resolver uses. There is exactly one
// counter implementation in this codebase, and the layout tier reads its
// output out of the Labels table by counter value rather than re-deriving it.
const nativeNoteNumberingMaxLabels = 4096

// nativeNoteNumberFormat reads an exact `<w:xxxnotePr><w:numFmt w:val="..."/>`
// pair. Anything else -- extra attributes, extra children, a missing or
// unformattable w:val -- states note properties this tier does not model.
func nativeNoteNumberFormat(node *nativeXMLNode, wordNS string) (string, bool) {
	if !nativeExactContainer(node) || len(node.Children) != 1 {
		return "", false
	}
	child := node.Children[0]
	if child.Name != (xml.Name{Space: wordNS, Local: "numFmt"}) || !nativeExactLeaf(child, xml.Name{Space: wordNS, Local: "val"}) {
		return "", false
	}
	format, ok := nativeAttr(child, wordNS, "val")
	if !ok || format == "" {
		return "", false
	}
	// Reject any alphabet whose first counter the shared formatter cannot
	// produce, including bullet, none and the chicago symbol sequence.
	if _, ok := nativeFormatNumberingCounter(1, format); !ok {
		return "", false
	}
	return format, true
}

// recordNoteNumberFormat accumulates one section's statement for a note kind.
// Sections that disagree leave the kind unmodeled and disclosed, so the
// document never paints one section's alphabet over another's.
func (extractor *nativeExtractor) recordNoteNumberFormat(kind, format string) bool {
	if extractor.noteNumberFormats == nil {
		extractor.noteNumberFormats = map[string]string{}
		extractor.noteNumberConflicts = map[string]bool{}
	}
	if prior, seen := extractor.noteNumberFormats[kind]; seen && prior != format {
		extractor.noteNumberConflicts[kind] = true
		return false
	}
	extractor.noteNumberFormats[kind] = format
	return true
}

// nativeNoteNumberingRecords builds the per-kind label table the layout tier
// indexes by counter value. Labels[i] is the label for counter value i+1, and
// the table is exactly as long as the document has content notes of that kind,
// which is the largest counter value the tier can assign.
func (extractor *nativeExtractor) nativeNoteNumberingRecords() []NativeNoteNumberingV1 {
	records := []NativeNoteNumberingV1{}
	for _, kind := range []string{"endnote", "footnote"} {
		format, ok := extractor.noteNumberFormats[kind]
		if !ok || extractor.noteNumberConflicts[kind] {
			continue
		}
		count := 0
		for _, story := range extractor.notes {
			if story.Kind == kind && (story.NoteRole == "" || story.NoteRole == "content") {
				count++
			}
		}
		if count == 0 || count > nativeNoteNumberingMaxLabels {
			continue
		}
		labels := make([]string, 0, count)
		for value := 1; value <= count; value++ {
			label, ok := nativeFormatNumberingCounter(value, format)
			if !ok || label == "" || strings.ContainsAny(label, "\x00\r\n") {
				labels = nil
				break
			}
			labels = append(labels, label)
		}
		if labels == nil {
			continue
		}
		records = append(records, NativeNoteNumberingV1{Kind: kind, Format: format, Labels: labels})
	}
	if len(records) == 0 {
		return nil
	}
	return records
}
