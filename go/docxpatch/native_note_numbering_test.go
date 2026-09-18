package docxpatch

import (
	"strings"
	"testing"
)

// The section-level note properties this extractor reads, with the two content
// endnotes the shared fixture already carries (ids 2 and 3 after the second is
// appended below), so the label table has a value to format for each of them.
func noteNumberingParts(t *testing.T, sectionPr ...string) map[string]string {
	t.Helper()
	parts := transitionalNativeParts()
	parts["Custom/Notes/End.XML"] = `<w:endnotes xmlns:w="` + testW + `"><w:endnote w:id="2"><w:p><w:r><w:t>Endnote β</w:t></w:r></w:p></w:endnote><w:endnote w:id="3"><w:p><w:r><w:t>Endnote γ</w:t></w:r></w:p></w:endnote></w:endnotes>`
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:r><w:endnoteReference w:id="2"/></w:r>`, `<w:r><w:endnoteReference w:id="2"/></w:r><w:r><w:endnoteReference w:id="3"/></w:r>`, 1)
	for index, markup := range sectionPr {
		old := `<w:sectPr><w:type w:val="continuous"/>`
		if index == 1 {
			old = `<w:sectPr><w:type w:val="nextPage"/>`
		}
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], old, old+markup, 1)
	}
	return parts
}

func extractNoteNumbering(t *testing.T, parts map[string]string) *NativeDocumentV1 {
	t.Helper()
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatalf("extract: %v", err)
	}
	return doc
}

func noteNumberingSectionDiagnostics(doc *NativeDocumentV1) []string {
	messages := []string{}
	for _, entry := range doc.Unsupported {
		if entry.Code == "UNMODELED_SECTION_PROPERTY" {
			messages = append(messages, entry.Message)
		}
	}
	return messages
}

func TestNativeNoteNumberingReadsSectionNumberFormat(t *testing.T) {
	doc := extractNoteNumbering(t, noteNumberingParts(t, `<w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr>`))
	if len(doc.NoteNumbering) != 1 {
		t.Fatalf("note_numbering = %#v; want one endnote record", doc.NoteNumbering)
	}
	record := doc.NoteNumbering[0]
	if record.Kind != "endnote" || record.Format != "lowerRoman" || strings.Join(record.Labels, ",") != "i,ii" {
		t.Fatalf("note_numbering = %#v; want endnote lowerRoman [i ii]", record)
	}
	for _, message := range noteNumberingSectionDiagnostics(doc) {
		if strings.Contains(message, "numbering-format") {
			t.Fatalf("a modeled numbering format must not also be preserved as unmodeled: %q", message)
		}
	}
}

func TestNativeNoteNumberingKeepsDecimalUnmodeledMarkupRefused(t *testing.T) {
	for name, markup := range map[string]string{
		// A note property other than the counter alphabet still states placement
		// or restart semantics with no input here.
		"extra child":    `<w:endnotePr><w:numFmt w:val="lowerRoman"/><w:numRestart w:val="eachPage"/></w:endnotePr>`,
		"attribute":      `<w:endnotePr w:pos="sectEnd"><w:numFmt w:val="lowerRoman"/></w:endnotePr>`,
		"numFmt extra":   `<w:endnotePr><w:numFmt w:val="custom" w:format="x"/></w:endnotePr>`,
		"empty":          `<w:endnotePr/>`,
		"unknown format": `<w:endnotePr><w:numFmt w:val="chicago"/></w:endnotePr>`,
		"bullet":         `<w:endnotePr><w:numFmt w:val="bullet"/></w:endnotePr>`,
	} {
		doc := extractNoteNumbering(t, noteNumberingParts(t, markup))
		if len(doc.NoteNumbering) != 0 {
			t.Fatalf("%s: note_numbering = %#v; want no record", name, doc.NoteNumbering)
		}
		if len(noteNumberingSectionDiagnostics(doc)) == 0 {
			t.Fatalf("%s: unmodeled note properties must stay disclosed", name)
		}
	}
}

func TestNativeNoteNumberingRefusesConflictingSections(t *testing.T) {
	doc := extractNoteNumbering(t, noteNumberingParts(t,
		`<w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr>`,
		`<w:endnotePr><w:numFmt w:val="upperRoman"/></w:endnotePr>`))
	if len(doc.NoteNumbering) != 0 {
		t.Fatalf("note_numbering = %#v; want no record when sections disagree", doc.NoteNumbering)
	}
	found := false
	for _, message := range noteNumberingSectionDiagnostics(doc) {
		found = found || strings.Contains(message, "more than one numbering format")
	}
	if !found {
		t.Fatalf("conflicting numbering formats must be disclosed: %v", noteNumberingSectionDiagnostics(doc))
	}
}

func TestNativeNoteNumberingContractRejectsMalformedRecords(t *testing.T) {
	base := extractNoteNumbering(t, noteNumberingParts(t, `<w:endnotePr><w:numFmt w:val="lowerRoman"/></w:endnotePr>`))
	if len(base.NoteNumbering) != 1 || len(base.NoteNumbering[0].Labels) != 2 {
		t.Fatalf("note_numbering = %#v; want one endnote record with two labels", base.NoteNumbering)
	}
	for name, mutate := range map[string]func(doc *NativeDocumentV1){
		"unknown kind":    func(doc *NativeDocumentV1) { doc.NoteNumbering[0].Kind = "comment" },
		"missing format":  func(doc *NativeDocumentV1) { doc.NoteNumbering[0].Format = "" },
		"empty labels":    func(doc *NativeDocumentV1) { doc.NoteNumbering[0].Labels = nil },
		"blank label":     func(doc *NativeDocumentV1) { doc.NoteNumbering[0].Labels[1] = "" },
		"duplicate kinds": func(doc *NativeDocumentV1) { doc.NoteNumbering = append(doc.NoteNumbering, doc.NoteNumbering[0]) },
	} {
		clone := *base
		clone.NoteNumbering = append([]NativeNoteNumberingV1(nil), NativeNoteNumberingV1{Kind: base.NoteNumbering[0].Kind, Format: base.NoteNumbering[0].Format, Labels: append([]string(nil), base.NoteNumbering[0].Labels...)})
		mutate(&clone)
		if issues := ValidateNativeDocumentV1(&clone); len(issues) == 0 {
			t.Fatalf("%s: contract must reject the mutated numbering record", name)
		}
	}
}
