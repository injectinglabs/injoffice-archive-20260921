package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

const nativeIndexEntryMarkup = `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
	`<w:r><w:instrText xml:space="preserve"> XE "</w:instrText></w:r>` +
	`<w:proofErr w:type="spellStart"/>` +
	`<w:r w:rsidRPr="003224E4"><w:rPr><w:b/></w:rPr><w:instrText>galleries:aa</w:instrText></w:r>` +
	`<w:proofErr w:type="spellEnd"/>` +
	`<w:r><w:instrText xml:space="preserve">" </w:instrText></w:r>` +
	`<w:r><w:fldChar w:fldCharType="end"/></w:r>`

func nativeIndexEntryDocument(t *testing.T, ns, markup string) []byte {
	t.Helper()
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"/>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p><w:r><w:t xml:space="preserve">galleries</w:t></w:r>` + markup + `<w:r><w:t xml:space="preserve"> to insert</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	if ns == wordMLStrict {
		for key, value := range parts {
			parts[key] = strings.ReplaceAll(value, relBaseTransitional, relBaseStrict)
		}
	}
	return buildNativeDOCX(t, nativeEntries(parts))
}

func nativeFieldSemanticsRecords(doc *NativeDocumentV1) int {
	count := 0
	for _, entry := range doc.Unsupported {
		if entry.Code == "FIELD_SEMANTICS" {
			count++
		}
	}
	return count
}

func nativeRunProjection(doc *NativeDocumentV1) []string {
	projection := []string{}
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		text := ""
		if run.Text != nil {
			text = *run.Text
		}
		projection = append(projection, run.Kind+"|"+run.Control+"|"+text)
	}
	return projection
}

// An XE field has no result, so admitting it must leave the paragraph's painted
// content exactly as it is without the field markup at all.
func TestNativeIndexEntryFieldPaintsNothing(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		admitted, err := ExtractNativeDocumentV1(nativeIndexEntryDocument(t, ns, nativeIndexEntryMarkup))
		if err != nil {
			t.Fatal(err)
		}
		stripped, err := ExtractNativeDocumentV1(nativeIndexEntryDocument(t, ns, ""))
		if err != nil {
			t.Fatal(err)
		}
		if got, want := nativeRunProjection(admitted), nativeRunProjection(stripped); strings.Join(got, "\x00") != strings.Join(want, "\x00") {
			t.Fatalf("%s runs with the XE field = %#v, stripped = %#v", ns, got, want)
		}
		if count := nativeFieldSemanticsRecords(admitted); count != 0 {
			t.Fatalf("%s XE field recorded %d FIELD_SEMANTICS diagnostics: %#v", ns, count, admitted.Unsupported)
		}
		// The field's source stays preserve-only: the paragraph carrying it is
		// never editable, exactly as a closed empty bookmark leaves it.
		if mode := admitted.Body.Blocks[0].Paragraph.EditPolicy.Mode; mode != "read-only" {
			t.Fatalf("%s paragraph edit mode = %q, want read-only", ns, mode)
		}
	}
}

func TestNativeIndexEntryFieldSourcePreserved(t *testing.T) {
	data := nativeIndexEntryDocument(t, wordMLTransitional, nativeIndexEntryMarkup)
	before := bytes.Clone(data)
	if _, err := ExtractNativeDocumentV1(data); err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(before, data) {
		t.Fatal("source changed")
	}
}

// Everything that paints keeps refusing. The instruction kinds below all state
// a result Word paints, and the structural variants either give the XE field a
// result of its own or put painted content inside its boundaries.
func TestNativeIndexEntryFieldRefusesEverythingThatPaints(t *testing.T) {
	cases := map[string]string{
		"page":      strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` PAGE `, 1),
		"numpages":  strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` NUMPAGES `, 1),
		"ref":       strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` REF `, 1),
		"pageref":   strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` PAGEREF `, 1),
		"styleref":  strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` STYLEREF `, 1),
		"seq":       strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` SEQ `, 1),
		"toc":       strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` TOC `, 1),
		"index":     strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` INDEX `, 1),
		"hyperlink": strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` HYPERLINK `, 1),
		// XE is the keyword, not a prefix of one, and not an argument.
		"prefix":   strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` XEROX "`, 1),
		"argument": strings.Replace(nativeIndexEntryMarkup, ` XE "`, ` QUOTE XE "`, 1),
		// A separate gives the field a cached result, which is painted content.
		"result": strings.Replace(nativeIndexEntryMarkup, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
			`<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`, 1),
		// Text inside the boundaries is painted, whatever the instruction says.
		"visible": strings.Replace(nativeIndexEntryMarkup, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
			`<w:r><w:t>visible</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`, 1),
		// A nested field is another field, whose own result is unread here.
		"nested": strings.Replace(nativeIndexEntryMarkup, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`,
			`<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`, 1),
		// An unterminated field states no boundary at all.
		"unterminated": strings.Replace(nativeIndexEntryMarkup, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`, "", 1),
		// Markup this layer has not read may carry meaning; the run's own
		// attributes and the boundary character's stay exact.
		"foreign-attribute": strings.Replace(nativeIndexEntryMarkup, `<w:fldChar w:fldCharType="begin"/>`, `<w:fldChar w:fldCharType="begin" w:dirty="true"/>`, 1),
		"foreign-run":       strings.Replace(nativeIndexEntryMarkup, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`, `<w:r extra="x"><w:fldChar w:fldCharType="end"/></w:r>`, 1),
		"foreign-sibling":   strings.Replace(nativeIndexEntryMarkup, `<w:r><w:fldChar w:fldCharType="end"/></w:r>`, `<w:customXml/><w:r><w:fldChar w:fldCharType="end"/></w:r>`, 1),
		"nested-instrText":  strings.Replace(nativeIndexEntryMarkup, `<w:instrText>galleries:aa</w:instrText>`, `<w:instrText><w:t>galleries:aa</w:t></w:instrText>`, 1),
	}
	for name, markup := range cases {
		doc, err := ExtractNativeDocumentV1(nativeIndexEntryDocument(t, wordMLTransitional, markup))
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		if nativeFieldSemanticsRecords(doc) == 0 {
			t.Fatalf("%s: field admitted with no FIELD_SEMANTICS record: %#v", name, doc.Unsupported)
		}
	}
}
