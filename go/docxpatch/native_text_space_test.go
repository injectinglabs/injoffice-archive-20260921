package docxpatch

import (
	"strings"
	"testing"
)

func TestNativeTextReplacementPreservesEdgeSpace(t *testing.T) {
	for _, opening := range []string{`<w:t>`, `<w:t xml:space='preserve'>`} {
		for _, text := range []string{" leading", "trailing ", " both ", "", "plain"} {
			t.Run(opening+text, func(t *testing.T) {
				original := opening + "Before</w:t>"
				main := nativeMutationMain(`<w:p><w:r><w:rPr><w:b/></w:rPr>` + original + `</w:r></w:p>`)
				parts := nativeMutationParts(main)
				source := buildNativeDOCX(t, nativeEntries(parts))
				doc, err := ExtractNativeDocumentV1(source)
				if err != nil {
					t.Fatal(err)
				}
				run := doc.Body.Blocks[0].Paragraph.Runs[0]
				result, err := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: text}})
				if err != nil {
					t.Fatal(err)
				}
				wantOpening := opening
				if opening == `<w:t>` && nativeMutationNeedsPreservedSpace(text) {
					wantOpening = `<w:t xml:space="preserve">`
				}
				want := strings.Replace(main, original, wantOpening+text+`</w:t>`, 1)
				if got := string(readNativeZipPart(t, result.Package, "word/document.xml")); got != want {
					t.Fatalf("XML splice:\n%s\nwant:\n%s", got, want)
				}
				for part := range parts {
					if part != "word/document.xml" {
						assertNativeRawPartPreserved(t, source, result.Package, part)
					}
				}
				reread, err := ExtractNativeDocumentV1(result.Package)
				if err != nil {
					t.Fatal(err)
				}
				got := reread.Body.Blocks[0].Paragraph.Runs[0].Text
				if got == nil || *got != text {
					t.Fatalf("re-extracted text = %v, want %q", got, text)
				}
			})
		}
	}
}

func TestNativeTextReplacementSpaceLexicalForms(t *testing.T) {
	for _, raw := range []string{`<w:t/>`, `<w:t />`, `<w:t data-note="xml:space='preserve'">old</w:t>`} {
		result, err := replaceNativeTextElement([]byte(raw), " new ")
		if err != nil {
			t.Fatal(err)
		}
		if !strings.Contains(string(result), ` xml:space="preserve"> new </w:t>`) {
			t.Fatalf("missing actual preservation attribute: %s", result)
		}
	}
}
