package docxpatch

import (
	"strings"
	"testing"
)

// A style that supplies a character indent together with Word's own twip
// conversion, exactly as Word writes an East-Asian list style.
const characterIndentStyles = `<w:styles xmlns:w="` + wordMLTransitional + `">` +
	`<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>` +
	`<w:style w:type="paragraph" w:styleId="ListParagraph"><w:pPr><w:ind w:leftChars="200" w:left="480"/></w:pPr></w:style>` +
	`</w:styles>`

func characterIndentParts(indent string) map[string]string {
	parts := resolvedStylesTestParts(characterIndentStyles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="ListParagraph"/>` + indent + `</w:pPr><w:r><w:t>test</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	return parts
}

func characterIndentDiagnostics(t *testing.T, indent string) (*NativeResolvedLayoutInputV1, []string) {
	t.Helper()
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(characterIndentParts(indent))))
	if err != nil {
		t.Fatal(err)
	}
	codes := []string{}
	for _, diagnostic := range resolved.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	return resolved, codes
}

func characterIndentUnsupported(t *testing.T, indent string) []string {
	t.Helper()
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(characterIndentParts(indent))))
	if err != nil {
		t.Fatal(err)
	}
	codes := []string{}
	for _, entry := range doc.Unsupported {
		codes = append(codes, entry.Code)
	}
	return codes
}

// w:leftChars="0" is Word's ordinary way of cancelling a style's character
// indent. Zero characters is zero twips for every font, so it needs no metric
// and must not leave the paragraph unresolved. Cancelling the character
// channel is not the same as cancelling the indent: Word renders the cjklist
// benchmark documents, whose list paragraphs carry exactly this markup over a
// style with w:leftChars="200" w:left="480", with the label at the 480-twip
// text margin, so the inherited absolute indent survives. Only an absolute
// companion on the same w:ind element is superseded by the zero.
func TestNativeZeroCharacterIndentResolvesWithoutFontMetrics(t *testing.T) {
	resolved, codes := characterIndentDiagnostics(t, `<w:ind w:leftChars="0"/>`)
	for _, code := range codes {
		if code == "CHARACTER_INDENT_PRESERVED" {
			t.Fatalf("zero character indent still refuses: %v", codes)
		}
	}
	if len(resolved.Paragraphs) != 1 {
		t.Fatalf("paragraphs = %#v", resolved.Paragraphs)
	}
	left := resolved.Paragraphs[0].Properties.IndentLeftTwips
	if left == nil || *left != 480 {
		t.Fatalf("zero character indent discarded the inherited absolute indent: %#v", left)
	}
	if codes := characterIndentUnsupported(t, `<w:ind w:leftChars="0"/>`); strings.Contains(strings.Join(codes, ","), "UNMODELED_PARAGRAPH_INDENT") {
		t.Fatalf("extraction still refuses a zero character indent: %v", codes)
	}
	companion, codes := characterIndentDiagnostics(t, `<w:ind w:leftChars="0" w:left="720"/>`)
	for _, code := range codes {
		if code == "CHARACTER_INDENT_PRESERVED" {
			t.Fatalf("zero character indent with a companion still refuses: %v", codes)
		}
	}
	if left := companion.Paragraphs[0].Properties.IndentLeftTwips; left == nil || *left != 0 {
		t.Fatalf("zero character indent did not supersede its own element's twip indent: %#v", left)
	}
}

// Word writes the twip conversion of a non-zero character measure alongside it.
// That companion is the producer's own conversion, so it is adopted verbatim.
func TestNativeCharacterIndentAdoptsProducerTwipCompanion(t *testing.T) {
	resolved, codes := characterIndentDiagnostics(t, `<w:ind w:firstLineChars="200" w:firstLine="480"/>`)
	for _, code := range codes {
		if code == "CHARACTER_INDENT_PRESERVED" {
			t.Fatalf("attested character indent still refuses: %v", codes)
		}
	}
	first := resolved.Paragraphs[0].Properties.FirstLineTwips
	if first == nil || *first != 480 {
		t.Fatalf("first-line indent = %#v, want the attested 480 twips", first)
	}
	if codes := characterIndentUnsupported(t, `<w:ind w:firstLineChars="200" w:firstLine="480"/>`); strings.Contains(strings.Join(codes, ","), "UNMODELED_PARAGRAPH_INDENT") {
		t.Fatalf("extraction still refuses an attested character indent: %v", codes)
	}
}

// A non-zero character measure with no conversion in the source still needs a
// character width this layer does not have. It keeps refusing rather than
// guessing one.
func TestNativeUnattestedCharacterIndentKeepsRefusing(t *testing.T) {
	for name, indent := range map[string]string{
		"no companion":      `<w:ind w:leftChars="200"/>`,
		"invalid measure":   `<w:ind w:leftChars="two" w:left="480"/>`,
		"invalid companion": `<w:ind w:leftChars="200" w:left="wide"/>`,
	} {
		t.Run(name, func(t *testing.T) {
			_, codes := characterIndentDiagnostics(t, indent)
			found := false
			for _, code := range codes {
				if code == "CHARACTER_INDENT_PRESERVED" {
					found = true
				}
			}
			if !found {
				t.Fatalf("unattested character indent was resolved: %v", codes)
			}
			if !strings.Contains(strings.Join(characterIndentUnsupported(t, indent), ","), "UNMODELED_PARAGRAPH_INDENT") {
				t.Fatalf("extraction accepted an unattested character indent: %v", characterIndentUnsupported(t, indent))
			}
		})
	}
}
