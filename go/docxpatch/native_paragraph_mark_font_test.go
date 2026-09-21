package docxpatch

import (
	"strings"
	"testing"
)

func TestParagraphMarkComplexFontSlotStaysScriptQualified(t *testing.T) {
	// ECMA-376 17.3.2.30: the mark's own w:rtl selects the complex-script slot.
	// Paragraph-level w:bidi (17.3.1.6) only orders the line and leaves plain
	// Latin text in the ascii slot.
	for _, test := range []struct {
		name, text, extra, markExtra, cs string
		sourceSafe, scriptBlocked        bool
	}{
		{"latin", "ABC", "", "", "Arial", true, false},
		{"empty", "", "", "", "Arial", true, false},
		{"paragraph bidi", "ABC", "<w:bidi/>", "", "Arial", true, false},
		// The mark's own switch still selects the complex-script slot, and the
		// mark's w:cs names the face it selects, so it resolves instead of
		// refusing. A switch with no face in the slot still refuses.
		{"mark rtl", "ABC", "", "<w:rtl/>", "Arial", true, false},
		{"paragraph bidi with mark rtl", "ABC", "<w:bidi/>", "<w:rtl/>", "Arial", true, false},
		{"mark cs switch", "ABC", "", "<w:cs/>", "Arial", true, false},
		{"mixed", "A漢", "", "", "Arial", true, true},
		{"invalid", "ABC", "", "", "", false, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr>` + test.extra + `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="` + test.cs + `"/>` + test.markExtra + `</w:rPr></w:pPr><w:r><w:t>` + test.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			data := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			layout, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if hasResolutionDiagnostic(layout, "SCRIPT_FONT_PRESERVED") != test.scriptBlocked {
				t.Fatalf("script qualification: %#v", layout.Diagnostics)
			}
			assertNativeTextPreservationPolicy(t, data, doc, true)
			assertNativePropertyPatchRefused(t, data, doc)
			encoded, err := EncodeNativeDocumentV1(doc)
			if err != nil {
				t.Fatal(err)
			}
			if strings.Contains(string(encoded), `"code":"UNMODELED_PARAGRAPH_MARK_PROPERTIES"`) == test.sourceSafe {
				t.Fatalf("source guard wrong: %s", encoded)
			}
		})
	}
}
