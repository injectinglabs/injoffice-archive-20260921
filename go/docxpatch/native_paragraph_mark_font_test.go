package docxpatch

import (
	"strings"
	"testing"
)

func TestParagraphMarkComplexFontSlotStaysScriptQualified(t *testing.T) {
	for _, test := range []struct {
		name, text, extra, cs     string
		sourceSafe, scriptBlocked bool
	}{
		{"latin", "ABC", "", "Arial", true, false},
		{"empty", "", "", "Arial", true, false},
		{"rtl", "ABC", "<w:bidi/>", "Arial", true, true},
		{"mixed", "A漢", "", "Arial", true, true},
		{"invalid", "ABC", "", "", false, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr>` + test.extra + `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="` + test.cs + `"/></w:rPr></w:pPr><w:r><w:t>` + test.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
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
			if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" {
				t.Fatal("paragraph mark source became mutable")
			}
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
