package docxpatch

import (
	"strings"
	"testing"
)

func TestPagePrinterCodePreservesExplicitGeometry(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, test := range []struct {
			name, attrs string
			valid       bool
		}{
			{"absent", "", true},
			{"zero", ` w:code="0"`, true},
			{"a4", ` w:code="9"`, true},
			{"maximum", ` w:code="118"`, true},
			{"negative", ` w:code="-1"`, false},
			{"large", ` w:code="119"`, false},
			{"invalid", ` w:code="paper"`, false},
			{"foreign", ` xmlns:x="urn:foreign" x:code="9"`, false},
			{"unknown", ` w:code="9" w:extra="1"`, false},
		} {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"` + test.attrs + `/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>`
			if ns == wordMLStrict {
				for key, value := range parts {
					parts[key] = strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
				}
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			before := string(data)
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			encoded, err := EncodeNativeDocumentV1(doc)
			if err != nil {
				t.Fatal(err)
			}
			if strings.Contains(string(encoded), `"code":"UNMODELED_SECTION_PROPERTY"`) == test.valid {
				t.Fatalf("%s: unexpected section qualification", test.name)
			}
			if test.valid && (*doc.Sections[0].Page.WidthTwips != 11906 || *doc.Sections[0].Page.HeightTwips != 16838) {
				t.Fatalf("%s: printer code replaced authored geometry", test.name)
			}
			if string(data) != before {
				t.Fatal("source changed")
			}
		}
	}
}
