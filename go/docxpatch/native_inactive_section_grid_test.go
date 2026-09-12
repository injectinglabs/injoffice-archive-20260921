package docxpatch

import (
	"strings"
	"testing"
)

func TestInactiveSectionGridQualification(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, test := range []struct {
			name, grid string
			valid      bool
		}{
			{"empty", `<w:docGrid/>`, true},
			{"omitted", `<w:docGrid w:linePitch="360"/>`, true},
			{"default", `<w:docGrid w:type="default" w:linePitch="360" w:charSpace="-100"/>`, true},
			{"active", `<w:docGrid w:type="lines" w:linePitch="360"/>`, false},
			{"active-characters", `<w:docGrid w:type="linesAndChars"/>`, false},
			{"active-snap", `<w:docGrid w:type="snapToChars"/>`, false},
			{"unknown-type", `<w:docGrid w:type="futureGrid"/>`, false},
			{"empty-type", `<w:docGrid w:type=""/>`, false},
			{"invalid", `<w:docGrid w:linePitch="abc"/>`, false},
			{"overflow", `<w:docGrid w:charSpace="2147483648"/>`, false},
			{"foreign", `<w:docGrid xmlns:x="urn:foreign" x:linePitch="360"/>`, false},
			{"unknown", `<w:docGrid w:extra="1"/>`, false},
			{"children", `<w:docGrid><w:unknown/></w:docGrid>`, false},
			{"text", `<w:docGrid>not metadata</w:docGrid>`, false},
			{"duplicate", `<w:docGrid/><w:docGrid/>`, false},
		} {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + ns + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p/><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/>` + test.grid + `</w:sectPr></w:body></w:document>`
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
			blocked := false
			for _, d := range doc.Unsupported {
				if d.Code == "UNMODELED_SECTION_PROPERTY" || d.Code == "DUPLICATE_SECTION_PROPERTY" {
					blocked = true
				}
			}
			if blocked == test.valid {
				t.Fatalf("%s: grid qualification incorrect", test.name)
			}
			if *doc.Sections[0].Page.WidthTwips != 11906 || *doc.Sections[0].Page.HeightTwips != 16838 {
				t.Fatal("grid changed page geometry")
			}
			if string(data) != before {
				t.Fatal("source changed")
			}
		}
	}
}
