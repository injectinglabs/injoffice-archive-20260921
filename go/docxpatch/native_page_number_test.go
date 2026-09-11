package docxpatch

import (
	"strings"
	"testing"
)

func TestNativeSourceRefusesInternalLayoutFieldMarker(t *testing.T) {
	for _, marker := range []string{"", "PAGE"} {
		doc, err := DecodeNativeDocumentV1(nativeFixture(t))
		if err != nil {
			t.Fatal(err)
		}
		doc.Body.Blocks[0].Paragraph.Runs[0].LayoutPageField = &marker
		found := false
		for _, issue := range ValidateNativeDocumentV1(doc) {
			if strings.Contains(issue.Path, "layout_page_field") {
				found = true
			}
		}
		if !found {
			t.Fatal("internal marker was admitted as authored source")
		}
	}
}

func TestNativeDecimalPageNumberStart(t *testing.T) {
	for _, tc := range []struct {
		markup string
		pass   bool
	}{
		{`<w:pgNumType w:fmt="decimal" w:start="7"/>`, true},
		{`<w:pgNumType w:start="0"/>`, true},
		{`<w:pgNumType w:fmt="upperRoman" w:start="7"/>`, false},
		{`<w:pgNumType w:start="-1"/>`, false},
		{`<w:pgNumType w:start="1000000"/>`, false},
		{`<w:pgNumType w:start="07"/>`, false},
		{`<w:pgNumType w:start="7" w:chapStyle="1"/>`, false},
		{`<w:pgNumType w:start="7"/><w:pgNumType w:start="8"/>`, false},
	} {
		t.Run(tc.markup, func(t *testing.T) {
			parts := transitionalNativeParts()
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:sectPr>`, `<w:sectPr>`+tc.markup, 1)
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			refused := false
			for _, issue := range doc.Unsupported {
				if issue.Code == "UNMODELED_SECTION_PROPERTY" || issue.Code == "DUPLICATE_SECTION_PROPERTY" {
					refused = true
				}
			}
			if refused == tc.pass {
				t.Fatalf("pass=%v unsupported=%+v", tc.pass, doc.Unsupported)
			}
			if tc.pass && doc.Sections[0].PageNumberStart == nil {
				t.Fatal("missing decimal source start")
			}
		})
	}
}
