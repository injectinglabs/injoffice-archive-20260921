package docxpatch

import (
	"strings"
	"testing"
)

// w:bCs and w:iCs are the complex-script companions of w:b and w:i. The layout
// resolver already carries them as deferred script slots, so extraction must
// record them as preserved rather than as unmodeled run markup that refuses the
// whole body. Malformed, duplicate and nested toggles keep refusing by name.
func TestExtractNativeComplexScriptTogglesArePreserved(t *testing.T) {
	for _, test := range []struct {
		name     string
		markup   string
		unmodled bool
	}{
		{name: "implicit and explicit toggles", markup: `<w:bCs/><w:iCs w:val="0"/>`},
		{name: "malformed toggle value", markup: `<w:bCs w:val="maybe"/>`, unmodled: true},
		{name: "duplicate toggle", markup: `<w:bCs/><w:bCs/>`, unmodled: true},
		{name: "toggle with unknown attribute", markup: `<w:iCs w:val="1" w:unknown="1"/>`, unmodled: true},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := transitionalNativeParts()
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:rPr><w:b/>`, `<w:rPr>`+test.markup+`<w:b/>`, 1)
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			refused := false
			for _, entry := range doc.Unsupported {
				if entry.Code == "UNMODELED_RUN_PROPERTY" || entry.Code == "PARTIAL_RUN_PROPERTIES" {
					refused = true
				}
			}
			if refused != test.unmodled {
				t.Fatalf("unmodeled run-property refusal = %v, want %v: %#v", refused, test.unmodled, doc.Unsupported)
			}
			paragraph := doc.Body.Blocks[0].Paragraph
			if paragraph == nil || paragraph.EditPolicy.Mode != "read-only" {
				t.Fatalf("a preserved complex-script toggle must keep the paragraph read-only: %#v", paragraph)
			}
		})
	}
}
