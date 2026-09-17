package docxpatch

import (
	"strings"
	"testing"
)

// w:webHidden hides a run in Word's Web Layout view only, so paginated layout
// draws it and the fact is recorded rather than refused. Anything but the exact
// CT_OnOff leaf states something this reading does not cover.
func TestNativeWebHiddenRunProperty(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
	}{
		{"bare", `<w:webHidden/>`, true},
		{"explicit on", `<w:webHidden w:val="1"/>`, true},
		{"explicit off", `<w:webHidden w:val="false"/>`, true},
		{"invalid value", `<w:webHidden w:val="maybe"/>`, false},
		{"unknown attribute", `<w:webHidden w:other="1"/>`, false},
		{"nested markup", `<w:webHidden><w:b/></w:webHidden>`, false},
		{"text content", `<w:webHidden>x</w:webHidden>`, false},
		{"duplicate", `<w:webHidden/><w:webHidden/>`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/>` + tc.markup + `</w:rPr><w:t>Contents</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			data := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if got := hasUnsupportedCode(doc, "WEB_LAYOUT_HIDDEN_RUN_PRESERVED"); got != tc.accepted {
				t.Fatalf("extract WEB_LAYOUT_HIDDEN_RUN_PRESERVED=%v, want %v: %#v", got, tc.accepted, doc.Unsupported)
			}
			if got := hasResolutionDiagnostic(resolved, "WEB_LAYOUT_HIDDEN_RUN_PRESERVED"); got != tc.accepted {
				t.Fatalf("resolve WEB_LAYOUT_HIDDEN_RUN_PRESERVED=%v, want %v: %#v", got, tc.accepted, resolved.Diagnostics)
			}
			if got := hasUnsupportedCode(doc, "UNMODELED_RUN_PROPERTY"); got == tc.accepted {
				t.Fatalf("extract UNMODELED_RUN_PROPERTY=%v, want %v: %#v", got, !tc.accepted, doc.Unsupported)
			}
			if got := hasResolutionDiagnostic(resolved, "UNMODELED_RUN_PROPERTY"); got == tc.accepted {
				t.Fatalf("resolve UNMODELED_RUN_PROPERTY=%v, want %v: %#v", got, !tc.accepted, resolved.Diagnostics)
			}
			// The accepted leaf is disclosed without making the run's exposed
			// properties partial, and never makes the source writable.
			if got := hasUnsupportedCode(doc, "PARTIAL_RUN_PROPERTIES"); got == tc.accepted {
				t.Fatalf("PARTIAL_RUN_PROPERTIES=%v, want %v: %#v", got, !tc.accepted, doc.Unsupported)
			}
			paragraph := doc.Body.Blocks[0].Paragraph
			if paragraph.EditPolicy.Mode != "read-only" || len(paragraph.EditPolicy.AllowedOperations) != 0 {
				t.Fatalf("preserved source became mutable: %#v", paragraph.EditPolicy)
			}
		})
	}
}

// The qualifying shape is the same in both WordprocessingML namespaces, and the
// helper never qualifies a different property that happens to share the shape.
func TestNativeWebLayoutHiddenRunHelper(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, markup := range []string{`<w:webHidden/>`, `<w:vanish/>`, `<w:webHidden w:val="bad"/>`} {
			root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`">`+markup+`</w:rPr>`))
			if err != nil {
				t.Fatal(err)
			}
			want := markup == `<w:webHidden/>`
			if got := nativeWebLayoutHiddenRun(root.Children[0], root, ns); got != want {
				t.Fatalf("%s: qualified=%v, want %v", markup, got, want)
			}
		}
	}
}

// A foreign-namespace element of the same local name is not this property.
func TestNativeWebLayoutHiddenRunRejectsForeignNamespace(t *testing.T) {
	root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+wordMLTransitional+`" xmlns:x="urn:foreign"><x:webHidden/></w:rPr>`))
	if err != nil {
		t.Fatal(err)
	}
	if nativeWebLayoutHiddenRun(root.Children[0], root, wordMLTransitional) {
		t.Fatal("foreign webHidden was qualified")
	}
	if !strings.HasPrefix(root.Children[0].Name.Space, "urn:") {
		t.Fatalf("fixture namespace: %q", root.Children[0].Name.Space)
	}
}
