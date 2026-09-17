package docxpatch

import (
	"strings"
	"testing"
)

// A w:bdr that names ST_Border none or nil asks for no border: Word paints no
// stroke and reserves no space, so the run occupies the same box as the same
// run without the element. Every other value paints a border this tier does not
// model, and anything but an exact single leaf states something this reading
// does not cover.
func TestNativeAbsentRunBorderProperty(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
	}{
		{"none", `<w:bdr w:val="none"/>`, true},
		{"nil", `<w:bdr w:val="nil"/>`, true},
		{"none with inert stroke attributes", `<w:bdr w:val="none" w:sz="0" w:space="0" w:color="auto" w:frame="1"/>`, true},
		{"none with inert theme attributes", `<w:bdr w:val="none" w:themeColor="accent1" w:themeTint="99" w:themeShade="11" w:shadow="0"/>`, true},
		{"painted single border", `<w:bdr w:val="single" w:sz="4" w:space="0" w:color="auto"/>`, false},
		{"painted double border", `<w:bdr w:val="double"/>`, false},
		{"missing value", `<w:bdr w:sz="0"/>`, false},
		{"unknown attribute", `<w:bdr w:val="none" w:other="1"/>`, false},
		{"nested markup", `<w:bdr w:val="none"><w:b/></w:bdr>`, false},
		{"text content", `<w:bdr w:val="none">x</w:bdr>`, false},
		{"duplicate", `<w:bdr w:val="none"/><w:bdr w:val="none"/>`, false},
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
			if got := hasUnsupportedCode(doc, "RUN_BORDER_ABSENT_PRESERVED"); got != tc.accepted {
				t.Fatalf("extract RUN_BORDER_ABSENT_PRESERVED=%v, want %v: %#v", got, tc.accepted, doc.Unsupported)
			}
			if got := hasResolutionDiagnostic(resolved, "RUN_BORDER_ABSENT_PRESERVED"); got != tc.accepted {
				t.Fatalf("resolve RUN_BORDER_ABSENT_PRESERVED=%v, want %v: %#v", got, tc.accepted, resolved.Diagnostics)
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
func TestNativeAbsentRunBorderHelper(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, markup := range []string{`<w:bdr w:val="none"/>`, `<w:bdr w:val="nil"/>`, `<w:bdr w:val="single"/>`, `<w:bdr/>`, `<w:u w:val="none"/>`} {
			root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`">`+markup+`</w:rPr>`))
			if err != nil {
				t.Fatal(err)
			}
			want := markup == `<w:bdr w:val="none"/>` || markup == `<w:bdr w:val="nil"/>`
			if got := nativeAbsentRunBorder(root.Children[0], root, ns); got != want {
				t.Fatalf("%s: qualified=%v, want %v", markup, got, want)
			}
		}
	}
}

// A foreign-namespace element of the same local name is not this property.
func TestNativeAbsentRunBorderRejectsForeignNamespace(t *testing.T) {
	root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+wordMLTransitional+`" xmlns:x="urn:foreign"><x:bdr w:val="none" xmlns:w="`+wordMLTransitional+`"/></w:rPr>`))
	if err != nil {
		t.Fatal(err)
	}
	if nativeAbsentRunBorder(root.Children[0], root, wordMLTransitional) {
		t.Fatal("foreign bdr was qualified")
	}
	if !strings.HasPrefix(root.Children[0].Name.Space, "urn:") {
		t.Fatalf("fixture namespace: %q", root.Children[0].Name.Space)
	}
}
