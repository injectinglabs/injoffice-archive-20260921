package docxpatch

import (
	"strings"
	"testing"
)

const nativeLigatureTestNSDecl = ` xmlns:w14="` + nativeWordML2010 + `"`

// A docDefaults ligature mode is scoped to the document, so refusing it blocks
// every paragraph. Only the value the declared shaper defaults already apply is
// accepted; the remaining modes name feature selections v1 cannot perform.
func TestNativeLigatureModeDocDefaults(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
	}{
		{"standard contextual", `<w14:ligatures w14:val="standardContextual"/>`, true},
		{"none", `<w14:ligatures w14:val="none"/>`, false},
		{"all", `<w14:ligatures w14:val="all"/>`, false},
		{"historical", `<w14:ligatures w14:val="historical"/>`, false},
		{"discretional", `<w14:ligatures w14:val="standardDiscretional"/>`, false},
		{"absent value", `<w14:ligatures/>`, false},
		{"foreign value attribute", `<w14:ligatures w:val="standardContextual"/>`, false},
		{"extra attribute", `<w14:ligatures w14:val="standardContextual" w14:extra="1"/>`, false},
		{"nested markup", `<w14:ligatures w14:val="standardContextual"><w14:x/></w14:ligatures>`, false},
		{"duplicate element", `<w14:ligatures w14:val="standardContextual"/><w14:ligatures w14:val="standardContextual"/>`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `"` + nativeLigatureTestNSDecl + `><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/>` + tc.markup + `</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(styles))))
			if err != nil {
				t.Fatal(err)
			}
			if got := hasResolutionDiagnostic(resolved, "LIGATURE_MODE_MATCHES_SHAPER"); got != tc.accepted {
				t.Fatalf("LIGATURE_MODE_MATCHES_SHAPER=%v, want %v: %#v", got, tc.accepted, resolved.Diagnostics)
			}
			if got := hasResolutionDiagnostic(resolved, "FOREIGN_RUN_PROPERTY"); got == tc.accepted {
				t.Fatalf("FOREIGN_RUN_PROPERTY=%v, want %v: %#v", got, !tc.accepted, resolved.Diagnostics)
			}
			for _, diagnostic := range resolved.Diagnostics {
				if diagnostic.Code != "LIGATURE_MODE_MATCHES_SHAPER" {
					continue
				}
				if diagnostic.ScopeID != resolved.DocumentID || diagnostic.Preservation != "preserve-verbatim" {
					t.Fatalf("ligature diagnostic lost its document scope or preservation: %#v", diagnostic)
				}
			}
		})
	}
}

// The same value on a direct run rPr is accepted by the extractor without
// marking the run's properties partial, and every other value stays foreign.
func TestNativeLigatureModeDirectRunProperties(t *testing.T) {
	for _, tc := range []struct {
		value    string
		accepted bool
	}{{"standardContextual", true}, {"none", false}, {"all", false}} {
		t.Run(tc.value, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"` + nativeLigatureTestNSDecl + `><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/><w14:ligatures w14:val="` + tc.value + `"/></w:rPr><w:t>Ligature</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if got := hasUnsupportedCode(doc, "LIGATURE_MODE_MATCHES_SHAPER"); got != tc.accepted {
				t.Fatalf("LIGATURE_MODE_MATCHES_SHAPER=%v, want %v: %#v", got, tc.accepted, doc.Unsupported)
			}
			for _, code := range []string{"FOREIGN_RUN_PROPERTY", "PARTIAL_RUN_PROPERTIES"} {
				if got := hasUnsupportedCode(doc, code); got == tc.accepted {
					t.Fatalf("%s=%v, want %v: %#v", code, got, !tc.accepted, doc.Unsupported)
				}
			}
		})
	}
}

// The accepted mode is the only ligature markup the helper qualifies, and it
// stays qualified in both WordprocessingML namespaces it can accompany.
func TestNativeShaperDefaultLigatureMode(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, markup := range []string{`<w14:ligatures w14:val="standardContextual"/>`, `<w14:ligatures w14:val="none"/>`} {
			root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`"`+nativeLigatureTestNSDecl+`>`+markup+`</w:rPr>`))
			if err != nil {
				t.Fatal(err)
			}
			want := strings.Contains(markup, "standardContextual")
			if got := nativeShaperDefaultLigatureMode(root.Children[0], root); got != want {
				t.Fatalf("%s: qualified=%v, want %v", markup, got, want)
			}
		}
	}
}
