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
		// A named ST_Ligatures value this tier does not apply is recorded as
		// LIGATURE_MODE_UNAPPLIED, not as foreign markup, so the approximate
		// tier can paint the run and disclose the unapplied mode. Only markup
		// outside that closed reading stays FOREIGN_RUN_PROPERTY.
		foreign bool
	}{
		{"standard contextual", `<w14:ligatures w14:val="standardContextual"/>`, true, false},
		{"none", `<w14:ligatures w14:val="none"/>`, false, false},
		{"all", `<w14:ligatures w14:val="all"/>`, false, false},
		{"historical", `<w14:ligatures w14:val="historical"/>`, false, false},
		{"discretional", `<w14:ligatures w14:val="standardDiscretional"/>`, false, false},
		{"absent value", `<w14:ligatures/>`, false, true},
		{"unknown value", `<w14:ligatures w14:val="ligaturesEverywhere"/>`, false, true},
		{"foreign value attribute", `<w14:ligatures w:val="standardContextual"/>`, false, true},
		{"extra attribute", `<w14:ligatures w14:val="standardContextual" w14:extra="1"/>`, false, true},
		{"nested markup", `<w14:ligatures w14:val="standardContextual"><w14:x/></w14:ligatures>`, false, true},
		{"duplicate element", `<w14:ligatures w14:val="standardContextual"/><w14:ligatures w14:val="standardContextual"/>`, false, true},
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
			if got := hasResolutionDiagnostic(resolved, "FOREIGN_RUN_PROPERTY"); got != tc.foreign {
				t.Fatalf("FOREIGN_RUN_PROPERTY=%v, want %v: %#v", got, tc.foreign, resolved.Diagnostics)
			}
			if got := hasResolutionDiagnostic(resolved, "LIGATURE_MODE_UNAPPLIED"); got != (!tc.accepted && !tc.foreign) {
				t.Fatalf("LIGATURE_MODE_UNAPPLIED=%v, want %v: %#v", got, !tc.accepted && !tc.foreign, resolved.Diagnostics)
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
			// A named mode this tier does not apply is its own code now, so the
			// run stays unsafe (PARTIAL_RUN_PROPERTIES) but is never foreign.
			if hasUnsupportedCode(doc, "FOREIGN_RUN_PROPERTY") {
				t.Fatalf("a named ST_Ligatures value must not be foreign markup: %#v", doc.Unsupported)
			}
			if got := hasUnsupportedCode(doc, "LIGATURE_MODE_UNAPPLIED"); got == tc.accepted {
				t.Fatalf("LIGATURE_MODE_UNAPPLIED=%v, want %v: %#v", got, !tc.accepted, doc.Unsupported)
			}
			if got := hasUnsupportedCode(doc, "PARTIAL_RUN_PROPERTIES"); got == tc.accepted {
				t.Fatalf("PARTIAL_RUN_PROPERTIES=%v, want %v: %#v", got, !tc.accepted, doc.Unsupported)
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
