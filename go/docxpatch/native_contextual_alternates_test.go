package docxpatch

import (
	"strings"
	"testing"
)

const nativeContextualAlternatesTestNSDecl = ` xmlns:w14="` + nativeWordML2010 + `"`

// A docDefaults contextual-alternates request is scoped to the document, so
// refusing it blocks every paragraph. Only an enabled request is accepted: it
// names the calt the declared shaper defaults already apply. Disabling calt
// turns a default feature off, which v1 has no input for.
func TestNativeContextualAlternatesDocDefaults(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
	}{
		{"absent value defaults on", `<w14:cntxtAlts/>`, true},
		{"explicit true", `<w14:cntxtAlts w14:val="true"/>`, true},
		{"explicit 1", `<w14:cntxtAlts w14:val="1"/>`, true},
		{"explicit on", `<w14:cntxtAlts w14:val="on"/>`, true},
		{"explicit false", `<w14:cntxtAlts w14:val="false"/>`, false},
		{"explicit 0", `<w14:cntxtAlts w14:val="0"/>`, false},
		{"explicit off", `<w14:cntxtAlts w14:val="off"/>`, false},
		{"invalid value", `<w14:cntxtAlts w14:val="maybe"/>`, false},
		{"foreign value attribute", `<w14:cntxtAlts w:val="true"/>`, false},
		{"extra attribute", `<w14:cntxtAlts w14:val="true" w14:extra="1"/>`, false},
		{"nested markup", `<w14:cntxtAlts><w14:x/></w14:cntxtAlts>`, false},
		{"duplicate element", `<w14:cntxtAlts/><w14:cntxtAlts/>`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `"` + nativeContextualAlternatesTestNSDecl + `><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/>` + tc.markup + `</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(styles))))
			if err != nil {
				t.Fatal(err)
			}
			if got := hasResolutionDiagnostic(resolved, "CONTEXTUAL_ALTERNATES_MATCH_SHAPER"); got != tc.accepted {
				t.Fatalf("CONTEXTUAL_ALTERNATES_MATCH_SHAPER=%v, want %v: %#v", got, tc.accepted, resolved.Diagnostics)
			}
			if got := hasResolutionDiagnostic(resolved, "FOREIGN_RUN_PROPERTY"); got == tc.accepted {
				t.Fatalf("FOREIGN_RUN_PROPERTY=%v, want %v: %#v", got, !tc.accepted, resolved.Diagnostics)
			}
			for _, diagnostic := range resolved.Diagnostics {
				if diagnostic.Code != "CONTEXTUAL_ALTERNATES_MATCH_SHAPER" {
					continue
				}
				if diagnostic.ScopeID != resolved.DocumentID || diagnostic.Preservation != "preserve-verbatim" {
					t.Fatalf("contextual alternates diagnostic lost its document scope or preservation: %#v", diagnostic)
				}
			}
		})
	}
}

// The enabled request on a direct run rPr is accepted by the extractor without
// marking the run's properties partial; disabling calt stays foreign.
func TestNativeContextualAlternatesDirectRunProperties(t *testing.T) {
	for _, tc := range []struct {
		markup   string
		accepted bool
	}{{`<w14:cntxtAlts/>`, true}, {`<w14:cntxtAlts w14:val="1"/>`, true}, {`<w14:cntxtAlts w14:val="0"/>`, false}} {
		t.Run(tc.markup, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"` + nativeContextualAlternatesTestNSDecl + `><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/>` + tc.markup + `</w:rPr><w:t>Alternates</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if got := hasUnsupportedCode(doc, "CONTEXTUAL_ALTERNATES_MATCH_SHAPER"); got != tc.accepted {
				t.Fatalf("CONTEXTUAL_ALTERNATES_MATCH_SHAPER=%v, want %v: %#v", got, tc.accepted, doc.Unsupported)
			}
			for _, code := range []string{"FOREIGN_RUN_PROPERTY", "PARTIAL_RUN_PROPERTIES"} {
				if got := hasUnsupportedCode(doc, code); got == tc.accepted {
					t.Fatalf("%s=%v, want %v: %#v", code, got, !tc.accepted, doc.Unsupported)
				}
			}
		})
	}
}

// The helper qualifies only an enabled exact leaf, in both WordprocessingML
// namespaces it can accompany, and qualifies no neighbouring w14 property that
// selects a non-default OpenType feature.
func TestNativeShaperDefaultContextualAlternates(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, markup := range []string{
			`<w14:cntxtAlts/>`,
			`<w14:cntxtAlts w14:val="0"/>`,
			`<w14:ligatures w14:val="standard"/>`,
			`<w14:numForm w14:val="lining"/>`,
			`<w14:numForm w14:val="oldStyle"/>`,
			`<w14:numSpacing w14:val="tabular"/>`,
			`<w14:numSpacing w14:val="proportional"/>`,
			`<w14:stylisticSets><w14:styleSet w14:id="2"/></w14:stylisticSets>`,
		} {
			root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`"`+nativeContextualAlternatesTestNSDecl+`>`+markup+`</w:rPr>`))
			if err != nil {
				t.Fatal(err)
			}
			want := markup == `<w14:cntxtAlts/>`
			if got := nativeShaperDefaultContextualAlternates(root.Children[0], root); got != want {
				t.Fatalf("%s: qualified=%v, want %v", markup, got, want)
			}
			if strings.Contains(markup, "cntxtAlts") && nativeShaperDefaultLigatureMode(root.Children[0], root) {
				t.Fatalf("%s: contextual alternates must not qualify as a ligature mode", markup)
			}
		}
	}
}
