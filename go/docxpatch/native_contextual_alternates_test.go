package docxpatch

import (
	"strings"
	"testing"
)

const nativeContextualAlternatesTestNSDecl = ` xmlns:w14="` + nativeWordML2010 + `"`

// A docDefaults contextual-alternates request is scoped to the document, so
// refusing it blocks every paragraph. Only an enabled request STATES the calt
// the declared shaper defaults already apply. Disabling calt turns a default
// feature off, which v1 has no input for, so the run is painted WITH calt and
// the unapplied property is disclosed instead; malformed markup stays foreign.
func TestNativeContextualAlternatesDocDefaults(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
		code     string
	}{
		{"absent value defaults on", `<w14:cntxtAlts/>`, true, ""},
		{"explicit true", `<w14:cntxtAlts w14:val="true"/>`, true, ""},
		{"explicit 1", `<w14:cntxtAlts w14:val="1"/>`, true, ""},
		{"explicit on", `<w14:cntxtAlts w14:val="on"/>`, true, ""},
		{"explicit false", `<w14:cntxtAlts w14:val="false"/>`, false, "CONTEXTUAL_ALTERNATES_UNAPPLIED"},
		{"explicit 0", `<w14:cntxtAlts w14:val="0"/>`, false, "CONTEXTUAL_ALTERNATES_UNAPPLIED"},
		{"explicit off", `<w14:cntxtAlts w14:val="off"/>`, false, "CONTEXTUAL_ALTERNATES_UNAPPLIED"},
		{"invalid value", `<w14:cntxtAlts w14:val="maybe"/>`, false, "FOREIGN_RUN_PROPERTY"},
		{"foreign value attribute", `<w14:cntxtAlts w:val="true"/>`, false, "FOREIGN_RUN_PROPERTY"},
		{"extra attribute", `<w14:cntxtAlts w14:val="true" w14:extra="1"/>`, false, "FOREIGN_RUN_PROPERTY"},
		{"nested markup", `<w14:cntxtAlts><w14:x/></w14:cntxtAlts>`, false, "FOREIGN_RUN_PROPERTY"},
		{"duplicate element", `<w14:cntxtAlts/><w14:cntxtAlts/>`, false, "FOREIGN_RUN_PROPERTY"},
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
			if tc.code != "" && !hasResolutionDiagnostic(resolved, tc.code) {
				t.Fatalf("resolver did not record %s: %#v", tc.code, resolved.Diagnostics)
			}
			if got := hasResolutionDiagnostic(resolved, "FOREIGN_RUN_PROPERTY"); got != (tc.code == "FOREIGN_RUN_PROPERTY") {
				t.Fatalf("FOREIGN_RUN_PROPERTY=%v, want %v: %#v", got, tc.code == "FOREIGN_RUN_PROPERTY", resolved.Diagnostics)
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
// marking the run's properties partial; disabling calt is disclosed as an
// unapplied property instead, and malformed markup stays foreign.
func TestNativeContextualAlternatesDirectRunProperties(t *testing.T) {
	for _, tc := range []struct {
		markup   string
		accepted bool
		code     string
	}{
		{`<w14:cntxtAlts/>`, true, ""},
		{`<w14:cntxtAlts w14:val="1"/>`, true, ""},
		{`<w14:cntxtAlts w14:val="0"/>`, false, "CONTEXTUAL_ALTERNATES_UNAPPLIED"},
		{`<w14:cntxtAlts w14:val="maybe"/>`, false, "FOREIGN_RUN_PROPERTY"},
	} {
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
			if tc.code != "" && !hasUnsupportedCode(doc, tc.code) {
				t.Fatalf("extractor did not record %s: %#v", tc.code, doc.Unsupported)
			}
			if got := hasUnsupportedCode(doc, "FOREIGN_RUN_PROPERTY"); got != (tc.code == "FOREIGN_RUN_PROPERTY") {
				t.Fatalf("FOREIGN_RUN_PROPERTY=%v, want %v: %#v", got, tc.code == "FOREIGN_RUN_PROPERTY", doc.Unsupported)
			}
			if got := hasUnsupportedCode(doc, "PARTIAL_RUN_PROPERTIES"); got == tc.accepted {
				t.Fatalf("PARTIAL_RUN_PROPERTIES=%v, want %v: %#v", got, !tc.accepted, doc.Unsupported)
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
