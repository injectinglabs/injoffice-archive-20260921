package docxpatch

import (
	"strings"
	"testing"
)

const nativeUnappliedTypographyTestNSDecl = ` xmlns:w14="` + nativeWordML2010 + `"`

// Word 2010 run-typography extensions this tier records under their own code so
// the approximate tier can paint the run with the feature unapplied. Every
// accepted case must stop being FOREIGN_RUN_PROPERTY and start being its own
// code; every rejected case must stay foreign markup, because the codes below
// are the ones the approximate tier paints around.
func TestNativeUnappliedTypographicRunFeatureQualification(t *testing.T) {
	for _, tc := range []struct {
		name   string
		markup string
		code   string
	}{
		{"one stylistic set", `<w14:stylisticSets><w14:styleSet w14:id="2"/></w14:stylisticSets>`, "STYLISTIC_SET_UNAPPLIED"},
		{"several stylistic sets", `<w14:stylisticSets><w14:styleSet w14:id="2"/><w14:styleSet w14:id="14"/></w14:stylisticSets>`, "STYLISTIC_SET_UNAPPLIED"},
		{"ligature mode standard", `<w14:ligatures w14:val="standard"/>`, "LIGATURE_MODE_UNAPPLIED"},
		{"ligature mode none", `<w14:ligatures w14:val="none"/>`, "LIGATURE_MODE_UNAPPLIED"},
		{"ligature mode all", `<w14:ligatures w14:val="all"/>`, "LIGATURE_MODE_UNAPPLIED"},
		{"number form lining", `<w14:numForm w14:val="lining"/>`, "NUMBER_FORM_UNAPPLIED"},
		{"number form old style", `<w14:numForm w14:val="oldStyle"/>`, "NUMBER_FORM_UNAPPLIED"},
		{"number spacing tabular", `<w14:numSpacing w14:val="tabular"/>`, "NUMBER_SPACING_UNAPPLIED"},
		{"number spacing proportional", `<w14:numSpacing w14:val="proportional"/>`, "NUMBER_SPACING_UNAPPLIED"},
		{"3d text effect", `<w14:props3d w14:extrusionH="63500"><w14:bevelT w14:w="38100"/><w14:contourClr><w14:srgbClr w14:val="92D050"/></w14:contourClr></w14:props3d>`, "TEXT_EFFECT_3D_UNAPPLIED"},
		{"bare 3d text effect", `<w14:props3d/>`, "TEXT_EFFECT_3D_UNAPPLIED"},

		// The shaper-default values keep their own existing codes, which state
		// shaping this tier performs rather than a deviation it paints through.
		{"shaper default ligature mode", `<w14:ligatures w14:val="standardContextual"/>`, "LIGATURE_MODE_MATCHES_SHAPER"},
		{"shaper default contextual alternates", `<w14:cntxtAlts/>`, "CONTEXTUAL_ALTERNATES_MATCH_SHAPER"},

		// Everything outside the closed set stays foreign markup and keeps
		// refusing on both tiers.
		{"disabled contextual alternates", `<w14:cntxtAlts w14:val="0"/>`, "FOREIGN_RUN_PROPERTY"},
		{"text outline", `<w14:textOutline w14:w="9525"/>`, "FOREIGN_RUN_PROPERTY"},
		{"text fill", `<w14:textFill><w14:solidFill><w14:srgbClr w14:val="FF0000"/></w14:solidFill></w14:textFill>`, "FOREIGN_RUN_PROPERTY"},
		{"shadow", `<w14:shadow w14:blurRad="50800"/>`, "FOREIGN_RUN_PROPERTY"},
		{"unknown ligature mode", `<w14:ligatures w14:val="ligaturesEverywhere"/>`, "FOREIGN_RUN_PROPERTY"},
		{"ligature mode without a value", `<w14:ligatures/>`, "FOREIGN_RUN_PROPERTY"},
		{"unknown number form", `<w14:numForm w14:val="roman"/>`, "FOREIGN_RUN_PROPERTY"},
		{"unknown number spacing", `<w14:numSpacing w14:val="wide"/>`, "FOREIGN_RUN_PROPERTY"},
		{"number form in the wrong namespace", `<w14:numForm w:val="oldStyle"/>`, "FOREIGN_RUN_PROPERTY"},
		{"number form with an extra attribute", `<w14:numForm w14:val="oldStyle" w14:extra="1"/>`, "FOREIGN_RUN_PROPERTY"},
		{"number form with nested markup", `<w14:numForm w14:val="oldStyle"><w14:x/></w14:numForm>`, "FOREIGN_RUN_PROPERTY"},
		{"duplicate number form", `<w14:numForm w14:val="lining"/><w14:numForm w14:val="oldStyle"/>`, "FOREIGN_RUN_PROPERTY"},
		{"empty stylistic sets", `<w14:stylisticSets/>`, "FOREIGN_RUN_PROPERTY"},
		{"stylistic set without an id", `<w14:stylisticSets><w14:styleSet/></w14:stylisticSets>`, "FOREIGN_RUN_PROPERTY"},
		{"stylistic set with a nonnumeric id", `<w14:stylisticSets><w14:styleSet w14:id="ss02"/></w14:stylisticSets>`, "FOREIGN_RUN_PROPERTY"},
		{"stylistic set with foreign children", `<w14:stylisticSets><w14:styleSet w14:id="2"/><w14:other/></w14:stylisticSets>`, "FOREIGN_RUN_PROPERTY"},
		{"stylistic sets with an attribute", `<w14:stylisticSets w14:id="2"><w14:styleSet w14:id="2"/></w14:stylisticSets>`, "FOREIGN_RUN_PROPERTY"},
		// A WordprocessingML element nested inside the 3D effect would be
		// content this tier must not silently discard with the effect.
		{"3d text effect hiding word markup", `<w14:props3d><w:t>hidden</w:t></w14:props3d>`, "FOREIGN_RUN_PROPERTY"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"` + nativeUnappliedTypographyTestNSDecl + `><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/>` + tc.markup + `</w:rPr><w:t>Typography 456</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			document, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if !hasUnsupportedCode(document, tc.code) {
				t.Fatalf("extractor did not record %s: %#v", tc.code, document.Unsupported)
			}
			if tc.code != "FOREIGN_RUN_PROPERTY" && hasUnsupportedCode(document, "FOREIGN_RUN_PROPERTY") {
				t.Fatalf("%s must replace FOREIGN_RUN_PROPERTY, not accompany it: %#v", tc.code, document.Unsupported)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if !hasResolutionDiagnostic(resolved, tc.code) {
				t.Fatalf("resolver did not record %s: %#v", tc.code, resolved.Diagnostics)
			}
			if tc.code != "FOREIGN_RUN_PROPERTY" && hasResolutionDiagnostic(resolved, "FOREIGN_RUN_PROPERTY") {
				t.Fatalf("%s must replace FOREIGN_RUN_PROPERTY in the resolver too: %#v", tc.code, resolved.Diagnostics)
			}
		})
	}
}

// The disclosure a reader sees must name the property that was not applied and
// the measured advance error, or the painted page is silently wrong.
func TestNativeUnappliedTypographicRunFeatureDiscloses(t *testing.T) {
	for _, tc := range []struct {
		markup string
		code   string
		says   []string
	}{
		{`<w14:stylisticSets><w14:styleSet w14:id="2"/></w14:stylisticSets>`, "STYLISTIC_SET_UNAPPLIED", []string{"ss02", "NOT applied", "28,745"}},
		{`<w14:stylisticSets><w14:styleSet w14:id="14"/></w14:stylisticSets>`, "STYLISTIC_SET_UNAPPLIED", []string{"ss14", "NOT applied"}},
		{`<w14:numForm w14:val="oldStyle"/>`, "NUMBER_FORM_UNAPPLIED", []string{"oldStyle", "NOT applied", "68"}},
		{`<w14:numSpacing w14:val="proportional"/>`, "NUMBER_SPACING_UNAPPLIED", []string{"proportional", "NOT applied", "312 font units"}},
		{`<w14:ligatures w14:val="standard"/>`, "LIGATURE_MODE_UNAPPLIED", []string{"standard", "NOT applied"}},
		{`<w14:props3d/>`, "TEXT_EFFECT_3D_UNAPPLIED", []string{"w14:props3d", "NOT applied"}},
	} {
		t.Run(tc.code+" "+tc.markup, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"` + nativeUnappliedTypographyTestNSDecl + `><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Arial"/>` + tc.markup + `</w:rPr><w:t>456</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			document, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			found := ""
			for _, entry := range document.Unsupported {
				if entry.Code == tc.code {
					found = entry.Message
				}
			}
			if found == "" {
				t.Fatalf("no %s record: %#v", tc.code, document.Unsupported)
			}
			for _, want := range tc.says {
				if !strings.Contains(found, want) {
					t.Fatalf("%s disclosure does not name %q: %s", tc.code, want, found)
				}
			}
		})
	}
}
