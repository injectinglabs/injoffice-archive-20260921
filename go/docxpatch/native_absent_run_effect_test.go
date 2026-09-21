package docxpatch

import (
	"testing"
)

// Word writes a numbering level's or a style's run properties out in full, one
// element per property, each carrying the value that means "off". Every such
// element asks for exactly the glyphs, advances and ink an omitted element
// already produces, so the run occupies the same box as the same run without
// it. `mixednumberings.docx` is the corpus case: one abstract numbering level
// carries w:caps, w:smallCaps, w:strike, w:dstrike, w:spacing, w:position,
// w:effect, w:em and w:specVanish, every one of them at its off value, and each
// refused the whole document as UNMODELED_RUN_PROPERTY. Every other value does
// change glyphs, advances or ink and keeps refusing.
//
// One exception is now modeled rather than refused: a nonzero w:spacing is
// character tracking, which the resolver resolves into a painted advance (see
// native_character_spacing_test.go). The writable v1 contract still exposes no
// tracking, so the extractor keeps reporting it as an unmodeled run property -
// `resolverModels` below is exactly that split.
func TestNativeAbsentRunEffectProperty(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
		// True when the resolver resolves the property into layout instead of
		// leaving it unmodeled. Only nonzero character tracking does.
		resolverModels bool
	}{
		{"caps off", `<w:caps w:val="0"/>`, true, false},
		{"caps off spelled false", `<w:caps w:val="false"/>`, true, false},
		{"small caps off", `<w:smallCaps w:val="off"/>`, true, false},
		{"strike off", `<w:strike w:val="0"/>`, true, false},
		{"double strike off", `<w:dstrike w:val="0"/>`, true, false},
		{"special vanish off", `<w:specVanish w:val="0"/>`, true, false},
		{"zero character spacing", `<w:spacing w:val="0"/>`, true, false},
		{"zero baseline position", `<w:position w:val="0"/>`, true, false},
		{"no text effect", `<w:effect w:val="none"/>`, true, false},
		{"no emphasis mark", `<w:em w:val="none"/>`, true, false},

		{"caps on", `<w:caps w:val="1"/>`, false, false},
		{"caps by omission", `<w:caps/>`, false, false},
		{"small caps on", `<w:smallCaps w:val="true"/>`, false, false},
		{"strike on", `<w:strike w:val="on"/>`, false, false},
		{"positive character spacing", `<w:spacing w:val="20"/>`, false, true},
		{"negative character spacing", `<w:spacing w:val="-20"/>`, false, true},
		{"raised baseline", `<w:position w:val="6"/>`, false, false},
		{"blinking text effect", `<w:effect w:val="blinkBackground"/>`, false, false},
		{"dot emphasis mark", `<w:em w:val="dot"/>`, false, false},
		{"unknown attribute", `<w:caps w:val="0" w:other="1"/>`, false, false},
		{"nested markup", `<w:caps w:val="0"><w:b/></w:caps>`, false, false},
		{"duplicate", `<w:caps w:val="0"/><w:caps w:val="0"/>`, false, false},
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
			if got := hasUnsupportedCode(doc, "RUN_EFFECT_ABSENT_PRESERVED"); got != tc.accepted {
				t.Fatalf("extract RUN_EFFECT_ABSENT_PRESERVED=%v, want %v: %#v", got, tc.accepted, doc.Unsupported)
			}
			if got := hasResolutionDiagnostic(resolved, "RUN_EFFECT_ABSENT_PRESERVED"); got != tc.accepted {
				t.Fatalf("resolve RUN_EFFECT_ABSENT_PRESERVED=%v, want %v: %#v", got, tc.accepted, resolved.Diagnostics)
			}
			if got := hasUnsupportedCode(doc, "UNMODELED_RUN_PROPERTY"); got == tc.accepted {
				t.Fatalf("extract UNMODELED_RUN_PROPERTY=%v, want %v: %#v", got, !tc.accepted, doc.Unsupported)
			}
			if want := !tc.accepted && !tc.resolverModels; hasResolutionDiagnostic(resolved, "UNMODELED_RUN_PROPERTY") != want {
				t.Fatalf("resolve UNMODELED_RUN_PROPERTY want %v: %#v", want, resolved.Diagnostics)
			}
			if got := resolved.Runs[0].Properties.LetterSpacingTwips != nil; got != tc.resolverModels {
				t.Fatalf("resolve letter_spacing_twips=%v, want %v", got, tc.resolverModels)
			}
			// The accepted leaf is disclosed without making the run's exposed
			// properties partial. Text editing preserves the source decoration.
			// These decorations (including unmodeled source shapes) do not own
			// text. Keep paint diagnostics, but prove that editing retains them.
			if hasUnsupportedCode(doc, "PARTIAL_RUN_PROPERTIES") != (tc.name == "duplicate") {
				t.Fatalf("decoration incorrectly blocks text: %#v", doc.Unsupported)
			}
			assertNativeTextPreservationPolicy(t, data, doc, tc.name != "duplicate")
		})
	}
}

// The qualifying shape is the same in both WordprocessingML namespaces, and the
// helper never qualifies a property outside the closed set that happens to
// share the shape.
func TestNativeAbsentRunEffectHelper(t *testing.T) {
	accepted := map[string]bool{`<w:caps w:val="0"/>`: true, `<w:em w:val="none"/>`: true, `<w:position w:val="0"/>`: true}
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for _, markup := range []string{`<w:caps w:val="0"/>`, `<w:em w:val="none"/>`, `<w:position w:val="0"/>`, `<w:caps w:val="1"/>`, `<w:sz w:val="0"/>`, `<w:u w:val="none"/>`, `<w:vanish w:val="0"/>`} {
			root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`">`+markup+`</w:rPr>`))
			if err != nil {
				t.Fatal(err)
			}
			if got := nativeAbsentRunEffect(root.Children[0], root, ns); got != accepted[markup] {
				t.Fatalf("%s: qualified=%v, want %v", markup, got, accepted[markup])
			}
		}
	}
}
