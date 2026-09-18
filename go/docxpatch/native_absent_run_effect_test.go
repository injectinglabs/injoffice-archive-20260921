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
func TestNativeAbsentRunEffectProperty(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
	}{
		{"caps off", `<w:caps w:val="0"/>`, true},
		{"caps off spelled false", `<w:caps w:val="false"/>`, true},
		{"small caps off", `<w:smallCaps w:val="off"/>`, true},
		{"strike off", `<w:strike w:val="0"/>`, true},
		{"double strike off", `<w:dstrike w:val="0"/>`, true},
		{"special vanish off", `<w:specVanish w:val="0"/>`, true},
		{"zero character spacing", `<w:spacing w:val="0"/>`, true},
		{"zero baseline position", `<w:position w:val="0"/>`, true},
		{"no text effect", `<w:effect w:val="none"/>`, true},
		{"no emphasis mark", `<w:em w:val="none"/>`, true},

		{"caps on", `<w:caps w:val="1"/>`, false},
		{"caps by omission", `<w:caps/>`, false},
		{"small caps on", `<w:smallCaps w:val="true"/>`, false},
		{"strike on", `<w:strike w:val="on"/>`, false},
		{"positive character spacing", `<w:spacing w:val="20"/>`, false},
		{"negative character spacing", `<w:spacing w:val="-20"/>`, false},
		{"raised baseline", `<w:position w:val="6"/>`, false},
		{"blinking text effect", `<w:effect w:val="blinkBackground"/>`, false},
		{"dot emphasis mark", `<w:em w:val="dot"/>`, false},
		{"unknown attribute", `<w:caps w:val="0" w:other="1"/>`, false},
		{"nested markup", `<w:caps w:val="0"><w:b/></w:caps>`, false},
		{"duplicate", `<w:caps w:val="0"/><w:caps w:val="0"/>`, false},
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
