package docxpatch

import (
	"strings"
	"testing"
)

// Word writes a full CT_RPr of explicit "off" leaves whenever direct character
// formatting is cleared. Each accepted value states the absence of a formatting
// operation, so applying it is applying nothing; every enabled value reshapes
// or repaints the run and keeps its refusal.
func TestNativeAbsentRunFormatting(t *testing.T) {
	for _, tc := range []struct {
		name     string
		markup   string
		accepted bool
	}{
		// The absent forms Word itself writes.
		{"caps off", `<w:caps w:val="0"/>`, true},
		{"caps off false", `<w:caps w:val="false"/>`, true},
		{"smallCaps off", `<w:smallCaps w:val="0"/>`, true},
		{"strike off", `<w:strike w:val="0"/>`, true},
		{"dstrike off", `<w:dstrike w:val="0"/>`, true},
		{"outline off", `<w:outline w:val="0"/>`, true},
		{"shadow off", `<w:shadow w:val="0"/>`, true},
		{"emboss off", `<w:emboss w:val="0"/>`, true},
		{"imprint off", `<w:imprint w:val="0"/>`, true},
		{"specVanish off", `<w:specVanish w:val="0"/>`, true},
		{"effect none", `<w:effect w:val="none"/>`, true},
		{"em none", `<w:em w:val="none"/>`, true},
		{"spacing zero", `<w:spacing w:val="0"/>`, true},
		{"position zero", `<w:position w:val="0"/>`, true},
		{"shd clear auto", `<w:shd w:val="clear" w:color="auto" w:fill="auto"/>`, true},
		// Enabled values reshape or repaint the run and must keep refusing.
		{"caps on", `<w:caps/>`, false},
		{"caps on explicit", `<w:caps w:val="1"/>`, false},
		{"smallCaps on", `<w:smallCaps/>`, false},
		{"strike on", `<w:strike/>`, false},
		{"dstrike on", `<w:dstrike w:val="true"/>`, false},
		{"outline on", `<w:outline/>`, false},
		{"shadow on", `<w:shadow/>`, false},
		{"emboss on", `<w:emboss/>`, false},
		{"imprint on", `<w:imprint/>`, false},
		{"specVanish on", `<w:specVanish/>`, false},
		{"effect animated", `<w:effect w:val="blinkBackground"/>`, false},
		{"effect absent value", `<w:effect/>`, false},
		{"em dot", `<w:em w:val="dot"/>`, false},
		{"em absent value", `<w:em/>`, false},
		{"spacing expanded", `<w:spacing w:val="15"/>`, false},
		{"spacing condensed", `<w:spacing w:val="-10"/>`, false},
		{"spacing absent value", `<w:spacing/>`, false},
		{"spacing padded zero", `<w:spacing w:val="00"/>`, false},
		{"position raised", `<w:position w:val="6"/>`, false},
		{"position lowered", `<w:position w:val="-6"/>`, false},
		{"shd filled", `<w:shd w:val="clear" w:color="auto" w:fill="000080"/>`, false},
		{"shd patterned", `<w:shd w:val="pct20" w:color="auto" w:fill="auto"/>`, false},
		{"shd unfilled", `<w:shd w:val="clear" w:color="auto"/>`, false},
		// Malformed, decorated, repeated markup states something else.
		{"invalid toggle", `<w:caps w:val="maybe"/>`, false},
		{"unknown attribute", `<w:caps w:val="0" w:other="1"/>`, false},
		{"nested markup", `<w:caps w:val="0"><w:b/></w:caps>`, false},
		{"text content", `<w:caps w:val="0">x</w:caps>`, false},
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
			if got := hasUnsupportedCode(doc, "RUN_FORMATTING_ABSENCE_PRESERVED"); got != tc.accepted {
				t.Fatalf("extract RUN_FORMATTING_ABSENCE_PRESERVED=%v, want %v: %#v", got, tc.accepted, doc.Unsupported)
			}
			if got := hasResolutionDiagnostic(resolved, "RUN_FORMATTING_ABSENCE_PRESERVED"); got != tc.accepted {
				t.Fatalf("resolve RUN_FORMATTING_ABSENCE_PRESERVED=%v, want %v: %#v", got, tc.accepted, resolved.Diagnostics)
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
// helper never qualifies a property that merely shares an attribute shape.
func TestNativeAbsentRunFormattingHelper(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		for markup, want := range map[string]bool{
			`<w:caps w:val="0"/>`:             true,
			`<w:spacing w:val="0"/>`:          true,
			`<w:caps/>`:                       false,
			`<w:b w:val="0"/>`:                false,
			`<w:vanish w:val="0"/>`:           false,
			`<w:vertAlign w:val="baseline"/>`: false,
			// A w:bdr that states no border is the neighbouring
			// nativeAbsentRunBorder reading, never this one.
			`<w:bdr w:val="none" w:sz="0" w:space="0"/>`: false,
			`<w:highlight w:val="none"/>`:                false,
			`<w:spacing w:after="0" w:line="240"/>`:      false,
		} {
			root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+ns+`">`+markup+`</w:rPr>`))
			if err != nil {
				t.Fatal(err)
			}
			if got := nativeAbsentRunFormatting(root.Children[0], root, ns); got != want {
				t.Fatalf("%s: qualified=%v, want %v", markup, got, want)
			}
		}
	}
}

// A foreign-namespace element of the same local name is not this property.
func TestNativeAbsentRunFormattingRejectsForeignNamespace(t *testing.T) {
	root, err := parseNativeXML("test.xml", []byte(`<w:rPr xmlns:w="`+wordMLTransitional+`" xmlns:x="urn:foreign"><x:caps w:val="0" xmlns:w="`+wordMLTransitional+`"/></w:rPr>`))
	if err != nil {
		t.Fatal(err)
	}
	if nativeAbsentRunFormatting(root.Children[0], root, wordMLTransitional) {
		t.Fatal("foreign caps was qualified")
	}
	if !strings.HasPrefix(root.Children[0].Name.Space, "urn:") {
		t.Fatalf("fixture namespace: %q", root.Children[0].Name.Space)
	}
}
