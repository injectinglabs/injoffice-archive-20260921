package docxpatch

import (
	"strings"
	"testing"
)

// w:spacing on w:rPr is character tracking (ECMA-376 17.3.2.35): one signed
// whole-twip measurement added to the advance of every character of the run.
// It resolves through the same cascade as every other run-property measurement
// - docDefaults, then the paragraph style chain, then the character style, then
// the run's own w:rPr - and the nearest layer that states one wins outright.
func TestNativeCharacterSpacingCascade(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		styles := `<w:styles xmlns:w="` + ns + `"><w:docDefaults><w:rPrDefault><w:rPr><w:spacing w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:spacing w:val="30"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:type="character" w:styleId="Char"><w:rPr><w:spacing w:val="-20"/></w:rPr></w:style></w:styles>`
		parts := resolvedStylesTestParts(styles)
		parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/><w:rPr><w:spacing w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Char"/><w:spacing w:val="15"/></w:rPr><w:t>AV</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Char"/></w:rPr><w:t>AV</w:t></w:r><w:r><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
		if ns == wordMLStrict {
			for k, v := range parts {
				parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
			}
		}
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := string(data)
		layout, err := ResolveNativeDocumentLayoutV1(data)
		if err != nil {
			t.Fatal(err)
		}
		// run0 states its own, run1 inherits the character style, run2 the
		// paragraph style chain - never the docDefaults it overrides.
		for i, want := range []int{15, -20, 30} {
			got := layout.Runs[i].Properties.LetterSpacingTwips
			if got == nil || *got != want {
				t.Fatalf("%s run%d: %v want %d", ns, i, got, want)
			}
		}
		mark := layout.Paragraphs[0].ParagraphMarkProperties.LetterSpacingTwips
		if mark == nil || *mark != 18 {
			t.Fatalf("%s mark %v", ns, mark)
		}
		if string(data) != before {
			t.Fatal("source changed")
		}
	}
}

// An authored zero keeps exactly the treatment it had before tracking was
// modeled: it states the absence of its own effect, so it resolves to no
// tracking at all and is reported as preserved, not applied. Nothing about a
// document whose only w:spacing is zero may move.
func TestNativeCharacterSpacingZeroStaysInert(t *testing.T) {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr><w:spacing w:val="0"/></w:rPr><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
	layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if layout.Runs[0].Properties.LetterSpacingTwips != nil {
		t.Fatal("an authored zero resolved to a tracking value")
	}
	if !hasNativeDiagnosticCode(layout, "RUN_EFFECT_ABSENT_PRESERVED") {
		t.Fatal("an authored zero must stay reported as an absent run effect")
	}
	if hasNativeDiagnosticCode(layout, "UNMODELED_RUN_PROPERTY") {
		t.Fatal("an authored zero must not be reported as unmodeled")
	}
}

// Everything this tier cannot read as one exact whole-twip measurement stays
// unmodeled markup and keeps refusing, rather than being guessed into geometry.
func TestNativeCharacterSpacingInvalidSource(t *testing.T) {
	for _, value := range []string{``, `w:val=""`, `w:val="15pt"`, `w:val="+15"`, `w:val="015"`, `w:val="1.5"`, `w:val="31681"`, `w:val="-31681"`, `w:val="15" extra="1"`} {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr><w:spacing ` + value + `/></w:rPr><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
		layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if layout.Runs[0].Properties.LetterSpacingTwips != nil {
			t.Fatalf("%q resolved to a tracking value", value)
		}
		if !hasNativeDiagnosticCode(layout, "UNMODELED_RUN_PROPERTY") {
			t.Fatalf("%q must stay preserved and unmodeled", value)
		}
	}
	// Two w:spacing children make the layer ambiguous, so neither is resolved.
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr><w:spacing w:val="15"/><w:spacing w:val="20"/></w:rPr><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
	layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if layout.Runs[0].Properties.LetterSpacingTwips != nil {
		t.Fatal("a duplicated w:spacing resolved to a tracking value")
	}
}

func hasNativeDiagnosticCode(layout *NativeResolvedLayoutInputV1, code string) bool {
	for _, diagnostic := range layout.Diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}
