package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// nativeLevelTextFixture adds an explicit, fully modeled presentation lvl1pPr
// and strips the slide runs down to bold/language so the exact tier cannot
// complete them without the presentation level.
func nativeLevelTextFixture(t *testing.T, strict bool, localLevel bool, ambiguous bool, customize ...func(map[string]string)) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		drawing := nsDrawingTransitional
		if strict {
			drawing = nsDrawingStrict
		}
		style := `<p:defaultTextStyle xmlns:a="` + drawing + `"><a:lvl1pPr algn="l" marL="0" indent="0"><a:buNone/><a:defRPr b="0" i="0" sz="2000"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle>`
		parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`, style+`</p:presentation>`, 1)
		part := "relocated/slides/slide-a.xml"
		slide := parts[part]
		slide = strings.Replace(slide, `<a:pPr algn="ctr" lvl="0"><a:buNone/></a:pPr>`, `<a:pPr lvl="0"/>`, 1)
		slide = strings.Replace(slide, `<a:rPr b="1" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr>`, `<a:rPr b="1" lang="tr-TR"/>`, 1)
		slide = strings.Replace(slide, `<a:rPr b="0" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr>`, `<a:rPr lang="en-US"/>`, 1)
		list := `<a:lstStyle>`
		if localLevel {
			list += `<a:lvl1pPr algn="r"><a:defRPr sz="2800"/></a:lvl1pPr>`
		}
		if ambiguous {
			list += `<a:defPPr algn="r"><a:defRPr sz="2600"/></a:defPPr>`
		}
		list += `</a:lstStyle>`
		parts[part] = strings.Replace(slide, `<a:lstStyle/>`, list, 1)
		for _, update := range customize {
			update(parts)
		}
	}})
}

// requireNativeTextRefusedOrUnsupported accepts the two shapes a refused text
// box takes: a retained refused element without paragraphs, or a shape-level
// refusal that leaves the slide with an unsupported-shape passthrough.
func requireNativeTextRefusedOrUnsupported(t *testing.T, deck NativePPTXDeck, context string) {
	t.Helper()
	if len(deck.Slides[0].Elements) == 0 {
		for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
			if diagnostic.Code == "pptx.unsupported-shape" {
				return
			}
		}
		t.Fatalf("%s: shape dropped without a refusal", context)
	}
	e := deck.Slides[0].Elements[0]
	if e.Compatibility.Status != NativeCompatibilityStatusRefused || len(*e.Paragraphs) != 0 {
		t.Fatalf("%s: unqualified or malformed default markup painted: %#v", context, e.Compatibility)
	}
}

func TestNativePresentationLevelMetadataRemainsStrict(t *testing.T) {
	for _, item := range []struct{ name, from, to string }{
		{"direct text", `<a:lvl1pPr algn="l"`, `unmodeled<a:lvl1pPr algn="l"`},
		{"malformed inherited size", `sz="2000"`, `sz="bad"`},
		{"malformed inherited kerning", `sz="2000"`, `sz="2000" kern="bad"`},
		{"duplicate level", `</p:defaultTextStyle>`, `<a:lvl1pPr/></p:defaultTextStyle>`},
		{"unmodeled unused level", `</p:defaultTextStyle>`, `<a:lvl2pPr unsupported="1"/></p:defaultTextStyle>`},
		{"unmodeled layout attribute", `<a:lvl1pPr algn="l"`, `<a:lvl1pPr defTabSz="914400" algn="l"`},
		{"presentation defPPr", `</p:defaultTextStyle>`, `<a:defPPr><a:defRPr sz="2600"/></a:defPPr></p:defaultTextStyle>`},
		{"presentation attribute", `<p:defaultTextStyle `, `<p:defaultTextStyle unsupported="1" `},
	} {
		t.Run(item.name, func(t *testing.T) {
			original := nativeLevelTextFixture(t, false, true, false, func(parts map[string]string) {
				if !strings.Contains(parts["relocated/deck.xml"], item.from) {
					t.Fatal("fixture drifted")
				}
				parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], item.from, item.to, 1)
			})
			deck, err := ExtractNativePPTX(original, nativeTestExtractOptions())
			if err != nil {
				return
			}
			requireNativeTextRefusedOrUnsupported(t, deck, item.name)
		})
	}
}

func TestNativePresentationMatchingLevelStyles(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, local := range []bool{false, true} {
			original := nativeLevelTextFixture(t, strict, local, false)
			before := bytes.Clone(original)
			deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			e := deck.Slides[0].Elements[0]
			p := (*e.Paragraphs)[0]
			size := int64(2000)
			align := NativeTextAlignLeft
			if local {
				size = 2800
				align = NativeTextAlignRight
			}
			if e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || *p.Align != align || len(p.Runs) != 2 || *p.Runs[0].FontSizeHundredthPt != size || *p.Runs[1].FontSizeHundredthPt != size || *p.Runs[0].FontFamily != "Arial" || *p.Runs[0].Color != "112233" || !*p.Runs[0].Bold || *p.Runs[1].Bold || *p.Runs[0].Language != "tr-TR" {
				t.Fatalf("strict=%v local=%v: level cascade lost: %+v %+v", strict, local, p, e.Compatibility)
			}
			codes := nativeDiagnosticCodes(e)
			if codes[nativePresentationTextStylePreviewCode] != 1 || codes[nativeInheritedTextPreviewCode] != 0 {
				t.Fatalf("strict=%v local=%v: projection not disclosed exactly once: %+v", strict, local, e.Compatibility.Diagnostics)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid projected deck: %#v", issues)
			}
			paragraphs := nativeMutationParagraphs("Replace")
			_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
			if err == nil || !strings.Contains(err.Error(), "preview-only") {
				t.Fatalf("resolved defaults granted editing: %v", err)
			}
			if !bytes.Equal(original, before) {
				t.Fatal("source mutated")
			}
		}
	}
}

// Local PowerPoint probes selected a presentation lvl1 over a local defPPr
// while a local lvl1 won. Do not encode a guessed defPPr source-layer merge as
// an exact projection; only matching levels are admitted. This is a refusal
// policy, not a qualification of PowerPoint's behaviour.
func TestNativePresentationDefPPrAmbiguityRemainsRefused(t *testing.T) {
	for _, strict := range []bool{false, true} {
		deck, err := ExtractNativePPTX(nativeLevelTextFixture(t, strict, false, true), nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		requireNativeTextRefusedOrUnsupported(t, deck, "local defPPr")
	}
}

// Self-contained text is exact and editable regardless of presentation
// defaults, including the defPPr and layout attributes PowerPoint writes.
func TestNativePresentationLevelsDoNotDemoteSelfContainedText(t *testing.T) {
	for _, style := range []string{
		`<a:lvl1pPr algn="l" marL="0" indent="0"><a:buNone/><a:defRPr b="0" i="0" sz="2000"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr>`,
		`<a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>`,
	} {
		original := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`, `<p:defaultTextStyle xmlns:a="`+nsDrawingTransitional+`">`+style+`</p:defaultTextStyle></p:presentation>`, 1)
		}})
		deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		e := deck.Slides[0].Elements[0]
		if e.Compatibility.Status != NativeCompatibilityStatusEditable || len(e.Compatibility.Diagnostics) != 0 || len(*e.Paragraphs) != 1 || *(*e.Paragraphs)[0].Runs[0].FontSizeHundredthPt != 3200 || *(*e.Paragraphs)[0].Align != NativeTextAlignCenter {
			t.Fatalf("self-contained text lost exactness: %+v %+v", e.Compatibility, e.Paragraphs)
		}
		paragraphs := nativeMutationParagraphs("Replace")
		if _, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}}); err != nil {
			t.Fatalf("self-contained text lost editing: %v", err)
		}
	}
}

// Placeholders resolve through the master chain, which this projection does
// not model; their text is never completed from presentation levels.
func TestNativePresentationLevelsSkipPlaceholders(t *testing.T) {
	style := `<p:defaultTextStyle xmlns:a="` + nsDrawingTransitional + `"><a:lvl1pPr algn="l" marL="0" indent="0"><a:buNone/><a:defRPr b="0" i="0" sz="2000"><a:latin typeface="Arial"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill></a:defRPr></a:lvl1pPr></p:defaultTextStyle>`
	for _, incomplete := range []bool{false, true} {
		input := nativePlaceholderFixture(t, false, func(parts map[string]string) {
			parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`, style+`</p:presentation>`, 1)
			if incomplete {
				// Without the master typeface the chain cannot complete the run.
				master := strings.Replace(parts["relocated/masters/master.xml"], `<a:latin typeface="+mn-lt"/>`, ``, 1)
				if master == parts["relocated/masters/master.xml"] {
					t.Fatal("fixture drifted")
				}
				parts["relocated/masters/master.xml"] = master
			}
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		for _, e := range deck.Slides[0].Elements {
			if nativeDiagnosticCodes(e)[nativePresentationTextStylePreviewCode] != 0 {
				t.Fatalf("incomplete=%v: placeholder text was projected from presentation levels: %+v", incomplete, e.Compatibility.Diagnostics)
			}
		}
		if incomplete {
			requireNativeTextRefusedOrUnsupported(t, deck, "incomplete placeholder")
			continue
		}
		if len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("placeholder missing: %+v", deck.Slides[0].Compatibility)
		}
		e := deck.Slides[0].Elements[0]
		if e.Placeholder == nil || nativeDiagnosticCodes(e)["pptx.inherited-placeholder-preview"] != 1 || *(*e.Paragraphs)[0].Runs[0].FontSizeHundredthPt != 2400 || *(*e.Paragraphs)[0].Runs[0].FontFamily != "Calibri" {
			t.Fatalf("placeholder chain changed under presentation defaults: %+v %+v", e.Compatibility.Diagnostics, e.Paragraphs)
		}
	}
}

// The opt-in inherited preview keeps its own composed layers and disclosure;
// the exact projection never runs beneath it.
func TestNativePresentationLevelsYieldToInheritedPreview(t *testing.T) {
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(nativeLevelTextFixture(t, false, true, false), options)
	if err != nil {
		t.Fatal(err)
	}
	e := deck.Slides[0].Elements[0]
	codes := nativeDiagnosticCodes(e)
	if codes[nativePresentationTextStylePreviewCode] != 0 || codes[nativeInheritedTextPreviewCode] != 1 || len(*e.Paragraphs) != 1 || *(*e.Paragraphs)[0].Runs[0].FontSizeHundredthPt != 2800 {
		t.Fatalf("inherited preview lane changed: %+v %+v", e.Compatibility.Diagnostics, e.Paragraphs)
	}
}
