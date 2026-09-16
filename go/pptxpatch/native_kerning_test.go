package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// nativeKerningFixture places an authored kerning threshold on the first run's
// direct a:rPr, or on a local a:lstStyle level defRPr that the run inherits.
func nativeKerningFixture(t *testing.T, strict bool, value string, level bool) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		if level {
			parts[part] = strings.Replace(parts[part], `<a:lstStyle/>`, `<a:lstStyle><a:lvl1pPr><a:defRPr kern="`+value+`"/></a:lvl1pPr></a:lstStyle>`, 1)
			return
		}
		parts[part] = strings.Replace(parts[part], `<a:rPr b="1"`, `<a:rPr kern="`+value+`" b="1"`, 1)
	}})
}

func TestNativeKerningThresholdExtractedFromRunAndLevelStyles(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, level := range []bool{false, true} {
			for _, value := range []string{"0", "1200", "400000"} {
				original := nativeKerningFixture(t, strict, value, level)
				before := bytes.Clone(original)
				deck, err := ExtractNativePPTX(original, nativeTestExtractOptions())
				if err != nil {
					t.Fatalf("strict=%v level=%v kern=%s: %v", strict, level, value, err)
				}
				e := deck.Slides[0].Elements[0]
				if e.Compatibility.Status != NativeCompatibilityStatusEditable || len(e.Compatibility.Diagnostics) != 0 {
					t.Fatalf("strict=%v level=%v kern=%s: kerned text is not exact: %#v", strict, level, value, e.Compatibility)
				}
				runs := (*e.Paragraphs)[0].Runs
				if len(runs) != 2 {
					t.Fatalf("strict=%v level=%v kern=%s: runs lost: %#v", strict, level, value, runs)
				}
				first := runs[0]
				if first.KerningThresholdHundredthPt == nil || *first.KerningThresholdHundredthPt != mustParseNativeInt(t, value) {
					t.Fatalf("strict=%v level=%v kern=%s: threshold lost: %#v", strict, level, value, first)
				}
				if *first.FontSizeHundredthPt != 3200 || !*first.Bold || *first.FontFamily != "Aptos" || *first.Text != "Hello " {
					t.Fatalf("strict=%v level=%v kern=%s: sibling run properties changed: %#v", strict, level, value, first)
				}
				second := runs[1]
				if level {
					// The level default reaches every run of that level.
					if second.KerningThresholdHundredthPt == nil || *second.KerningThresholdHundredthPt != mustParseNativeInt(t, value) {
						t.Fatalf("strict=%v kern=%s: level threshold did not cascade to the second run: %#v", strict, value, second)
					}
				} else if second.KerningThresholdHundredthPt != nil {
					t.Fatalf("strict=%v kern=%s: threshold leaked to an unkerned run: %#v", strict, value, second)
				}
				if issues := ValidateNativePPTX(deck); len(issues) != 0 {
					t.Fatalf("strict=%v level=%v kern=%s: invalid deck: %#v", strict, level, value, issues)
				}
				if !bytes.Equal(before, original) {
					t.Fatal("original bytes changed")
				}
			}
		}
	}
}

func TestNativeKerningThresholdLocalRunOverridesLevelStyle(t *testing.T) {
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		slide := strings.Replace(parts[part], `<a:lstStyle/>`, `<a:lstStyle><a:lvl1pPr><a:defRPr kern="1200"/></a:lvl1pPr></a:lstStyle>`, 1)
		parts[part] = strings.Replace(slide, `<a:rPr b="1"`, `<a:rPr kern="0" b="1"`, 1)
	}})
	deck, err := ExtractNativePPTX(original, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	runs := (*deck.Slides[0].Elements[0].Paragraphs)[0].Runs
	if runs[0].KerningThresholdHundredthPt == nil || *runs[0].KerningThresholdHundredthPt != 0 {
		t.Fatalf("direct zero threshold did not override the level default: %#v", runs[0])
	}
	if runs[1].KerningThresholdHundredthPt == nil || *runs[1].KerningThresholdHundredthPt != 1200 {
		t.Fatalf("level threshold lost on the run without a direct value: %#v", runs[1])
	}
}

func TestNativeKerningThresholdInvalidSourceRefused(t *testing.T) {
	for _, level := range []bool{false, true} {
		for _, value := range []string{"-1", "400001", "1.5", "01", "+1", "bad", "", "9223372036854775808"} {
			deck, err := ExtractNativePPTX(nativeKerningFixture(t, false, value, level), nativeTestExtractOptions())
			if err != nil {
				continue
			}
			if len(deck.Slides[0].Elements) == 0 {
				// A malformed direct run attribute refuses the whole shape into
				// slide passthrough, like every other malformed run attribute.
				unsupported := false
				for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
					unsupported = unsupported || diagnostic.Code == "pptx.unsupported-shape"
				}
				if !unsupported {
					t.Fatalf("level=%v: invalid kerning threshold %q dropped the shape without a refusal", level, value)
				}
				continue
			}
			e := deck.Slides[0].Elements[0]
			if e.Compatibility.Status != NativeCompatibilityStatusRefused || len(*e.Paragraphs) != 0 {
				t.Fatalf("level=%v: invalid kerning threshold %q was painted: %#v", level, value, e.Compatibility)
			}
		}
	}
	// A valid direct value must not hide a malformed inherited level value.
	deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		slide := strings.Replace(parts[part], `<a:lstStyle/>`, `<a:lstStyle><a:lvl1pPr><a:defRPr kern="bad"/></a:lvl1pPr></a:lstStyle>`, 1)
		parts[part] = strings.Replace(slide, `<a:rPr b="1"`, `<a:rPr kern="1200" b="1"`, 1)
	}}), nativeTestExtractOptions())
	if err == nil && len(deck.Slides[0].Elements) != 0 && deck.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatal("malformed inherited threshold hidden by a valid local value")
	}
}

func TestNativeKerningThresholdMutationPolicy(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := nativeKerningFixture(t, strict, "1200", false)
		before := bytes.Clone(original)
		deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		e := deck.Slides[0].Elements[0]
		if got := (*e.Paragraphs)[0].Runs[0].KerningThresholdHundredthPt; got == nil || *got != 1200 {
			t.Fatal("source threshold lost under mutation extract options")
		}
		// Extracted paragraphs carry the threshold, so writing them back
		// unchanged is refused rather than silently dropping the kerning.
		echoed := *e.Paragraphs
		request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &echoed}}}
		if _, err := ApplyNativePPTXMutations(original, request); err == nil || !strings.Contains(err.Error(), "kerning") {
			t.Fatalf("echoed kerning threshold was not refused explicitly: %v", err)
		}
		// Replacement text without a threshold still targets the element.
		paragraphs := nativeMutationParagraphs("AV")
		request.Operations[0].Paragraphs = &paragraphs
		result, err := ApplyNativePPTXMutations(original, request)
		if err != nil {
			t.Fatal(err)
		}
		reopened, err := ExtractNativePPTX(result, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		got := (*reopened.Slides[0].Elements[0].Paragraphs)[0].Runs[0]
		if *got.Text != "AV" || got.KerningThresholdHundredthPt != nil {
			t.Fatalf("replacement run did not round trip without kerning: %#v", got)
		}
		if !bytes.Equal(before, original) {
			t.Fatal("original bytes changed")
		}
	}
}

func mustParseNativeInt(t *testing.T, value string) int64 {
	t.Helper()
	parsed, err := parseCanonicalNativeInt(value, 0, 400000)
	if err != nil {
		t.Fatal(err)
	}
	return parsed
}
