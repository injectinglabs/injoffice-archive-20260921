package pptxpatch

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func nativeStyledTextFixture(t *testing.T, strict bool, marker string, customize func(map[string]string)) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		drawing, presentation := nsDrawingTransitional, nsPresentationTransitional
		if strict {
			drawing, presentation = nsDrawingStrict, nsPresentationStrict
		}
		parts["relocated/themes/theme.xml"] = nativeExactThemeXML(drawing)
		parts["relocated/masters/master.xml"] = nativeExactMasterWithColorMapXML(presentation)
		slide := parts["relocated/slides/slide-a.xml"]
		bullet := `<a:buNone/>`
		if marker != "" {
			bullet = `<a:buChar char="` + marker + `"/>`
		}
		list := `<a:lstStyle><a:defPPr algn="r"><a:defRPr b="0" i="1" sz="1800"><a:latin typeface="+mn-lt"/><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:defRPr></a:defPPr><a:lvl1pPr algn="l" marL="300000" indent="-100000">` + bullet + `<a:defRPr sz="2000"/></a:lvl1pPr></a:lstStyle>`
		slide = strings.Replace(slide, `<a:lstStyle/>`, list, 1)
		slide = strings.Replace(slide, `<a:pPr algn="ctr" lvl="0"><a:buNone/></a:pPr>`, `<a:pPr lvl="0"><a:defRPr sz="2400"/></a:pPr>`, 1)
		slide = strings.Replace(slide, `<a:rPr b="1" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr>`, `<a:rPr b="1"/>`, 1)
		slide = strings.Replace(slide, `<a:rPr b="0" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr>`, ``, 1)
		parts["relocated/slides/slide-a.xml"] = slide
		if customize != nil {
			customize(parts)
		}
	}})
}

func TestNativeLocalStyleCascadeAndAuthoredBullet(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, marker := range []string{"", "▪", "🙂"} {
			input := nativeStyledTextFixture(t, strict, marker, nil)
			before := append([]byte(nil), input...)
			deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusEditable || element.Paragraphs == nil {
				t.Fatalf("style cascade refused: %+v", element)
			}
			paragraph := (*element.Paragraphs)[0]
			if *paragraph.Align != NativeTextAlignLeft || *paragraph.MarginLeftEmu != 300000 || *paragraph.IndentEmu != -100000 || *paragraph.Bullet != (marker != "") {
				t.Fatalf("paragraph defaults not resolved: %+v", paragraph)
			}
			if marker != "" && (paragraph.BulletCharacter == nil || *paragraph.BulletCharacter != marker) {
				t.Fatalf("authored marker was replaced: %+v", paragraph)
			}
			for i, run := range paragraph.Runs {
				if *run.FontSizeHundredthPt != 2400 || *run.Italic != true || *run.FontFamily != "Calibri" || *run.Color != "2F6FED" || *run.Bold != (i == 0) {
					t.Fatalf("run precedence wrong: %+v", run)
				}
			}
			if !bytes.Equal(before, input) {
				t.Fatal("rendering cascade mutated source package")
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatalf("invalid native result: %+v", issues)
			}
		}
	}
}

func nativeTextCheckingFixture(t *testing.T, strict bool, metadata string) []byte {
	return nativeStyledTextFixture(t, strict, "▪", func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		parts[part] = strings.Replace(parts[part], `<a:rPr b="1"/>`, `<a:rPr b="1" `+metadata+`/>`, 1)
	})
}

func TestNativeTextCheckingFlagsPreviewWithoutMutationPermission(t *testing.T) {
	for _, strict := range []bool{false, true} {
		baseline, err := ExtractNativePPTX(nativeStyledTextFixture(t, strict, "▪", nil), nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		for _, metadata := range []string{`dirty="0" smtClean="1"`, `dirty="true" smtClean="false"`} {
			original := nativeTextCheckingFixture(t, strict, metadata)
			before := bytes.Clone(original)
			deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.TextBody == nil || element.Paragraphs == nil {
				t.Fatalf("checking flags did not yield a read-only paint projection: %+v", element)
			}
			if !nativeParagraphsEqual(*baseline.Slides[0].Elements[0].Paragraphs, *element.Paragraphs) {
				t.Fatal("proofing flags changed glyph inputs")
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid contract: %+v", issues)
			}
			paragraphs := nativeMutationParagraphs("replacement")
			_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
			if err == nil {
				t.Fatal("preview-only flags allowed lossy text replacement")
			}
			if !bytes.Equal(original, before) {
				t.Fatal("preview/refused replacement changed original bytes")
			}
		}
	}
}

func TestNativeTextCheckingFlagsRemainStrict(t *testing.T) {
	for _, metadata := range []string{`dirty="yes"`, `smtClean="2"`, `dirty="0" dirty="1"`, `kumimoji="1"`, `lang="en-US"`, `unknown="0"`, `x:dirty="0" xmlns:x="urn:other"`} {
		deck, err := ExtractNativePPTX(nativeTextCheckingFixture(t, false, metadata), nativeTestExtractOptions())
		if err == nil {
			for _, element := range deck.Slides[0].Elements {
				if element.Paragraphs != nil && len(*element.Paragraphs) > 0 {
					t.Fatalf("unmodeled metadata painted: %s", metadata)
				}
			}
		}
	}
	input := nativeStyledTextFixture(t, false, "▪", func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		parts[part] = strings.Replace(parts[part], `<a:defRPr sz="2400"/>`, `<a:defRPr sz="2400" dirty="invalid"/>`, 1)
		parts[part] = strings.Replace(parts[part], `<a:rPr b="1"/>`, `<a:rPr b="1" dirty="0"/>`, 1)
	})
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err == nil && len(deck.Slides[0].Elements) > 0 && len(*deck.Slides[0].Elements[0].Paragraphs) > 0 {
		t.Fatal("local override hid malformed inherited checking flag")
	}
}

func TestNativeTextCheckingLocalFlagsCannotBeLostBehindDefaults(t *testing.T) {
	for _, strict := range []bool{false, true} {
		input := nativeStyledTextFixture(t, strict, "", func(parts map[string]string) {
			part := "relocated/slides/slide-a.xml"
			parts[part] = strings.Replace(parts[part], `<a:defRPr b="0"`, `<a:defRPr dirty="0" b="0"`, 1)
			parts[part] = strings.Replace(parts[part], `<a:rPr b="1"/>`, `<a:rPr b="1" smtClean="0"/>`, 1)
		})
		before := bytes.Clone(input)
		deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		element := deck.Slides[0].Elements[0]
		paragraphs := nativeMutationParagraphs("Replacement")
		_, err = ApplyNativePPTXMutations(input, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
		if err == nil || !strings.Contains(err.Error(), "text checking metadata is preserve-only") {
			t.Fatalf("expected explicit prewrite checking flag refusal, got %v", err)
		}
		if !bytes.Equal(input, before) {
			t.Fatal("refused mutation changed source bytes")
		}
	}
}

func TestNativeTextCheckingFlagsInShapesAndTablesRemainPreserveOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		cell := strings.Replace(nativeExactTableCellXML("Checking flags", "l", "FFFFFF"), `<a:rPr `, `<a:rPr dirty="0" smtClean="1" `, 1)
		table := nativeExactTableGraphicFrameXML(3, "Checking table", []int64{1000000}, []int64{300000}, [][]string{{cell}}, "")
		textStart, textEnd := strings.Index(cell, "<a:txBody>"), strings.Index(cell, "</a:txBody>")+len("</a:txBody>")
		body := strings.ReplaceAll(cell[textStart:textEnd], "a:txBody", "p:txBody")
		shape := nativeAutoShapeXML(4, "Checking shape", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
		shape = strings.Replace(shape, "</p:sp>", body+"</p:sp>", 1)
		deck, err := ExtractNativePPTX(nativeTableFixture(t, strict, table+shape), nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if len(deck.Slides[0].Elements) != 3 {
			t.Fatal("lost sibling elements")
		}
		for _, element := range deck.Slides[0].Elements[1:] {
			if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("checking metadata widened mutation authority: %+v", element)
			}
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid projection: %+v", issues)
		}
	}
}

func TestNativeTextStyleRefusesUnmodeledOrAmbiguousDefaults(t *testing.T) {
	for _, mutation := range []struct{ from, to string }{
		{`<a:defRPr sz="2400"/>`, `<a:defRPr sz="2400" u="sng"/>`},
		{`<a:rPr b="1"/>`, `<a:rPr b="1"><a:highlight/></a:rPr>`},
		{`<a:buChar char="▪"/>`, `<a:buChar char="▪"/><a:buNone/>`},
		{`<a:buChar char="▪"/>`, `<a:buAutoNum type="arabicPeriod"/>`},
		{`<a:buChar char="▪"/>`, `<a:buChar char="two"/>`},
		{`<a:buChar char="▪"/>`, `<a:buChar char="▪"/><a:buFont typeface="Symbol"/>`},
		{`sz="1800"`, `sz="invalid"`},
		{`b="0" i="1"`, `b="invalid" i="1"`},
		{`algn="r"`, `algn="invalid"`},
		{`<a:buChar char="▪"/>`, `<a:buChar char="▪" unknown="1"/>`},
	} {
		input := nativeStyledTextFixture(t, false, "▪", func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], mutation.from, mutation.to, 1)
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err == nil && (deck.Compatibility.Status == NativeCompatibilityStatusEditable || (len(deck.Slides[0].Elements) > 0 && deck.Slides[0].Elements[0].Compatibility.Status == NativeCompatibilityStatusEditable)) {
			t.Fatalf("unmodeled source was approximated: %s", mutation.to)
		}
	}
}

func TestNativeTextStyleCannotHideMalformedOverriddenLeaf(t *testing.T) {
	for _, leaf := range []struct{ valid, invalid string }{
		{`<a:latin typeface="+mn-lt"/>`, `<a:latin typeface="+mn-lt" unknown="1"/>`},
		{`<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`, `<a:solidFill><a:schemeClr val="accent1"><a:unknown/></a:schemeClr></a:solidFill>`},
	} {
		input := nativeStyledTextFixture(t, false, "▪", func(parts map[string]string) {
			part := "relocated/slides/slide-a.xml"
			source := strings.Replace(parts[part], leaf.valid, leaf.invalid, 1)
			source = strings.Replace(source, `<a:rPr b="1"/>`, `<a:rPr b="1">`+leaf.valid+`</a:rPr>`, 1)
			source = strings.Replace(source, `<a:r><a:t>world`, `<a:r><a:rPr>`+leaf.valid+`</a:rPr><a:t>world`, 1)
			parts[part] = source
		})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			continue
		}
		for _, element := range deck.Slides[0].Elements {
			if element.Compatibility.Status == NativeCompatibilityStatusEditable {
				t.Fatal("malformed overridden leaf became editable")
			}
		}
	}
}

// The Chrome integration may request these rights-cleared source bytes and the
// real extractor projection. Normal Go tests never write repository fixtures.
func TestNativeTextStyleBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_STYLE_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional real-file browser fixture export")
	}
	input := nativeTextCheckingFixture(t, false, `dirty="0" smtClean="0"`)
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	output, err := MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "styled-native.pptx"), input, 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "styled-native.json"), output, 0600); err != nil {
		t.Fatal(err)
	}
	// Separate native-paint fixture: no bullet marker/hanging-indent dependency.
	checking := nativeStyledTextFixture(t, false, "", func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		parts[part] = strings.Replace(parts[part], `<a:rPr b="1"/>`, `<a:rPr b="1" dirty="0" smtClean="0"/>`, 1)
	})
	if err := os.WriteFile(filepath.Join(dir, "checking-native.pptx"), checking, 0600); err != nil {
		t.Fatal(err)
	}
	input = nativePlaceholderFixture(t, false, nil)
	deck, err = ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	output, err = MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "placeholder-native.pptx"), input, 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "placeholder-native.json"), output, 0600); err != nil {
		t.Fatal(err)
	}
}
