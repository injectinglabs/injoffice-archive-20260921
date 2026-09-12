package pptxpatch

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func nativeLanguageFixture(t *testing.T, strict bool) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		text := strings.ReplaceAll(parts[part], `typeface="Aptos"`, `typeface="Calibri"`)
		text = strings.Replace(text, `<a:rPr b="1"`, `<a:rPr lang="tr-TR" b="1"`, 1)
		text = strings.Replace(text, `<a:rPr b="0"`, `<a:rPr lang="en-US" b="0"`, 1)
		parts[part] = text
	}})
}

func TestNativeAuthoredLanguageMutationRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := nativeLanguageFixture(t, strict)
		before := bytes.Clone(original)
		deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		e := deck.Slides[0].Elements[0]
		runs := (*e.Paragraphs)[0].Runs
		if len(runs) != 2 || runs[0].Language == nil || *runs[0].Language != "tr-TR" || runs[1].Language == nil || *runs[1].Language != "en-US" {
			t.Fatal("source language lost")
		}
		paragraphs := nativeMutationParagraphs("Turkish I")
		language := "tr-TR"
		paragraphs[0].Runs[0].Language = &language
		request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}}
		result, err := ApplyNativePPTXMutations(original, request)
		if err != nil {
			t.Fatal(err)
		}
		reopened, err := ExtractNativePPTX(result, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		got := (*reopened.Slides[0].Elements[0].Paragraphs)[0].Runs[0]
		if got.Language == nil || *got.Language != "tr-TR" || *got.Text != "Turkish I" {
			t.Fatalf("saved language/text did not round trip: %+v", got)
		}
		if !bytes.Equal(before, original) {
			t.Fatal("mutated original input bytes")
		}
		language = `en-US" bad="value`
		if _, err = ApplyNativePPTXMutations(original, request); err == nil {
			t.Fatal("malformed mutation language accepted")
		}
	}
}

func TestNativeLanguageContractAndMutationBounds(t *testing.T) {
	for _, language := range []string{"", "en_US", "en US", "a", "en-123456789", strings.Repeat("en-", 24)} {
		paragraphs := nativeMutationParagraphs("Bounded")
		paragraphs[0].Runs[0].Language = &language
		if err := validateNativeMutationParagraphs(paragraphs, &nativeMutationBudget{}); err == nil {
			t.Fatalf("invalid language accepted: %q", language)
		}
	}
}

func TestNativeLanguageBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_LANGUAGE_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional local native/browser fixture")
	}
	if err := os.WriteFile(filepath.Join(dir, "language.pptx"), nativeLanguageFixture(t, false), 0600); err != nil {
		t.Fatal(err)
	}
}
