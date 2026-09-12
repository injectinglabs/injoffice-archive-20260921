package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeSpellingErrorFlagRetainsSourceGuards(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, flag := range []string{`err="yes"`, `err="0" err="1"`, `x:err="1" xmlns:x="urn:other"`} {
			deck, err := ExtractNativePPTX(nativeTextCheckingFixture(t, strict, flag), nativeTestExtractOptions())
			if err == nil {
				for _, element := range deck.Slides[0].Elements {
					if element.Paragraphs != nil && len(*element.Paragraphs) > 0 {
						t.Fatalf("unsafe flag painted: %s", flag)
					}
				}
			}
		}
		input := nativeStyledTextFixture(t, strict, "", func(parts map[string]string) {
			part := "relocated/slides/slide-a.xml"
			parts[part] = strings.Replace(parts[part], `<a:defRPr b="0"`, `<a:defRPr dirty="0" b="0"`, 1)
			parts[part] = strings.Replace(parts[part], `<a:rPr b="1"/>`, `<a:rPr b="1" err="1"/>`, 1)
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
			t.Fatalf("lost local err behind unchanged default warning: %v", err)
		}
		if !bytes.Equal(before, input) {
			t.Fatal("source changed")
		}
		invalid := nativeStyledTextFixture(t, strict, "", func(parts map[string]string) {
			part := "relocated/slides/slide-a.xml"
			parts[part] = strings.Replace(parts[part], `<a:defRPr b="0"`, `<a:defRPr err="bad" b="0"`, 1)
			parts[part] = strings.Replace(parts[part], `<a:rPr b="1"/>`, `<a:rPr b="1" err="0"/>`, 1)
		})
		deck, err = ExtractNativePPTX(invalid, nativeTestExtractOptions())
		if err == nil {
			for _, element := range deck.Slides[0].Elements {
				if element.Paragraphs != nil && len(*element.Paragraphs) > 0 {
					t.Fatal("valid local err hid malformed inherited flag")
				}
			}
		}
	}
}
