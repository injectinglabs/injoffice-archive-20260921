package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeVerticalTextSourceQualification(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, mode := range []string{"vert", "vert270", "eaVert", "wordArtVert", "mongolianVert", "wordArtVertRtl"} {
			input := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
				part := "relocated/slides/slide-a.xml"
				parts[part] = strings.Replace(parts[part], `<a:bodyPr/>`, `<a:bodyPr vert="`+mode+`"/>`, 1)
			}})
			before := bytes.Clone(input)
			deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			e := deck.Slides[0].Elements[0]
			if mode != "vert" {
				if e.TextBody != nil {
					t.Fatalf("accepted mode%s", mode)
				}
				continue
			}
			if e.TextBody == nil || e.TextBody.WritingMode == nil || *e.TextBody.WritingMode != "vertical-clockwise" || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("not source-qualified %#v", e)
			}
			if e.Transform.QuarterTurns != nil {
				t.Fatal("vertical text rotated source shape")
			}
			paragraphs := nativeMutationParagraphs("replacement")
			_, err = ApplyNativePPTXMutations(input, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
			if err == nil {
				t.Fatal("vertical source text became mutable")
			}
			if !bytes.Equal(before, input) {
				t.Fatal("source changed")
			}
			deck.Slides[0].Elements[0].TextBody.WritingMode = stringPointer("stacked")
			if len(ValidateNativePPTX(deck)) == 0 {
				t.Fatal("unknown writing mode accepted")
			}
		}
	}
}
