package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func nativeSourceFrameFixture(t *testing.T, strict bool, body string) []byte {
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		parts[part] = strings.Replace(parts[part], `<a:bodyPr/>`, body, 1)
	}})
}

func TestNativeSourceFrameAutoFitIsExplicitReadOnly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := nativeSourceFrameFixture(t, strict, `<a:bodyPr wrap="none"><a:spAutoFit/></a:bodyPr>`)
		before := bytes.Clone(original)
		options := nativeMutationExtractOptions()
		ordinary, err := ExtractNativePPTX(original, options)
		if err != nil {
			t.Fatal(err)
		}
		if ordinary.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused {
			t.Fatal("strict extraction accepted autofit")
		}
		options.AllowSourceFrameAutoFitPreview = true
		preview, err := ExtractNativePPTX(original, options)
		if err != nil {
			t.Fatal(err)
		}
		e := preview.Slides[0].Elements[0]
		if e.TextBody == nil || e.TextBody.AutoFit != "shape-source-frame" || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(*e.Paragraphs) != 1 {
			t.Fatalf("missing explicit preview: %+v", e)
		}
		if !nativeInt64PointerEqual(e.Transform.Cx, ordinary.Slides[0].Elements[0].Transform.Cx) || !nativeInt64PointerEqual(e.Transform.Cy, ordinary.Slides[0].Elements[0].Transform.Cy) {
			t.Fatal("source frame resized")
		}
		if *preview.SourceRevision != *ordinary.SourceRevision || !bytes.Equal(original, before) {
			t.Fatal("preview changed source identity/bytes")
		}
		if issues := ValidateNativePPTX(preview); len(issues) > 0 {
			t.Fatalf("invalid approximate contract: %+v", issues)
		}
		paragraphs := nativeMutationParagraphs("No mutation authority")
		request := NativePPTXMutationRequest{ExpectedSourceRevision: *preview.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}}
		if _, err := ApplyNativePPTXMutations(original, request); err == nil {
			t.Fatal("approximation granted editing permission")
		}
		preview.Slides[0].Elements[0].Compatibility.Status = NativeCompatibilityStatusEditable
		if len(ValidateNativePPTX(preview)) == 0 {
			t.Fatal("approximation validated as editable")
		}
	}
}

func TestNativeSourceFrameAutoFitDoesNotHideOtherLayoutGaps(t *testing.T) {
	for _, body := range []string{`<a:bodyPr><a:spAutoFit/><a:noAutofit/></a:bodyPr>`, `<a:bodyPr><a:spAutoFit/><a:spAutoFit/></a:bodyPr>`, `<a:bodyPr><a:spAutoFit bad="1"/></a:bodyPr>`, `<a:bodyPr><a:spAutoFit>text</a:spAutoFit></a:bodyPr>`, `<a:bodyPr vert="vert270"><a:spAutoFit/></a:bodyPr>`, `<a:bodyPr><a:normAutofit/></a:bodyPr>`, `<a:bodyPr numCol="2"><a:spAutoFit/></a:bodyPr>`, `<a:bodyPr vertOverflow="clip"><a:spAutoFit/></a:bodyPr>`} {
		options := nativeMutationExtractOptions()
		options.AllowSourceFrameAutoFitPreview = true
		deck, err := ExtractNativePPTX(nativeSourceFrameFixture(t, false, body), options)
		if err == nil && len(deck.Slides[0].Elements) > 0 && deck.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused {
			t.Fatalf("unmodeled semantics hidden by approximate mode: %s", body)
		}
	}
}
