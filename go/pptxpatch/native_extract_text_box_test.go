package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func nativeRectangularTextBoxFixture(t *testing.T, strict bool, properties string, unsupportedText bool) []byte {
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		parts[part] = strings.Replace(parts[part], `</a:xfrm></p:spPr>`, `</a:xfrm>`+properties+`</p:spPr>`, 1)
		if unsupportedText {
			parts[part] = strings.Replace(parts[part], `<a:bodyPr/>`, `<a:bodyPr><a:spAutoFit/></a:bodyPr>`, 1)
			parts[part] = strings.Replace(parts[part], `<a:buNone/>`, `<a:buFont typeface="Wingdings"/><a:buChar char="q"/>`, 1)
		}
	}})
}

func TestNativeTransparentRectangularTextBoxes(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, properties := range []string{`<a:noFill/>`, `<a:prstGeom prst="rect"/><a:noFill/>`, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>`} {
			original := nativeRectangularTextBoxFixture(t, strict, properties, false)
			before := bytes.Clone(original)
			deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if len(deck.Slides[0].Elements) != 1 {
				t.Fatal("transparent text box disappeared")
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusEditable || element.TextBody == nil || element.Paragraphs == nil {
				t.Fatalf("exact text box was not materialized: %#v", element)
			}
			if !bytes.Equal(original, before) {
				t.Fatal("extraction changed source bytes")
			}
			paragraphs := nativeMutationParagraphs("Updated transparent box")
			produced, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
			if err != nil {
				t.Fatal(err)
			}
			after, err := ExtractNativePPTX(produced, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if !nativeParagraphsEqual(*after.Slides[0].Elements[0].Paragraphs, paragraphs) {
				t.Fatal("replacement did not round trip")
			}
			pkg, err := openNativeExtractPackage(produced)
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Contains(pkg.parts["relocated/slides/slide-a.xml"], []byte(properties)) {
				t.Fatal("replacement lost original shape markup")
			}
		}
	}
}

func TestNativeTransparentTextBoxUnsupportedContentRemainsVisibleRefusal(t *testing.T) {
	deck, err := ExtractNativePPTX(nativeRectangularTextBoxFixture(t, false, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/>`, true), nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatal("unsupported text disappeared instead of retaining a bounded placeholder")
	}
	element := deck.Slides[0].Elements[0]
	if element.Compatibility.Status != NativeCompatibilityStatusRefused || len(element.Passthrough) == 0 || element.Source == nil {
		t.Fatalf("unsupported text became editable or lost source: %#v", element)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid refusal contract: %#v", issues)
	}
}

func TestNativeTextBoxRejectsUnmodeledPaint(t *testing.T) {
	for _, properties := range []string{
		`<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom><a:noFill/>`,
		`<a:prstGeom prst="rect"><a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst></a:prstGeom><a:noFill/>`,
		`<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`,
		`<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`,
		`<a:noFill/><a:ln w="12700"><a:noFill/></a:ln>`,
		`<a:noFill/>rogue`,
	} {
		deck, err := ExtractNativePPTX(nativeRectangularTextBoxFixture(t, false, properties, false), nativeTestExtractOptions())
		if err == nil && len(deck.Slides[0].Elements) != 0 {
			t.Fatalf("unmodeled paint accepted as transparent text: %s", properties)
		}
	}
	for _, properties := range []string{`<a:noFill/><a:noFill/>`, `<a:prstGeom prst="rect"/><a:prstGeom prst="rect"/><a:noFill/>`, `<a:prstGeom prst="rect"><a:avLst/><a:avLst/></a:prstGeom><a:noFill/>`} {
		if _, err := ExtractNativePPTX(nativeRectangularTextBoxFixture(t, false, properties, false), nativeTestExtractOptions()); err == nil {
			t.Fatalf("duplicate accepted: %s", properties)
		}
	}
}
