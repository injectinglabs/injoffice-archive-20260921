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
			parts[part] = strings.Replace(parts[part], `<a:buNone/>`, `<a:buFont typeface="Wingdings" charset="invalid"/><a:buChar char="q"/>`, 1)
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
	var layout, content bool
	for _, diagnostic := range element.Compatibility.Diagnostics {
		layout = layout || (diagnostic.Code == "pptx.text-layout-unavailable" && strings.Contains(diagnostic.Message, "a:spAutoFit"))
		content = content || (diagnostic.Code == "pptx.text-content-unavailable" && strings.Contains(diagnostic.Message, "buFont"))
	}
	if !layout || !content {
		t.Fatalf("diagnostics lost the independent autofit and symbol-font blockers: %+v", element.Compatibility.Diagnostics)
	}
}

func TestNativeSourceFrameUnsupportedContentRetainsApproximationAndRefusal(t *testing.T) {
	for _, strict := range []bool{false, true} {
		original := nativeRectangularTextBoxFixture(t, strict, `<a:noFill/>`, true)
		before := bytes.Clone(original)
		options := nativeTestExtractOptions()
		options.AllowSourceFrameAutoFitPreview = true
		deck, err := ExtractNativePPTX(original, options)
		if err != nil {
			t.Fatal(err)
		}
		if len(deck.Slides[0].Elements) != 1 {
			t.Fatal("unsupported text lost its placeholder")
		}
		element := deck.Slides[0].Elements[0]
		if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.Source == nil || len(element.Passthrough) == 0 {
			t.Fatalf("source-frame option weakened the content refusal: %+v", element)
		}
		if element.TextBody == nil || element.TextBody.AutoFit != "shape-source-frame" {
			t.Fatal("source-frame layout was lost")
		}
		if element.Paragraphs != nil && len(*element.Paragraphs) != 0 {
			t.Fatal("unsupported symbol-font text was invented")
		}
		warnings, refusals := 0, 0
		for _, diagnostic := range element.Compatibility.Diagnostics {
			if diagnostic.Code == "pptx.autofit-source-frame-approximate" && diagnostic.Severity == NativeDiagnosticSeverityWarning {
				warnings++
			}
			if diagnostic.Code == "pptx.text-content-unavailable" && diagnostic.Severity == NativeDiagnosticSeverityRefusal && strings.Contains(diagnostic.Message, "buFont") {
				refusals++
			}
		}
		if warnings != 1 || refusals != 1 {
			t.Fatalf("expected independent approximation and refusal: %+v", element.Compatibility.Diagnostics)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("invalid partial-preview contract: %+v", issues)
		}
		if !bytes.Equal(original, before) {
			t.Fatal("preview changed original bytes")
		}
		strictDeck, err := ExtractNativePPTX(original, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		for _, diagnostic := range strictDeck.Slides[0].Elements[0].Compatibility.Diagnostics {
			if diagnostic.Code == "pptx.autofit-source-frame-approximate" {
				t.Fatal("strict extraction opted into approximation")
			}
		}
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

// nativePowerPointAuthoredTextBoxFixture reproduces the two things PowerPoint
// writes on an authored WordArt text box that the exact subset does not model:
// an a16:creationId extension on p:cNvPr and an a:prstTxWarp on a:bodyPr.
func nativePowerPointAuthoredTextBoxFixture(t *testing.T, strict, creationID, warp bool) []byte {
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		if creationID {
			parts[part] = strings.Replace(parts[part], `<p:cNvPr id="2" name="Title"/>`,
				`<p:cNvPr id="2" name="Title"><a:extLst><a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{6DEFC7CA-01A0-475F-BBF2-53EE54DA4331}"/></a:ext></a:extLst></p:cNvPr>`, 1)
		}
		if warp {
			parts[part] = strings.Replace(parts[part], `<a:bodyPr/>`,
				`<a:bodyPr><a:prstTxWarp prst="textDeflate"><a:avLst><a:gd name="adj" fmla="val 37500"/></a:avLst></a:prstTxWarp></a:bodyPr>`, 1)
		}
	}})
}

func TestNativePowerPointAuthoredTextBoxPaintsOnlyInTheApproximateTier(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, shape := range []struct {
			name             string
			creationID, warp bool
			codes            []string
		}{
			{name: "creationId", creationID: true, codes: []string{nativeTextNonVisualPreviewCode}},
			{name: "prstTxWarp", warp: true, codes: []string{nativeTextWarpApproximateCode}},
			{name: "both", creationID: true, warp: true, codes: []string{nativeTextNonVisualPreviewCode, nativeTextWarpApproximateCode}},
		} {
			original := nativePowerPointAuthoredTextBoxFixture(t, strict, shape.creationID, shape.warp)
			before := bytes.Clone(original)
			// Exact tier: the creationId keeps the whole shape out of the
			// model, the warp keeps only its text out. Neither is painted.
			exact, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatalf("%s/%v exact extract: %v", shape.name, strict, err)
			}
			if shape.creationID && len(exact.Slides[0].Elements) != 0 {
				t.Fatalf("%s/%v: exact tier modeled a shape with unmodeled nonvisual metadata", shape.name, strict)
			}
			if !shape.creationID && exact.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("%s/%v: exact tier accepted a:prstTxWarp: %#v", shape.name, strict, exact.Slides[0].Elements[0])
			}

			options := nativeTestExtractOptions()
			options.AllowSourceFrameAutoFitPreview = true
			options.AllowInheritedTextPreview = true
			deck, err := ExtractNativePPTX(original, options)
			if err != nil {
				t.Fatalf("%s/%v approximate extract: %v", shape.name, strict, err)
			}
			if len(deck.Slides[0].Elements) != 1 {
				t.Fatalf("%s/%v: approximate tier dropped the text box", shape.name, strict)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("%s/%v: approximate preview is not preserve-only: %s", shape.name, strict, element.Compatibility.Status)
			}
			if element.TextBody == nil || element.Paragraphs == nil || len(*element.Paragraphs) != 1 || len((*element.Paragraphs)[0].Runs) != 2 {
				t.Fatalf("%s/%v: authored text was not painted: %#v", shape.name, strict, element)
			}
			if *element.Transform.X != 914_400 || *element.Transform.Y != 457_200 || *element.Transform.Cx != 4_572_000 || *element.Transform.Cy != 914_400 {
				t.Fatalf("%s/%v: the saved frame moved: %#v", shape.name, strict, element.Transform)
			}
			for _, code := range shape.codes {
				found := false
				for _, diagnostic := range element.Compatibility.Diagnostics {
					found = found || (diagnostic.Code == code && diagnostic.Severity == NativeDiagnosticSeverityWarning)
				}
				if !found {
					t.Fatalf("%s/%v: %s was not disclosed: %#v", shape.name, strict, code, element.Compatibility.Diagnostics)
				}
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("%s/%v invalid approximate contract: %#v", shape.name, strict, issues)
			}
			if !bytes.Equal(original, before) {
				t.Fatalf("%s/%v: preview changed source bytes", shape.name, strict)
			}
			// The approximation stays read-only: the element cannot be edited.
			paragraphs := nativeMutationParagraphs("Edited")
			if _, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}}); err == nil {
				t.Fatalf("%s/%v: the read-only approximation accepted a mutation", shape.name, strict)
			}
		}
	}
}
