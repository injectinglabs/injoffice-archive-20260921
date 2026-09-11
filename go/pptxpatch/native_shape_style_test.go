package pptxpatch

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

const nativeShapeStyleRefs = `<p:style><a:lnRef idx="1"><a:srgbClr val="123456"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="ABCDEF"/></a:fillRef><a:effectRef idx="0"><a:srgbClr val="000000"/></a:effectRef><a:fontRef idx="minor"><a:srgbClr val="FFFFFF"/></a:fontRef></p:style>`

func nativeShapeStyleFixture(t *testing.T, strict bool, shape, fill string) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, shape+`</p:spTree>`, 1)
		drawing := nsDrawingTransitional
		if strict {
			drawing = nsDrawingStrict
		}
		parts["relocated/themes/theme.xml"] = fmt.Sprintf(`<a:theme xmlns:a="%s" name="Synthetic"><a:themeElements><a:fmtScheme name="Synthetic"><a:fillStyleLst>%s</a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst></a:fmtScheme></a:themeElements></a:theme>`, drawing, fill)
	}})
}

func TestNativeShapeThemeStylesProjectWithoutGrantingMutation(t *testing.T) {
	for _, strict := range []bool{false, true} {
		shape := nativeAutoShapeXML(3, "Themed pentagon", "pentagon", "", "", "")
		shape = strings.Replace(shape, `</p:sp>`, nativeShapeStyleRefs+`</p:sp>`, 1)
		original := nativeShapeStyleFixture(t, strict, shape, `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`)
		before := bytes.Clone(original)
		deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		e := nativeFixtureAutoShapes(deck.Slides[0])[0]
		if e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || e.Fill == nil || *e.Fill != "ABCDEF" || e.Stroke == nil || e.Stroke.Color != "123456" || e.Preset == nil || *e.Preset != NativeShapePresetPentagon {
			t.Fatalf("theme paint lost: %+v", e)
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatal(issues)
		}
		paragraphs := nativeMutationParagraphs("Replacement")
		_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
		if err == nil || !strings.Contains(err.Error(), "preview-only") {
			t.Fatalf("mutation was not refused: %v", err)
		}
		update := NativePPTXAutoShapeMutation{Transform: e.Transform, Preset: NativeShapePresetRect, Fill: e.Fill, Stroke: e.Stroke}
		_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "shape", Kind: NativePPTXUpdateAutoShape, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, AutoShape: &update}}})
		if err == nil || !strings.Contains(err.Error(), "preview-only") {
			t.Fatalf("shape update was not refused: %v", err)
		}
		if !bytes.Equal(original, before) {
			t.Fatal("projection changed source bytes")
		}
	}
}

func TestNativeShapeStyleRefusesUnmodeledMatrixAndKeepsOverrides(t *testing.T) {
	for _, tc := range []struct {
		name, fill, refs string
		refused          bool
	}{
		{"solid", `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`, nativeShapeStyleRefs, false},
		{"gradient", `<a:gradFill/>`, nativeShapeStyleRefs, true},
		{"placeholder transform", `<a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill>`, nativeShapeStyleRefs, true},
		{"effect", `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`, strings.Replace(nativeShapeStyleRefs, `effectRef idx="0"`, `effectRef idx="1"`, 1), true},
		{"index", `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`, strings.Replace(nativeShapeStyleRefs, `fillRef idx="1"`, `fillRef idx="4"`, 1), true},
		{"style direct text", `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`, strings.Replace(nativeShapeStyleRefs, `<p:style>`, `<p:style>unmodeled`, 1), true},
		{"reference direct text", `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`, strings.Replace(nativeShapeStyleRefs, `<a:fillRef idx="1">`, `<a:fillRef idx="1">unmodeled`, 1), true},
		{"selected leaf direct text", `<a:solidFill><a:schemeClr val="phClr">unmodeled</a:schemeClr></a:solidFill>`, nativeShapeStyleRefs, true},
		{"duplicate reference", `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`, strings.Replace(nativeShapeStyleRefs, `</p:style>`, `<a:fillRef idx="1"><a:srgbClr val="112233"/></a:fillRef></p:style>`, 1), true},
		{"duplicate selected color", `<a:solidFill><a:schemeClr val="phClr"/><a:schemeClr val="phClr"/></a:solidFill>`, nativeShapeStyleRefs, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			shape := nativeAutoShapeXML(3, "Explicit override", "rect", `<a:solidFill><a:srgbClr val="FEDCBA"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
			shape = strings.Replace(shape, `</p:sp>`, tc.refs+`</p:sp>`, 1)
			deck, err := ExtractNativePPTX(nativeShapeStyleFixture(t, false, shape, tc.fill), nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			e := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if (e.Compatibility.Status == NativeCompatibilityStatusRefused) != tc.refused {
				t.Fatalf("unexpected status: %+v", e.Compatibility)
			}
			if !tc.refused && (e.Fill == nil || *e.Fill != "FEDCBA" || e.Stroke != nil) {
				t.Fatalf("explicit override lost: %+v", e)
			}
		})
	}
}

func TestNativeShapePartialTextRetainsGeometryAndRefusesMutation(t *testing.T) {
	shape := nativeAutoShapeXML(3, "Geometry with vertical text", "rect", `<a:solidFill><a:srgbClr val="123456"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	shape = strings.Replace(shape, `</p:sp>`, `<p:txBody><a:bodyPr vert="vert"/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"><a:latin typeface="Arial"/></a:rPr><a:t>Not a horizontal substitute</a:t></a:r></a:p></p:txBody></p:sp>`, 1)
	original := nativeAutoShapeFixture(t, false, shape)
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	e := nativeFixtureAutoShapes(deck.Slides[0])[0]
	if e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || e.Fill == nil || e.Paragraphs == nil || len(*e.Paragraphs) != 0 || e.TextBody != nil {
		t.Fatalf("partial projection not honest: %+v", e)
	}
	paragraphs := nativeMutationParagraphs("Replacement")
	_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
	if err == nil || !strings.Contains(err.Error(), "preview-only") {
		t.Fatalf("omitted text became editable: %v", err)
	}
	update := NativePPTXAutoShapeMutation{Transform: e.Transform, Preset: NativeShapePresetRect, Fill: e.Fill}
	_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "shape", Kind: NativePPTXUpdateAutoShape, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, AutoShape: &update}}})
	if err == nil || !strings.Contains(err.Error(), "preview-only") {
		t.Fatalf("partial shape became editable: %v", err)
	}
	unsupported := strings.Replace(shape, `<a:solidFill><a:srgbClr val="123456"/></a:solidFill>`, `<a:gradFill/>`, 1)
	bad, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, unsupported), nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	if nativeFixtureAutoShapes(bad.Slides[0])[0].Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatal("text omission must not override unsupported paint refusal")
	}
}
