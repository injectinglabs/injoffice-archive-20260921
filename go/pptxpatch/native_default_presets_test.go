package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeDefaultPresetPreview(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, preset := range []string{"roundRect", "rightArrow", "hexagon"} {
			t.Run(preset, func(t *testing.T) {
				shape := nativeAutoShapeXML(3, "Default preset", preset, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
				original := nativeShapeStyleFixture(t, strict, shape, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`)
				before := bytes.Clone(original)
				deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
				if err != nil {
					t.Fatal(err)
				}
				e := nativeFixtureAutoShapes(deck.Slides[0])[0]
				if e.Preset == nil || string(*e.Preset) != preset || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
					t.Fatalf("preset not qualified read-only: %+v", e)
				}
				if issues := ValidateNativePPTX(deck); len(issues) > 0 {
					t.Fatal(issues)
				}
				update := NativePPTXAutoShapeMutation{Transform: e.Transform, Preset: NativeShapePresetRect, Fill: e.Fill, Stroke: e.Stroke}
				_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "shape", Kind: NativePPTXUpdateAutoShape, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, AutoShape: &update}}})
				if err == nil || !strings.Contains(err.Error(), "preview-only") {
					t.Fatalf("mutation not refused: %v", err)
				}
				if !bytes.Equal(original, before) {
					t.Fatal("source bytes changed")
				}
				for _, av := range []string{`<a:avLst><a:gd name="unknown" fmla="val 16667"/></a:avLst>`, `<a:avLst extra="1"/>`, `<a:avLst>text</a:avLst>`, `<a:avLst><a:gd name="adj" fmla="*/ w 1 2"/></a:avLst>`, `<a:avLst><x:gd xmlns:x="urn:foreign"/></a:avLst>`} {
					candidate := strings.Replace(shape, `<a:avLst/>`, av, 1)
					negative, err := ExtractNativePPTX(nativeShapeStyleFixture(t, strict, candidate, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`), nativeTestExtractOptions())
					if err != nil {
						t.Fatal(err)
					}
					got := nativeFixtureAutoShapes(negative.Slides[0])[0]
					if got.Preset != nil || got.Compatibility.Status != NativeCompatibilityStatusRefused {
						t.Fatalf("unsupported geometry qualified: %+v", got)
					}
				}
			})
		}
	}
}
