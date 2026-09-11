package pptxpatch

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestNativeQuarterTurnShapesPreviewAndMutationBoundary(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, item := range []struct {
			angle   int64
			quarter int64
		}{{5400000, 1}, {10800000, 2}, {16200000, 3}, {-5400000, 3}, {27000000, 1}} {
			raw := nativeAutoShapeXML(3, "Rotated triangle", "triangle", `<a:solidFill><a:srgbClr val="123456"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), fmt.Sprintf(` rot="%d"`, item.angle))
			original := nativeAutoShapeFixture(t, strict, raw)
			before := bytes.Clone(original)
			deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			element := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || element.Transform.QuarterTurns == nil || *element.Transform.QuarterTurns != item.quarter || element.Preset == nil {
				t.Fatalf("rotation lost: %+v", element)
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatalf("invalid: %+v", issues)
			}
			paragraphs := nativeMutationParagraphs("Not allowed")
			_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
			if err == nil || !strings.Contains(err.Error(), "quarter-turn transforms are preview-only") {
				t.Fatalf("rotation granted edit permission: %v", err)
			}
			if !bytes.Equal(original, before) {
				t.Fatal("source bytes changed")
			}
		}
	}
}

// Optional local native/browser proof; no external corpus bytes or fonts.
func TestNativeQuarterTurnBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_ROTATION_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional local render fixture")
	}
	var shapes strings.Builder
	for q := 0; q < 4; q++ {
		shape := nativeAutoShapeXML(10+q, fmt.Sprintf("Quarter %d", q), "triangle", `<a:solidFill><a:srgbClr val="123456"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), fmt.Sprintf(` rot="%d"`, q*5400000))
		start := strings.Index(shape, `<a:off `)
		end := strings.Index(shape[start:], `</a:xfrm>`) + start
		shape = shape[:start] + fmt.Sprintf(`<a:off x="%d" y="2000000"/><a:ext cx="2000000" cy="1200000"/>`, 200000+q*2500000) + shape[end:]
		text := `<p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"/><a:lstStyle/><a:p><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>AB</a:t></a:r></a:p></p:txBody>`
		shapes.WriteString(strings.Replace(shape, `</p:sp>`, text+`</p:sp>`, 1))
	}
	input := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		part := "relocated/slides/slide-a.xml"
		raw := parts[part]
		start := strings.Index(raw, `<p:sp>`)
		end := strings.Index(raw[start:], `</p:sp>`) + start + len(`</p:sp>`)
		parts[part] = raw[:start] + shapes.String() + raw[end:]
	}})
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "rotation.pptx"), input, 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "rotation.json"), encoded, 0600); err != nil {
		t.Fatal(err)
	}
}

func TestNativeQuarterTurnShapesRefuseFractionalAndArbitraryRotation(t *testing.T) {
	for _, angle := range []string{"12345", "5400000"} {
		raw := nativeAutoShapeXML(3, "Unsupported rotation", "triangle", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), ` rot="`+angle+`"`)
		if angle == "5400000" {
			raw = strings.Replace(raw, `cx="`, `cx="1`, 1)
		}
		// Explicit odd/even extents require a fractional center at 90 degrees.
		if angle == "5400000" {
			start := strings.Index(raw, `<a:ext `)
			end := strings.Index(raw[start:], `/>`) + start + 2
			raw = raw[:start] + `<a:ext cx="1000001" cy="1000000"/>` + raw[end:]
		}
		deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, raw), nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if nativeFixtureAutoShapes(deck.Slides[0])[0].Compatibility.Status != NativeCompatibilityStatusRefused {
			t.Fatal("unrepresentable rotation accepted")
		}
	}
}

func TestNativeQuarterTurnCannotEnterMutationPayload(t *testing.T) {
	shape := NativePPTXAutoShapeMutation{Preset: NativeShapePresetTriangle, Transform: NativeTransform{X: int64Pointer(0), Y: int64Pointer(0), Cx: int64Pointer(1000000), Cy: int64Pointer(500000), QuarterTurns: int64Pointer(2)}}
	if err := validateNativeAutoShapeMutation(shape); err == nil || !strings.Contains(err.Error(), "preview-only") {
		t.Fatalf("quarter-turn mutation accepted: %v", err)
	}
}

func TestNativeQuarterTurnSourceAngleRequiresDrawingMLInt32(t *testing.T) {
	for _, angle := range []string{"2160000000", "-2160000000"} {
		raw := nativeAutoShapeXML(3, "Oversized angle", "triangle", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), ` rot="`+angle+`"`)
		if _, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, raw), nativeTestExtractOptions()); err == nil {
			t.Fatal("out-of-range DrawingML angle accepted")
		}
	}
}
