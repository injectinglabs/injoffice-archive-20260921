package pptxpatch

import (
	"fmt"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestNativeSourceAffineShapeExtractionAndMutationRefusal(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, attrs := range []string{` rot="1800000"`, ` flipH="true"`, ` rot="-2700000" flipV="1"`} {
			raw := nativeAutoShapeXML(3, "Affine", "triangle", `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), attrs)
			deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, strict, raw), nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			shape := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if shape.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || shape.Transform.RotationAngle == nil || shape.Preset == nil || shape.Source == nil {
				t.Fatalf("missing source affine qualification: %#v", shape)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatal(issues)
			}
			if err := validateNativeAutoShapeMutation(NativePPTXAutoShapeMutation{Preset: NativeShapePresetTriangle, Transform: shape.Transform}); err == nil {
				t.Fatal("source affine became editable")
			}
		}
	}
}
func TestNativeSourceAffineGroupExtractionBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		presentation, drawing := nativeGroupTestNamespaces(strict)
		group := nativeNestedGroupXML(presentation, drawing, ` rot="1800000" flipH="1"`)
		group = strings.Replace(group, `<a:chExt cx="`, `<a:chExt cx="3`, 1)
		payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
		}})
		deck, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if len(deck.Slides[0].Elements) != 2 {
			t.Fatalf("group not projected: %#v", deck.Slides[0].Compatibility)
		}
		got := deck.Slides[0].Elements[1]
		if got.Kind != NativeElementKindGroup || got.Transform.RotationAngle == nil || *got.Transform.RotationAngle != 1800000 || got.Transform.FlipH == nil || !*got.Transform.FlipH || got.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(got.Children) != 2 {
			t.Fatalf("invalid transformed group: %#v", got)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatal(issues)
		}
	}
}

func TestNativeSourceAffineAncestorCannotGrantMutation(t *testing.T) {
	for _, fractional := range []bool{false, true} {
		attrs := ` rot="1800000"`
		if fractional {
			attrs = ""
		}
		group := nativeNestedGroupXML(nsPresentationTransitional, nsDrawingTransitional, attrs)
		if fractional {
			group = strings.Replace(group, `cx="6000000" cy="8000000"`, `cx="6000001" cy="8000000"`, 1)
		}
		input := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
		}})
		before := append([]byte(nil), input...)
		deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		child := deck.Slides[0].Elements[1].Children[0].Children[0]
		if child.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || nativeHasSourceAffine(child.Transform) {
			t.Fatal("inherited authority was not retained separately from local geometry")
		}
		paragraphs := nativeMutationParagraphs("Must not mutate")
		output, err := ApplyNativePPTXMutations(input, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: child.ID, ExpectedFingerprintSHA256: child.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
		if err == nil || !strings.Contains(err.Error(), "affine ancestor") || len(output) != 0 {
			t.Fatalf("ancestor mutation accepted: %v", err)
		}
		if string(input) != string(before) {
			t.Fatal("source bytes changed on refusal")
		}
	}
}
func TestNativeSourceAffineGraphicFrameGroupRefusesAtomically(t *testing.T) {
	for _, chart := range []bool{false, true} {
		for _, attrs := range []string{` rot="1800000"`, ` flipH="1"`, ``} {
			frame := nativeExactTableGraphicFrameXML(4, "Table", []int64{1000000}, []int64{500000}, [][]string{{nativeExactTableCellXML("Cell", "l", "FFFFFF")}}, "")
			if chart {
				frame = nativeChartGraphicFrameXML(false, 4, "Chart", "")
			}
			group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Frame Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm` + attrs + `><a:off x="0" y="0"/><a:ext cx="6000000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="6000000" cy="4000000"/></a:xfrm></p:grpSpPr>` + frame + nativeAutoShapeXML(5, "Unsupported late child", "rect", `<a:gradFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "") + `</p:grpSp>`
			if attrs == "" {
				group = strings.Replace(group, `<a:chExt cx="6000000"`, `<a:chExt cx="6000001"`, 1)
			}
			var input []byte
			if chart {
				input = nativeChartFixture(t, nativeChartFixtureOptions{frameXML: group})
			} else {
				input = nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
					parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
				}})
			}
			deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if len(deck.Slides[0].Elements) != 1 || len(deck.Assets) != 0 || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.group-child-refused-unavailable") {
				t.Fatalf("complex graphic frame leaked partial content: %#v", deck.Slides[0].Compatibility)
			}
		}
	}
}
func TestNativeSourceAffineTextBoxAndUprightBoundary(t *testing.T) {
	for _, upright := range []bool{false, true} {
		input := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			raw := parts["relocated/slides/slide-a.xml"]
			raw = strings.Replace(raw, `<a:xfrm>`, `<a:xfrm rot="1800000" flipH="1">`, 1)
			if upright {
				raw = strings.Replace(raw, `<a:bodyPr/>`, `<a:bodyPr upright="1"/>`, 1)
			}
			parts["relocated/slides/slide-a.xml"] = raw
		}})
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		text := deck.Slides[0].Elements[0]
		if text.Transform.RotationAngle == nil || *text.Transform.RotationAngle != 1800000 || text.Transform.FlipH == nil || !*text.Transform.FlipH {
			t.Fatal("text box source orientation lost")
		}
		if upright {
			if text.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || text.TextBody == nil || text.TextBody.Upright == nil || !*text.TextBody.Upright {
				t.Fatal("upright source metadata or preview-only authority lost")
			}
		} else if text.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || text.TextBody == nil {
			t.Fatal("text box affine qualification failed")
		}
	}
}

// Generated functional evidence for actual WASM/compiler/browser qualification.
// No external Office fixture or renderer is used.
func TestNativeSourceAffineBrowserFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_AFFINE_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional local browser fixture")
	}
	body := func(text string) string {
		return `<p:txBody><a:bodyPr lIns="10000" rIns="20000" tIns="10000" bIns="10000" wrap="none"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="0" i="0" sz="1800"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:latin typeface="DejaVu Sans"/></a:rPr><a:t>` + text + `</a:t></a:r></a:p></p:txBody>`
	}
	shape := func(id int, x, y, cx, cy int64, preset, attrs, label string) string {
		raw := nativeAutoShapeXML(id, label, preset, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), attrs)
		start := strings.Index(raw, `<a:off `)
		end := strings.Index(raw[start:], `</a:xfrm>`) + start
		raw = raw[:start] + fmt.Sprintf(`<a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/>`, x, y, cx, cy) + raw[end:]
		return strings.Replace(raw, `</p:sp>`, body(label)+`</p:sp>`, 1)
	}
	var content strings.Builder
	for i, attrs := range []string{``, ` rot="1800000" flipH="1"`, ` flipV="1"`, ` flipH="1" flipV="1"`} {
		content.WriteString(shape(10+i, 300000+int64(i)*2800000, 800000, 2000000, 1200000, "triangle", attrs, []string{"BASE", "H + 30", "V", "H + V"}[i]))
	}
	content.WriteString(shape(20, 500000, 3800000, 2200000, 1500000, "wedgeRoundRectCallout", ` rot="2100000"`, "CALLOUT"))
	group := func(id int, x, y, cx, cy, chx, chy int64, attrs, child string) string {
		return fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="%d" name="Affine Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/><a:chOff x="0" y="0"/><a:chExt cx="%d" cy="%d"/></a:xfrm></p:grpSpPr>%s</p:grpSp>`, id, attrs, x, y, cx, cy, chx, chy, child)
	}
	content.WriteString(group(30, 3500000, 3800000, 2500000, 1600000, 2000000, 2000000, ` rot="900000"`, shape(31, 200000, 200000, 1400000, 1400000, "triangle", ` rot="5400000"`, "AXES")))
	content.WriteString(group(40, 6800000, 3800000, 2200000, 1600000, 2000001, 1600003, ` flipH="1"`, shape(41, 0, 0, 2000000, 1400000, "hexagon", ` rot="600000"`, "RATIO")))
	content.WriteString(shape(50, 9500000, 4000000, 1800000, 1200000, "star5", ` rot="1500000"`, "STAR"))
	makeInput := func(shapes string) []byte {
		return nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			part := "relocated/slides/slide-a.xml"
			raw := parts[part]
			start := strings.Index(raw, `<p:sp>`)
			end := strings.Index(raw[start:], `</p:sp>`) + start + len(`</p:sp>`)
			parts[part] = raw[:start] + shapes + raw[end:]
		}})
	}
	input := makeInput(content.String())
	deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	for name, bytes := range map[string][]byte{"affine.pptx": input, "affine-go.json": encoded, "affine-upright-refused.pptx": makeInput(strings.Replace(shape(60, 1000000, 1000000, 3000000, 2000000, "triangle", ` rot="1800000"`, "UPRIGHT"), `<a:bodyPr `, `<a:bodyPr upright="1" `, 1))} {
		if err := os.WriteFile(filepath.Join(dir, name), bytes, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
