package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXProjectsDiagramDrawingFallback(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{strict: strict}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract diagram: %v", err)
			}
			slide := deck.Slides[0]
			group := nativeFixtureDiagramGroup(t, slide)
			if group.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(group.Children) != 5 || group.ChildTransform == nil {
				t.Fatalf("diagram drawing was not projected as one preserve-only group: %#v", group)
			}
			if *group.Transform.X != 1_524_000 || *group.Transform.Y != 1_397_000 || *group.Transform.Cx != 6_096_000 || *group.Transform.Cy != 4_064_000 {
				t.Fatalf("frame EMU changed: %#v", group.Transform)
			}
			if *group.ChildTransform.X != 0 || *group.ChildTransform.Y != 0 || *group.ChildTransform.Cx != 6_096_000 || *group.ChildTransform.Cy != 4_064_000 {
				t.Fatalf("child coordinate space must be the identity frame space: %#v", group.ChildTransform)
			}
			if len(group.Passthrough) != 1 || group.Passthrough[0].OwnerPart != slide.Source.PartName || group.Source == nil || group.Source.ObjectID != "cNvPr-4" || group.Name == nil || *group.Name != "Diagram 3" {
				t.Fatalf("frame bytes were not preserved with the source anchor: %#v", group)
			}
			if len(group.Compatibility.Diagnostics) != 1 || group.Compatibility.Diagnostics[0].Code != nativeDiagramDrawingPreviewCode || !strings.Contains(group.Compatibility.Diagnostics[0].Message, nativeDiagramDrawingPolicy) {
				t.Fatalf("diagram preview policy was not declared: %#v", group.Compatibility.Diagnostics)
			}
			seen := map[string]bool{}
			for index, child := range group.Children {
				if child.Kind != NativeElementKindShape || child.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || child.Geometry == nil || child.Preset != nil {
					t.Fatalf("child %d is not a read-only evaluated shape: %#v", index, child)
				}
				// PowerPoint omits the join, so cap/join/dash stay unset together
				// (contract all-or-none) instead of inventing a join.
				if child.Fill == nil || *child.Fill != "2F6FED" || child.Stroke == nil || child.Stroke.Color != "FFFFFF" || child.Stroke.WidthEMU == nil || *child.Stroke.WidthEMU != 12_700 || child.Stroke.Dash != nil || child.Stroke.Cap != nil || child.Stroke.Join != nil {
					t.Fatalf("child %d paint was not taken from the drawing part: fill=%v stroke=%#v", index, child.Fill, child.Stroke)
				}
				if child.Source == nil || child.Source.PartName != slide.Source.PartName || !strings.HasPrefix(child.Source.ObjectID, "cNvPr-4/dsp/") || seen[child.ID] {
					t.Fatalf("child %d lacks a unique slide-owned source anchor: %#v", index, child.Source)
				}
				seen[child.ID] = true
				if len(child.Compatibility.Diagnostics) != 1 || child.Compatibility.Diagnostics[0].Code != nativeDiagramDrawingPreviewCode {
					t.Fatalf("child %d diagnostics: %#v", index, child.Compatibility.Diagnostics)
				}
				if child.Paragraphs == nil || len(*child.Paragraphs) != 1 || len((*child.Paragraphs)[0].Runs) != 1 || child.TextBody == nil {
					t.Fatalf("child %d text was not projected: %#v", index, child.Paragraphs)
				}
				run := (*child.Paragraphs)[0].Runs[0]
				if run.Color == nil || *run.Color != "FFFFFF" || run.FontFamily == nil || *run.FontFamily != "Calibri" || run.FontSizeHundredthPt == nil || *run.FontSizeHundredthPt != 2000 {
					t.Fatalf("child %d run did not resolve the dsp:style font reference: %#v", index, run)
				}
			}
			first := group.Children[0]
			if *first.Transform.X != 2_133_600 || *first.Transform.Y != 0 || *first.Transform.Cx != 1_828_800 || *first.Transform.Cy != 914_400 || *(*first.Paragraphs)[0].Runs[0].Text != "Manager" {
				t.Fatalf("first shape position or text drifted from dsp:sp: %#v", first)
			}
			if *(*group.Children[4].Paragraphs)[0].Runs[0].Text != "Employee2" {
				t.Fatalf("shape order was not preserved: %#v", group.Children[4].Paragraphs)
			}
			for _, diagnostic := range slide.Compatibility.Diagnostics {
				if strings.HasPrefix(diagnostic.Code, "pptx.diagram-") && diagnostic.Code != nativeDiagramDrawingPreviewCode {
					t.Fatalf("projected diagram must not also be refused at slide level: %#v", diagnostic)
				}
			}
			if slide.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("slide status: %#v", slide.Compatibility.Status)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid diagram deck: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("encode: %v", err)
			}
			if _, err := DecodeNativePPTXJSON(encoded); err != nil {
				t.Fatalf("decode: %v", err)
			}
		})
	}
}

func TestExtractNativePPTXDiagramDrawingChildrenRemainPreviewOnlyForMutation(t *testing.T) {
	t.Parallel()
	original := nativeDiagramFixture(t, nativeDiagramFixtureOptions{})
	deck, err := ExtractNativePPTX(original, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract diagram: %v", err)
	}
	group := nativeFixtureDiagramGroup(t, deck.Slides[0])
	child := group.Children[0]
	text := "Renamed"
	align := NativeTextAlignCenter
	paragraphs := []NativeParagraph{{Runs: []NativeTextRun{{Text: &text}}, Align: &align, Level: int64Pointer(0), Bullet: boolPointer(false)}}
	for _, target := range []NativeElement{child, group} {
		request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{
			OperationID: "op-1", Kind: NativePPTXReplaceText, ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs,
		}}}
		if produced, err := ApplyNativePPTXMutations(original, request); err == nil || produced != nil {
			t.Fatalf("diagram drawing element %q accepted a mutation", target.ID)
		}
	}
}

func TestExtractNativePPTXDiagramEmptyDrawingRefusesFrameNotSlide(t *testing.T) {
	t.Parallel()
	// Mirrors LibreOffice's smartart-org-chart.pptx: PowerPoint wrote an empty
	// dsp:spTree, so nothing source-backed exists to paint.
	empty := `<dsp:drawing xmlns:dgm="` + nativeDiagramURITransitional + `" xmlns:dsp="` + nsDiagramDrawing + `" xmlns:a="` + nsDrawingTransitional + `"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/></dsp:spTree></dsp:drawing>`
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: empty}, "pptx.diagram-drawing-empty-unavailable")
}

func TestExtractNativePPTXDiagramWithoutDrawingPartIsRefused(t *testing.T) {
	t.Parallel()
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{omitDrawingRelationship: true, omitDataModelExt: true}, "pptx.diagram-drawing-unavailable")
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{omitDrawingRelationship: true}, "pptx.diagram-drawing-unavailable")
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{omitDataModelExt: true, secondDrawingRelationship: true}, "pptx.diagram-drawing-unavailable")
}

func TestExtractNativePPTXDiagramDrawingWithoutDataModelExtUsesTheOnlyDrawing(t *testing.T) {
	t.Parallel()
	deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{omitDataModelExt: true}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract diagram: %v", err)
	}
	if group := nativeFixtureDiagramGroup(t, deck.Slides[0]); len(group.Children) != 5 {
		t.Fatalf("single unambiguous drawing relationship was not used: %#v", group)
	}
}

func TestExtractNativePPTXDiagramDrawingMalformedPartIsRefused(t *testing.T) {
	t.Parallel()
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: `<dsp:drawing xmlns:dsp="` + nsDiagramDrawing + `"><dsp:spTree>`}, "pptx.diagram-drawing-markup-unavailable")
	nested := nativeDiagramDrawingXML(nsDrawingTransitional, `<dsp:grpSp/>`+nativeDiagramShapeXML(0, "Manager", ""))
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nested}, "pptx.diagram-drawing-markup-unavailable")
}

func TestExtractNativePPTXDiagramDrawingUnmodeledPaintRefusesWholeFrame(t *testing.T) {
	t.Parallel()
	gradient := strings.Replace(nativeDiagramShapeXML(1, "Employee", ""),
		`<a:solidFill><a:schemeClr val="accent1"><a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/></a:schemeClr></a:solidFill>`,
		`<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill>`, 1)
	if !strings.Contains(gradient, "gradFill") {
		t.Fatal("fixture did not swap the fill")
	}
	drawing := nativeDiagramDrawingXML(nsDrawingTransitional, nativeDiagramShapeXML(0, "Manager", "")+gradient)
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: drawing}, "pptx.diagram-drawing-fill-unavailable")

	hueShift := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), `<a:hueOff val="0"/>`, `<a:hueOff val="1200000"/>`, 1)
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, hueShift)}, "pptx.diagram-drawing-fill-unavailable")

	dashed := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), `<a:prstDash val="solid"/>`, `<a:prstDash val="dash"/>`, 1)
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, dashed)}, "pptx.diagram-drawing-line-unavailable")

	bevel := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), `<a:effectLst/>`, `<a:effectLst/><a:sp3d/>`, 1)
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, bevel)}, "pptx.diagram-drawing-effects-unavailable")

	hidden := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), `<dsp:cNvPr id="0" name=""/>`, `<dsp:cNvPr id="0" name="" hidden="1"/>`, 1)
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, hidden)}, "pptx.diagram-drawing-markup-unavailable")
}

func TestExtractNativePPTXDiagramDrawingOmitsNonExactTextButKeepsGeometry(t *testing.T) {
	t.Parallel()
	spaced := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), `<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>`, `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>`, 1)
	deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, spaced)}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract diagram: %v", err)
	}
	group := nativeFixtureDiagramGroup(t, deck.Slides[0])
	child := group.Children[0]
	if child.Paragraphs == nil || len(*child.Paragraphs) != 0 || child.TextBody != nil || child.Fill == nil || child.Geometry == nil {
		t.Fatalf("non-exact text must be omitted without dropping the shape: %#v", child)
	}
	codes := []string{}
	for _, diagnostic := range child.Compatibility.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	if strings.Join(codes, ",") != nativeDiagramDrawingPreviewCode+","+nativeDiagramDrawingTextOmittedCode {
		t.Fatalf("text omission was not declared: %v", codes)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid diagram deck: %#v", issues)
	}
}

func TestExtractNativePPTXDiagramDrawingInheritedTextUsesDeclaredPreviewOnly(t *testing.T) {
	t.Parallel()
	// PowerPoint omits b/i on drawing runs. The exact tier omits that text;
	// only the opt-in inherited preview supplies the declared defaults.
	implicit := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), ` b="0" i="0"`, ``, 1)
	drawing := nativeDiagramDrawingXML(nsDrawingTransitional, implicit)
	exact, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: drawing}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract exact diagram: %v", err)
	}
	exactChild := nativeFixtureDiagramGroup(t, exact.Slides[0]).Children[0]
	if len(*exactChild.Paragraphs) != 0 || len(exactChild.Compatibility.Diagnostics) != 2 || exactChild.Compatibility.Diagnostics[1].Code != nativeDiagramDrawingTextOmittedCode {
		t.Fatalf("exact tier must omit non-self-contained diagram text: %#v", exactChild)
	}
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	approximate, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: drawing}), options)
	if err != nil {
		t.Fatalf("extract approximate diagram: %v", err)
	}
	child := nativeFixtureDiagramGroup(t, approximate.Slides[0]).Children[0]
	if len(*child.Paragraphs) != 1 || *(*child.Paragraphs)[0].Runs[0].Text != "Manager" || (*child.Paragraphs)[0].Runs[0].Color == nil || *(*child.Paragraphs)[0].Runs[0].Color != "FFFFFF" {
		t.Fatalf("inherited preview did not resolve diagram text: %#v", child.Paragraphs)
	}
	codes := []string{}
	for _, diagnostic := range child.Compatibility.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	if strings.Join(codes, ",") != nativeDiagramDrawingPreviewCode+","+nativeInheritedTextPreviewCode {
		t.Fatalf("inherited diagram text was not labeled with the declared policy: %v", codes)
	}
	if issues := ValidateNativePPTX(approximate); len(issues) != 0 {
		t.Fatalf("invalid approximate diagram deck: %#v", issues)
	}
}

func TestExtractNativePPTXDiagramDrawingStyleMatrixIsExactOrRefused(t *testing.T) {
	t.Parallel()
	base := nativeDiagramShapeXML(0, "Manager", "")
	lineXML := `<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="lt1"><a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln>`
	if !strings.Contains(base, lineXML) || !strings.Contains(base, `<a:lnRef idx="2">`) {
		t.Fatal("fixture drifted")
	}
	// lnRef idx="2" has no entry in the one-line fixture theme: the style cannot
	// be resolved, and without an explicit a:ln the frame must be refused rather
	// than painting an outline-less shape.
	inherited := strings.Replace(base, lineXML, "", 1)
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, inherited)}, "pptx.diagram-drawing-style-unavailable")

	// A resolvable exact matrix entry supplies the outline the shape omits.
	resolvable := strings.Replace(inherited, `<a:lnRef idx="2">`, `<a:lnRef idx="1">`, 1)
	exactTheme := strings.Replace(nativeExactThemeXML(nsDrawingTransitional),
		`<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>`,
		`<a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst>`, 1)
	if !strings.Contains(exactTheme, `<a:miter lim="800000"/>`) {
		t.Fatal("theme fixture drifted")
	}
	// The fixture theme's bare matrix line is itself outside the exact subset,
	// so it must refuse the frame too.
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, resolvable)}, "pptx.diagram-drawing-style-unavailable")
	deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, resolvable), themeXML: exactTheme}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract style-resolved diagram: %v", err)
	}
	child := nativeFixtureDiagramGroup(t, deck.Slides[0]).Children[0]
	if child.Stroke == nil || child.Stroke.Color != "FFFFFF" || child.Stroke.WidthEMU == nil || *child.Stroke.WidthEMU != 6350 || child.Stroke.Join == nil || *child.Stroke.Join != NativeStrokeJoinMiter || child.Fill == nil || *child.Fill != "2F6FED" {
		t.Fatalf("style matrix outline was not resolved from the theme: %#v", child.Stroke)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid diagram deck: %#v", issues)
	}
}

func TestExtractNativePPTXDiagramDrawingKerningThresholdFollowsAutoShapeTiers(t *testing.T) {
	t.Parallel()
	kerned := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), ` kern="0"`, ` kern="1200"`, 1)
	if !strings.Contains(kerned, `kern="1200"`) {
		t.Fatal("fixture drifted")
	}
	drawing := nativeDiagramDrawingXML(nsDrawingTransitional, kerned)
	exact, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: drawing}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract exact diagram: %v", err)
	}
	exactChild := nativeFixtureDiagramGroup(t, exact.Slides[0]).Children[0]
	if len(*exactChild.Paragraphs) != 0 || len(exactChild.Compatibility.Diagnostics) != 2 || exactChild.Compatibility.Diagnostics[1].Code != nativeDiagramDrawingTextOmittedCode {
		t.Fatalf("exact tier must omit kerned diagram text like AutoShape runs: %#v", exactChild.Compatibility.Diagnostics)
	}
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	approximate, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: drawing}), options)
	if err != nil {
		t.Fatalf("extract approximate diagram: %v", err)
	}
	child := nativeFixtureDiagramGroup(t, approximate.Slides[0]).Children[0]
	if len(*child.Paragraphs) != 1 || *(*child.Paragraphs)[0].Runs[0].Text != "Manager" {
		t.Fatalf("inherited preview did not paint kerned diagram text: %#v", child.Paragraphs)
	}
	codes := []string{}
	for _, diagnostic := range child.Compatibility.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	if strings.Join(codes, ",") != nativeDiagramDrawingPreviewCode+","+nativeInheritedTextPreviewCode {
		t.Fatalf("kerned diagram text was not labeled with the inherited policy: %v", codes)
	}
}

func TestExtractNativePPTXDiagramDrawingRotationDeclaresAffinePreview(t *testing.T) {
	t.Parallel()
	for _, tc := range []struct {
		rotation string
		code     string
	}{{"5400000", "pptx.quarter-turn-preview"}, {"1200000", "pptx.source-affine-preview"}} {
		rotated := strings.Replace(nativeDiagramShapeXML(0, "Manager", ""), `<a:xfrm>`, `<a:xfrm rot="`+tc.rotation+`">`, 1)
		deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, rotated)}), nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("extract rotated diagram (%s): %v", tc.rotation, err)
		}
		child := nativeFixtureDiagramGroup(t, deck.Slides[0]).Children[0]
		codes := []string{}
		for _, diagnostic := range child.Compatibility.Diagnostics {
			codes = append(codes, diagnostic.Code)
		}
		if !strings.Contains(strings.Join(codes, ","), tc.code) {
			t.Fatalf("rotation %s did not declare %s: %v", tc.rotation, tc.code, codes)
		}
		if tc.rotation == "5400000" && (child.Transform.QuarterTurns == nil || *child.Transform.QuarterTurns != 1 || child.Transform.RotationAngle != nil) {
			t.Fatalf("quarter turn was not retained exactly: %#v", child.Transform)
		}
		if tc.rotation == "1200000" && (child.Transform.RotationAngle == nil || *child.Transform.RotationAngle != 1_200_000 || child.Transform.QuarterTurns != nil) {
			t.Fatalf("source rotation was not retained exactly: %#v", child.Transform)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("invalid rotated diagram deck (%s): %#v", tc.rotation, issues)
		}
	}
}

func TestExtractNativePPTXDiagramDrawingTextFrameOverrideOmitsText(t *testing.T) {
	t.Parallel()
	base := nativeDiagramShapeXML(0, "Manager", "")
	start := strings.Index(base, "<dsp:txXfrm>")
	if start < 0 {
		t.Fatal("fixture lacks dsp:txXfrm")
	}
	moved := base[:start] + `<dsp:txXfrm><a:off x="2133600" y="0"/><a:ext cx="1828800" cy="914400"/></dsp:txXfrm></dsp:sp>`
	deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, moved)}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract diagram with text frame override: %v", err)
	}
	child := nativeFixtureDiagramGroup(t, deck.Slides[0]).Children[0]
	if len(*child.Paragraphs) != 0 || child.TextBody != nil || len(child.Compatibility.Diagnostics) != 2 || child.Compatibility.Diagnostics[1].Code != nativeDiagramDrawingTextOmittedCode || !strings.Contains(child.Compatibility.Diagnostics[1].Message, "dsp:txXfrm") {
		t.Fatalf("text frame override must omit text with a declared reason: %#v", child.Compatibility.Diagnostics)
	}
	// Without dsp:txXfrm the preset text rectangle is authoritative.
	plain := base[:start] + `</dsp:sp>`
	deck, err = ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, plain)}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract diagram without text frame: %v", err)
	}
	if child := nativeFixtureDiagramGroup(t, deck.Slides[0]).Children[0]; len(*child.Paragraphs) != 1 {
		t.Fatalf("text without dsp:txXfrm was not projected: %#v", child.Compatibility.Diagnostics)
	}
}

func TestExtractNativePPTXDiagramDrawingShapeBudgetIsBounded(t *testing.T) {
	t.Parallel()
	var shapes strings.Builder
	for index := 0; index <= nativeMaxDiagramDrawingShapes; index++ {
		shapes.WriteString(`<dsp:sp><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></dsp:spPr></dsp:sp>`)
	}
	assertNativeDiagramRefused(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, shapes.String())}, "pptx.diagram-drawing-budget-unavailable")
}

func TestExtractNativePPTXDiagramDrawingNoFillConnectorShape(t *testing.T) {
	t.Parallel()
	connector := `<dsp:sp modelId="{C1}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr><a:xfrm><a:off x="3048000" y="914400"/><a:ext cx="45720" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="1D2427"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></dsp:spPr></dsp:sp>`
	deck, err := ExtractNativePPTX(nativeDiagramFixture(t, nativeDiagramFixtureOptions{drawingXML: nativeDiagramDrawingXML(nsDrawingTransitional, connector+nativeDiagramShapeXML(0, "Manager", ""))}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract diagram: %v", err)
	}
	group := nativeFixtureDiagramGroup(t, deck.Slides[0])
	connectorStroke := group.Children[0].Stroke
	if len(group.Children) != 2 || group.Children[0].Fill != nil || connectorStroke == nil || connectorStroke.Color != "1D2427" || connectorStroke.Join == nil || *connectorStroke.Join != NativeStrokeJoinMiter || connectorStroke.MiterLimit == nil || *connectorStroke.MiterLimit != 800_000 || connectorStroke.Cap == nil || *connectorStroke.Cap != NativeStrokeCapFlat || connectorStroke.Dash == nil || *connectorStroke.Dash != NativeStrokeDashSolid {
		t.Fatalf("no-fill connector shape was not projected exactly: %#v", group.Children)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid diagram deck: %#v", issues)
	}
}

func assertNativeDiagramRefused(t *testing.T, options nativeDiagramFixtureOptions, code string) {
	t.Helper()
	deck, err := ExtractNativePPTX(nativeDiagramFixture(t, options), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract refused diagram (%s): %v", code, err)
	}
	slide := deck.Slides[0]
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindGroup {
			t.Fatalf("refused diagram leaked a partial projection (%s): %#v", code, element)
		}
	}
	found := false
	for _, diagnostic := range slide.Compatibility.Diagnostics {
		if diagnostic.Code == code {
			found = true
		}
	}
	if !found || slide.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(slide.Passthrough) == 0 {
		codes := []string{}
		for _, diagnostic := range slide.Compatibility.Diagnostics {
			codes = append(codes, diagnostic.Code)
		}
		t.Fatalf("diagram frame was not refused with %s and preserved: %v passthrough=%d", code, codes, len(slide.Passthrough))
	}
	// The text box next to the frame must keep painting: refusal is per element.
	if len(slide.Elements) != 1 || slide.Elements[0].Kind != NativeElementKindText {
		t.Fatalf("sibling text element was lost by the diagram refusal: %#v", slide.Elements)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid refused diagram deck: %#v", issues)
	}
}

func nativeFixtureDiagramGroup(t *testing.T, slide NativeSlide) NativeElement {
	t.Helper()
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindGroup {
			return element
		}
	}
	t.Fatalf("slide has no projected diagram group: %#v", slide)
	return NativeElement{}
}

type nativeDiagramFixtureOptions struct {
	strict                    bool
	drawingXML                string
	omitDrawingRelationship   bool
	omitDataModelExt          bool
	secondDrawingRelationship bool
	frameXML                  string
	themeXML                  string
	// Optional overrides of the four diagram parts (native_diagram_layout_test.go).
	dataXML   string
	layoutXML string
	styleXML  string
	colorsXML string
}

func nativeDiagramDrawingXML(drawingNS, shapes string) string {
	diagramNS := nativeDiagramURITransitional
	if drawingNS == nsDrawingStrict {
		diagramNS = nativeDiagramURIStrict
	}
	return `<dsp:drawing xmlns:dgm="` + diagramNS + `" xmlns:dsp="` + nsDiagramDrawing + `" xmlns:a="` + drawingNS + `"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>` + shapes + `</dsp:spTree></dsp:drawing>`
}

// nativeDiagramShapeXML mirrors what PowerPoint writes into drawing1.xml:
// explicit frame-relative xfrm, preset, scheme paint with zero offsets, a
// style matrix reference, and text with explicit size but no typeface.
func nativeDiagramShapeXML(index int, text, extra string) string {
	x := 2_133_600 + int64(index%3)*2_286_000 - int64(index/3)*2_286_000
	if x < 0 {
		x = 0
	}
	y := int64(index/3) * 1_371_600
	// PowerPoint writes the preset text rectangle into dsp:txXfrm.
	geometry, err := EvaluateNativePPTXPresetGeometry("roundRect", 1_828_800, 914_400, nil)
	if err != nil {
		panic(err)
	}
	textRect := geometry.TextRect
	return fmt.Sprintf(`<dsp:sp modelId="{%08X-0000-0000-0000-000000000000}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>`+
		`<dsp:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>`+
		`<a:solidFill><a:schemeClr val="accent1"><a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/></a:schemeClr></a:solidFill>`+
		`<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="lt1"><a:hueOff val="0"/><a:satOff val="0"/><a:lumOff val="0"/><a:alphaOff val="0"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:effectLst/>%s</dsp:spPr>`+
		`<dsp:style><a:lnRef idx="2"><a:schemeClr val="lt1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></dsp:style>`+
		`<dsp:txBody><a:bodyPr spcFirstLastPara="0" vert="horz" wrap="square" lIns="36195" tIns="36195" rIns="36195" bIns="36195" numCol="1" spcCol="1270" anchor="ctr" anchorCtr="0"><a:noAutofit/></a:bodyPr><a:lstStyle/>`+
		`<a:p><a:pPr lvl="0" algn="ctr" defTabSz="889000"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:spcAft><a:spcPct val="0"/></a:spcAft><a:buNone/></a:pPr><a:r><a:rPr lang="en-US" sz="2000" b="0" i="0" kern="0" dirty="0"/><a:t>%s</a:t></a:r></a:p></dsp:txBody>`+
		`<dsp:txXfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></dsp:txXfrm></dsp:sp>`, index+1, x, y, extra, text, x+textRect.X, y+textRect.Y, textRect.CX, textRect.CY)
}

func nativeDiagramFixture(t *testing.T, options nativeDiagramFixtureOptions) []byte {
	t.Helper()
	drawingNS := nsDrawingTransitional
	relsNS := nsOfficeRelsTransitional
	if options.strict {
		drawingNS = nsDrawingStrict
		relsNS = nsOfficeRelsStrict
	}
	diagramNS := nativeDiagramURITransitional
	if options.strict {
		diagramNS = nativeDiagramURIStrict
	}
	drawing := options.drawingXML
	if drawing == "" {
		var shapes strings.Builder
		for index, text := range []string{"Manager", "Manager2", "Assistant", "Employee", "Employee2"} {
			shapes.WriteString(nativeDiagramShapeXML(index, text, ""))
		}
		drawing = nativeDiagramDrawingXML(drawingNS, shapes.String())
	}
	modelExt := `<dgm:extLst><a:ext uri="` + nsDiagramDrawing + `"><dsp:dataModelExt xmlns:dsp="` + nsDiagramDrawing + `" relId="rIdDrawing" minVer="` + diagramNS + `"/></a:ext></dgm:extLst>`
	if options.omitDataModelExt {
		modelExt = ""
	}
	dataPart := "relocated/diagrams/data1.xml"
	layoutPart := "relocated/diagrams/layout1.xml"
	stylePart := "relocated/diagrams/quickStyle1.xml"
	colorsPart := "relocated/diagrams/colors1.xml"
	drawingPart := "relocated/diagrams/drawing1.xml"
	dataXML := `<dgm:dataModel xmlns:dgm="` + diagramNS + `" xmlns:a="` + drawingNS + `"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/>` + modelExt + `</dgm:dataModel>`
	if options.dataXML != "" {
		dataXML = options.dataXML
	}
	layoutXML := `<dgm:layoutDef xmlns:dgm="` + diagramNS + `" uniqueId="urn:test/layout"/>`
	if options.layoutXML != "" {
		layoutXML = options.layoutXML
	}
	styleXML := `<dgm:styleDef xmlns:dgm="` + diagramNS + `" uniqueId="urn:test/style"/>`
	if options.styleXML != "" {
		styleXML = options.styleXML
	}
	colorsXML := `<dgm:colorsDef xmlns:dgm="` + diagramNS + `" uniqueId="urn:test/colors"/>`
	if options.colorsXML != "" {
		colorsXML = options.colorsXML
	}
	extra := []nativeExtractZipPart{
		{name: dataPart, data: dataXML},
		{name: layoutPart, data: layoutXML},
		{name: stylePart, data: styleXML},
		{name: colorsPart, data: colorsXML},
		{name: drawingPart, data: drawing},
	}
	slideRels := fmt.Sprintf(`<Relationship Id="rIdDm" Type="%s/diagramData" Target="../diagrams/data1.xml"/><Relationship Id="rIdLo" Type="%s/diagramLayout" Target="../diagrams/layout1.xml"/><Relationship Id="rIdQs" Type="%s/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/><Relationship Id="rIdCs" Type="%s/diagramColors" Target="../diagrams/colors1.xml"/>`, relsNS, relsNS, relsNS, relsNS)
	if !options.omitDrawingRelationship {
		slideRels += `<Relationship Id="rIdDrawing" Type="` + relDiagramDrawing + `" Target="../diagrams/drawing1.xml"/>`
	}
	if options.secondDrawingRelationship {
		slideRels += `<Relationship Id="rIdDrawing2" Type="` + relDiagramDrawing + `" Target="../diagrams/drawing1.xml"/>`
	}
	frame := options.frameXML
	if frame == "" {
		frame = `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="4" name="Diagram 3"/><p:cNvGraphicFramePr/><p:nvPr><p:extLst><p:ext uri="{D42A27DB-BD31-4B8C-83A1-F6EECF244321}"><p14:modId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="4063117179"/></p:ext></p:extLst></p:nvPr></p:nvGraphicFramePr>` +
			`<p:xfrm><a:off x="1524000" y="1397000"/><a:ext cx="6096000" cy="4064000"/></p:xfrm><a:graphic><a:graphicData uri="` + diagramNS + `"><dgm:relIds xmlns:dgm="` + diagramNS + `" xmlns:r="` + relsNS + `" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData></a:graphic></p:graphicFrame>`
	}
	return nativeExtractFixture(t, nativeExtractFixtureOptions{
		strict: options.strict, extraParts: extra,
		mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
				fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/><Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/><Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/><Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/><Override PartName="/%s" ContentType="%s"/></Types>`,
					dataPart, contentTypeDiagramData, layoutPart, stylePart, colorsPart, drawingPart, contentTypeDiagramDrawing), 1)
			parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`, slideRels+`</Relationships>`, 1)
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, frame+`</p:spTree>`, 1)
			parts["relocated/themes/theme.xml"] = nativeExactThemeXML(drawingNS)
			if options.themeXML != "" {
				parts["relocated/themes/theme.xml"] = options.themeXML
			}
		},
	})
}
