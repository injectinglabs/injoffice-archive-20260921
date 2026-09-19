package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

// A refused diagram must say which layout shape refused and which construct
// did it: "diagram shape ... is not modeled" alone identified neither.
func TestNativeDiagramLayoutRefusalNamesTheShapeAndTheConstruct(t *testing.T) {
	layout := strings.Replace(nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional),
		`<dgm:shape type="rect"><dgm:adjLst/></dgm:shape>`,
		`<dgm:shape xmlns:r="`+nsOfficeRelsTransitional+`" type="round2SameRect" r:blip="rIdPicture"><dgm:adjLst/></dgm:shape>`, 1)
	deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{layout: layout}), nativeDiagramLayoutApproximateOptions())
	if err != nil {
		t.Fatal(err)
	}
	message := ""
	for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
		if diagnostic.Code == nativeDiagramLayoutDefinitionCode {
			message = diagnostic.Message
		}
	}
	if message == "" {
		t.Fatalf("diagram was not refused with %s: %+v", nativeDiagramLayoutDefinitionCode, deck.Slides[0].Compatibility.Diagnostics)
	}
	// The slide-level passthrough names the refused frame as well.
	if !strings.HasPrefix(message, "p:graphicFrame ") || !strings.Contains(message, "diagram image shapes are not modeled") {
		t.Fatalf("diagram refusal did not name the slide shape and the construct: %s", message)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid refused diagram deck: %#v", issues)
	}
}

// dgm:shape@rot is APPLIED, not refused: it is a clockwise angle in degrees
// about the shape's centre, the same rotation a:xfrm@rot carries in
// sixty-thousandths of a degree. smartart-autofit-sync.pptx rendered blank
// only because a rotated round2SameRect refused the whole frame.
func TestNativeDiagramLayoutAppliesShapeRotation(t *testing.T) {
	t.Parallel()
	layout := strings.Replace(nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional),
		`<dgm:shape type="rect"><dgm:adjLst/></dgm:shape>`,
		`<dgm:shape type="rect" rot="180"><dgm:adjLst/></dgm:shape>`, 1)
	deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{layout: layout}), nativeDiagramLayoutApproximateOptions())
	if err != nil {
		t.Fatalf("extract rotated diagram: %v", err)
	}
	rotated := 0
	for _, child := range nativeFixtureDiagramGroup(t, deck.Slides[0]).Children {
		if child.Transform.RotationAngle == nil {
			continue
		}
		if *child.Transform.RotationAngle != 10_800_000 {
			t.Fatalf("rot=180 became %d sixty-thousandths of a degree", *child.Transform.RotationAngle)
		}
		rotated++
	}
	if rotated == 0 {
		t.Fatal("the authored rotation did not reach any laid-out shape")
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid rotated diagram deck: %#v", issues)
	}
}

// A preset adjust value is APPLIED, not refused: dgm:adj@idx addresses the
// preset's own avLst guides in document order and dgm:adj@val is the fraction
// a:avLst stores as thousandths of a percent. A roundRect at adj 0.1 has to
// draw a different outline from one at the catalog default of 16667.
func TestNativeDiagramLayoutAppliesPresetAdjustValues(t *testing.T) {
	t.Parallel()
	outline := func(adjust string) string {
		t.Helper()
		layout := strings.Replace(nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional),
			`<dgm:layoutNode name="rootText1" styleLbl="node0"><dgm:varLst><dgm:chPref val="3"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="rect"><dgm:adjLst/></dgm:shape>`,
			`<dgm:layoutNode name="rootText1" styleLbl="node0"><dgm:varLst><dgm:chPref val="3"/></dgm:varLst><dgm:alg type="tx"/><dgm:shape type="roundRect"><dgm:adjLst>`+adjust+`</dgm:adjLst></dgm:shape>`, 1)
		if !strings.Contains(layout, "roundRect") {
			t.Fatal("layout fixture drifted")
		}
		deck, err := ExtractNativePPTX(nativeDiagramLayoutFixture(t, nativeDiagramLayoutFixtureOptions{layout: layout}), nativeDiagramLayoutApproximateOptions())
		if err != nil {
			t.Fatalf("extract adjusted diagram: %v", err)
		}
		for _, child := range nativeFixtureDiagramGroup(t, deck.Slides[0]).Children {
			if child.Name != nil && *child.Name == "rootText1" && child.Geometry != nil && len(child.Geometry.Paths) != 0 {
				return fmt.Sprint(child.Geometry.Paths[0].Commands)
			}
		}
		t.Fatal("no adjusted shape was laid out")
		return ""
	}
	authored := outline(`<dgm:adj idx="1" val="0.1"/>`)
	if authored == outline("") {
		t.Fatal("the authored adjust value did not change the outline")
	}
	// An index the preset does not have refuses rather than drawing a default.
	layout := strings.Replace(nativeDiagramLayoutLayoutXML(nativeDiagramURITransitional),
		`<dgm:shape type="rect"><dgm:adjLst/></dgm:shape>`,
		`<dgm:shape type="roundRect"><dgm:adjLst><dgm:adj idx="4" val="0.1"/></dgm:adjLst></dgm:shape>`, 1)
	assertNativeDiagramLayoutRefused(t, nativeDiagramLayoutFixtureOptions{layout: layout}, nativeDiagramLayoutDefinitionCode)
}

// The preview transport caps a single diagnostic string at 2048 characters
// (apps/pptx-page-paint-worker/src/contract.ts) and prefixes the message with
// its code. A message that overruns that cap does not degrade: the whole
// slide preview fails bounded validation with HTTP 422, so EVERY diagram this
// tier lays out paints nothing. Growing the disclosure has to stay inside the
// budget, and 2048 is a hard wall, not a style rule.
func TestNativeDiagramLayoutDisclosureFitsTheTransportCap(t *testing.T) {
	const transportCap = 2048
	for _, message := range []struct{ code, text string }{
		{nativeDiagramLayoutPreviewCode, nativeDiagramLayoutGroupMessage},
		{nativeDiagramLayoutPreviewCode, nativeDiagramLayoutChildMessage},
	} {
		if size := len(message.code) + len(": ") + len(message.text); size > transportCap {
			t.Fatalf("%s diagnostic is %d characters, over the %d the preview transport accepts", message.code, size, transportCap)
		}
	}
}
