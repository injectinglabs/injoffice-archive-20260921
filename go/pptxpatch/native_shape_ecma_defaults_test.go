package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

// nativeECMADefaultShapeXML builds one AutoShape whose p:spPr carries the given
// attributes, geometry, and outline.
func nativeECMADefaultShapeXML(id int, spPrAttrs, geometry, line string) string {
	return fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="Shape %d"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr%s><a:xfrm><a:off x="%d" y="100000"/><a:ext cx="400000" cy="400000"/></a:xfrm>%s<a:solidFill><a:srgbClr val="FFFF7F"/></a:solidFill>%s</p:spPr></p:sp>`,
		id, id, spPrAttrs, id*500000, geometry, line)
}

func nativeECMADefaultRefusalCodes(element NativeElement) map[string]bool {
	codes := map[string]bool{}
	for _, diagnostic := range element.Compatibility.Diagnostics {
		if diagnostic.Severity == NativeDiagnosticSeverityRefusal {
			codes[diagnostic.Code] = true
		}
	}
	return codes
}

// TestExtractNativePPTXAutoShapeAppliesECMALineAndDisplayDefaults locks the
// layout-neutral omissions: a black-and-white display hint that keeps the
// shape's own colors, and the a:ln attributes and dash whose ECMA-376 defaults
// are the flat, single, centered, solid outline PowerPoint paints.
func TestExtractNativePPTXAutoShapeAppliesECMALineAndDisplayDefaults(t *testing.T) {
	t.Parallel()
	line := `<a:ln w="19080"><a:solidFill><a:srgbClr val="A0A060"/></a:solidFill><a:round/><a:headEnd/><a:tailEnd type="none"/></a:ln>`
	shape := nativeECMADefaultShapeXML(3, ` bwMode="auto"`, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`, line)
	deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract AutoShape: %v", err)
	}
	shapes := nativeFixtureAutoShapes(deck.Slides[0])
	if len(shapes) != 1 {
		t.Fatalf("expected one AutoShape, got %#v", shapes)
	}
	element := shapes[0]
	if element.Compatibility.Status != NativeCompatibilityStatusEditable {
		t.Fatalf("defaulted AutoShape was not editable: status=%q diagnostics=%#v", element.Compatibility.Status, element.Compatibility.Diagnostics)
	}
	if element.Stroke == nil {
		t.Fatalf("defaulted outline was dropped: %#v", element)
	}
	if element.Stroke.Cap == nil || *element.Stroke.Cap != NativeStrokeCapFlat {
		t.Fatalf("omitted a:ln@cap did not resolve to the flat ECMA default: %#v", element.Stroke)
	}
	if element.Stroke.Dash == nil || *element.Stroke.Dash != NativeStrokeDashSolid {
		t.Fatalf("omitted a:prstDash did not resolve to a solid outline: %#v", element.Stroke)
	}
	if element.Stroke.Join == nil || *element.Stroke.Join != NativeStrokeJoinRound || element.Stroke.WidthEMU == nil || *element.Stroke.WidthEMU != 19080 || element.Stroke.Color != "A0A060" {
		t.Fatalf("defaulted outline lost authored paint: %#v", element.Stroke)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid extracted deck: %#v", issues)
	}
}

// TestExtractNativePPTXAutoShapeStillRefusesNonNeutralLineAndDisplayMarkup keeps
// the tolerance honest: a value that is not the ECMA default, or a display mode
// that restates the paint, must still refuse rather than be defaulted away.
func TestExtractNativePPTXAutoShapeStillRefusesNonNeutralLineAndDisplayMarkup(t *testing.T) {
	t.Parallel()
	solid := `<a:solidFill><a:srgbClr val="A0A060"/></a:solidFill>`
	for _, testCase := range []struct {
		name      string
		spPrAttrs string
		line      string
		code      string
	}{
		{"grayscale display mode", ` bwMode="gray"`, `<a:ln w="19080">` + solid + `<a:round/></a:ln>`, "pptx.autoshape-properties-unavailable"},
		{"hidden display mode", ` bwMode="hidden"`, `<a:ln w="19080">` + solid + `<a:round/></a:ln>`, "pptx.autoshape-properties-unavailable"},
		{"named tail arrow", "", `<a:ln w="19080">` + solid + `<a:round/><a:tailEnd type="triangle"/></a:ln>`, "pptx.autoshape-line-unavailable"},
		{"compound outline", "", `<a:ln w="19080" cmpd="dbl">` + solid + `<a:round/></a:ln>`, "pptx.autoshape-line-unavailable"},
		{"inset outline alignment", "", `<a:ln w="19080" algn="in">` + solid + `<a:round/></a:ln>`, "pptx.autoshape-line-unavailable"},
		{"dashed outline", "", `<a:ln w="19080">` + solid + `<a:prstDash val="dash"/><a:round/></a:ln>`, "pptx.autoshape-dash-unavailable"},
	} {
		testCase := testCase
		t.Run(testCase.name, func(t *testing.T) {
			t.Parallel()
			shape := nativeECMADefaultShapeXML(3, testCase.spPrAttrs, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`, testCase.line)
			deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract AutoShape: %v", err)
			}
			shapes := nativeFixtureAutoShapes(deck.Slides[0])
			if len(shapes) != 1 {
				t.Fatalf("expected one AutoShape, got %#v", shapes)
			}
			element := shapes[0]
			if element.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatalf("non-neutral markup was tolerated: status=%q diagnostics=%#v", element.Compatibility.Status, element.Compatibility.Diagnostics)
			}
			if !nativeECMADefaultRefusalCodes(element)[testCase.code] {
				t.Fatalf("expected refusal %q, got %#v", testCase.code, element.Compatibility.Diagnostics)
			}
		})
	}
}

// TestExtractNativePPTXCustomGeometryIgnoresUnpaintedHandleAndConnectionLists
// locks that populated a:ahLst / a:cxnLst evaluate, because neither clause
// contributes a painted segment, while unknown markup inside them still refuses.
func TestExtractNativePPTXCustomGeometryIgnoresUnpaintedHandleAndConnectionLists(t *testing.T) {
	t.Parallel()
	guides := `<a:avLst><a:gd name="adj1" fmla="val 18750"/></a:avLst><a:gdLst><a:gd name="y1" fmla="*/ h adj1 100000"/></a:gdLst>`
	handles := `<a:ahLst><a:ahXY gdRefY="adj1" minY="-2147483647" maxY="2147483647"><a:pos x="l" y="y1"/></a:ahXY><a:ahPolar gdRefAng="adj1" minAng="0" maxAng="21600000"><a:pos x="hc" y="vc"/></a:ahPolar></a:ahLst>`
	connections := `<a:cxnLst><a:cxn ang="0"><a:pos x="r" y="vc"/></a:cxn><a:cxn ang="cd2"><a:pos x="l" y="vc"/></a:cxn></a:cxnLst>`
	paths := `<a:rect l="l" t="t" r="r" b="b"/><a:pathLst><a:path><a:moveTo><a:pt x="l" y="t"/></a:moveTo><a:lnTo><a:pt x="r" y="t"/></a:lnTo><a:lnTo><a:pt x="r" y="b"/></a:lnTo><a:close/></a:path></a:pathLst>`
	line := `<a:ln w="19080"><a:solidFill><a:srgbClr val="A0A060"/></a:solidFill><a:miter lim="800000"/><a:headEnd/><a:tailEnd/></a:ln>`
	for _, testCase := range []struct {
		name     string
		geometry string
		refused  bool
	}{
		{"populated handles and connections", `<a:custGeom>` + guides + handles + connections + paths + `</a:custGeom>`, false},
		{"unknown handle-list member", `<a:custGeom>` + guides + `<a:ahLst><a:ahRadial/></a:ahLst>` + connections + paths + `</a:custGeom>`, true},
		{"unknown connection-site attribute", `<a:custGeom>` + guides + handles + `<a:cxnLst><a:cxn ang="0" idx="2"><a:pos x="r" y="vc"/></a:cxn></a:cxnLst>` + paths + `</a:custGeom>`, true},
	} {
		testCase := testCase
		t.Run(testCase.name, func(t *testing.T) {
			t.Parallel()
			shape := nativeECMADefaultShapeXML(3, ` bwMode="auto"`, testCase.geometry, line)
			deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract AutoShape: %v", err)
			}
			shapes := nativeFixtureAutoShapes(deck.Slides[0])
			if len(shapes) != 1 {
				t.Fatalf("expected one AutoShape, got %#v", shapes)
			}
			element := shapes[0]
			refused := element.Compatibility.Status == NativeCompatibilityStatusRefused
			if refused != testCase.refused {
				t.Fatalf("status %q with diagnostics %#v", element.Compatibility.Status, element.Compatibility.Diagnostics)
			}
			if testCase.refused {
				if !nativeECMADefaultRefusalCodes(element)["pptx.autoshape-geometry-unavailable"] {
					t.Fatalf("expected a geometry refusal, got %#v", element.Compatibility.Diagnostics)
				}
				return
			}
			if element.Geometry == nil || len(element.Geometry.Paths) != 1 {
				t.Fatalf("custom geometry was not evaluated: %#v", element.Geometry)
			}
			if element.Stroke == nil || element.Fill == nil || *element.Fill != "FFFF7F" {
				t.Fatalf("evaluated shape lost its paint: %#v", element)
			}
			warnings := []string{}
			for _, diagnostic := range element.Compatibility.Diagnostics {
				warnings = append(warnings, diagnostic.Code)
			}
			if !strings.Contains(strings.Join(warnings, " "), "pptx.custom-geometry-preview") {
				t.Fatalf("evaluated geometry lost its declared preview policy: %#v", element.Compatibility.Diagnostics)
			}
		})
	}
}
