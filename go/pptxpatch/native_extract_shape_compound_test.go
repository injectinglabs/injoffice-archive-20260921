package pptxpatch

import (
	"fmt"
	"testing"
)

func nativeCompoundLine(compound string) string {
	return fmt.Sprintf(`<a:ln w="63500" cap="flat" cmpd="%s" algn="ctr"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`, compound)
}

// ECMA-376 Part 1 §20.1.10.14 ST_CompoundLine. A compound outline paints
// several concentric lines inside the one authored width; the renderer lays
// them out by moving the shape outline, which it can only do for a rectangle.
func TestExtractNativePPTXAutoShapeCompoundOutlineOnRectangle(t *testing.T) {
	t.Parallel()
	for source, want := range map[string]NativeStrokeCompound{
		"dbl":       NativeStrokeCompoundDouble,
		"thickThin": NativeStrokeCompoundThickThin,
		"thinThick": NativeStrokeCompoundThinThick,
		"tri":       NativeStrokeCompoundTriple,
	} {
		source, want := source, want
		t.Run(source, func(t *testing.T) {
			t.Parallel()
			shape := nativeAutoShapeXML(3, "Rectangle", "rect", `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`, nativeCompoundLine(source), "")
			deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract compound outline: %v", err)
			}
			shapes := nativeFixtureAutoShapes(deck.Slides[0])
			if len(shapes) != 1 || shapes[0].Compatibility.Status != NativeCompatibilityStatusEditable {
				t.Fatalf("compound outline on a rectangle was refused: %#v", shapes)
			}
			stroke := shapes[0].Stroke
			if stroke == nil || stroke.Compound == nil || *stroke.Compound != want || stroke.WidthEMU == nil || *stroke.WidthEMU != 63500 || stroke.Color != "7030A0" {
				t.Fatalf("compound outline did not survive extraction: %#v", stroke)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("compound outline contract is invalid: %#v", issues)
			}
		})
	}
}

// A single line is the ST_CompoundLine default and stays unstated, so every
// deck that predates compound outlines keeps the contract it already had.
func TestExtractNativePPTXSingleOutlineLeavesCompoundUnstated(t *testing.T) {
	t.Parallel()
	for _, line := range []string{nativeCompoundLine("sng"), `<a:ln w="63500" cap="flat" algn="ctr"><a:solidFill><a:srgbClr val="7030A0"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`} {
		shape := nativeAutoShapeXML(3, "Rectangle", "rect", `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`, line, "")
		deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("extract single outline: %v", err)
		}
		stroke := nativeFixtureAutoShapes(deck.Slides[0])[0].Stroke
		if stroke == nil || stroke.Compound != nil {
			t.Fatalf("a single outline must not state a compound: %#v", stroke)
		}
	}
}

func TestExtractNativePPTXCompoundOutlineRefusedOffRectangle(t *testing.T) {
	t.Parallel()
	for name, shape := range map[string]string{
		"ellipse":  nativeAutoShapeXML(3, "Ellipse", "ellipse", `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`, nativeCompoundLine("thinThick"), ""),
		"triangle": nativeAutoShapeXML(3, "Triangle", "triangle", `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`, nativeCompoundLine("thinThick"), ""),
		"unknown":  nativeAutoShapeXML(3, "Rectangle", "rect", `<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`, nativeCompoundLine("quad"), ""),
	} {
		name, shape := name, shape
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract refused compound outline: %v", err)
			}
			element := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.Stroke != nil {
				t.Fatalf("a compound outline the renderer cannot lay out must refuse: %#v", element)
			}
			found := false
			for _, diagnostic := range element.Compatibility.Diagnostics {
				if diagnostic.Code == "pptx.autoshape-line-unavailable" && diagnostic.Severity == NativeDiagnosticSeverityRefusal {
					found = true
				}
			}
			if !found {
				t.Fatalf("missing outline refusal: %#v", element.Compatibility.Diagnostics)
			}
		})
	}
}

// A connector shaft is a line, never a rectangle, so it cannot carry the bands.
func TestExtractNativePPTXCompoundConnectorLineIsRefused(t *testing.T) {
	t.Parallel()
	connector := `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>` + nativeCompoundLine("thinThick") + `</p:spPr></p:cxnSp>`
	deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, connector), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract compound connector: %v", err)
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind != NativeElementKindConnector {
			continue
		}
		if element.Compatibility.Status != NativeCompatibilityStatusRefused {
			t.Fatalf("a compound connector line must refuse: %#v", element)
		}
		return
	}
	t.Fatal("connector element disappeared")
}
