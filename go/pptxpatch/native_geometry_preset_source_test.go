package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestNativePresetSourceAdjustments(t *testing.T) {
	for _, ns := range []string{nativePresetDrawingNS, "http://purl.oclc.org/ooxml/drawingml/main"} {
		parse := func(body string) (*NativeEvaluatedGeometry, error) {
			node, err := parseNativeXML([]byte(fmt.Sprintf(`<a:prstGeom xmlns:a="%s" prst="triangle">%s</a:prstGeom>`, ns, body)), "preset.xml")
			if err != nil {
				return nil, err
			}
			return evaluateNativePresetSource(node, ns, 4000000, 3000000)
		}
		geometry, err := parse(`<a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst>`)
		if err != nil {
			t.Fatal(err)
		}
		if *geometry.Paths[0].Commands[1].X != 1000000 {
			t.Fatal("source adjustment not evaluated")
		}
		normal, err := parse("")
		if err != nil {
			t.Fatal(err)
		}
		if *normal.Paths[0].Commands[1].X != 2000000 {
			t.Fatal("omitted avLst did not use defaults")
		}
		for _, bad := range []string{`<a:avLst><a:gd name="unknown" fmla="val 1"/></a:avLst>`, `<a:avLst><a:gd name="adj" fmla="*/ 1 2 3"/></a:avLst>`, `<a:avLst extra="1"/>`, `<a:avLst/><a:avLst/>`, `<a:extLst/>`, `<a:avLst><a:gd name="adj" fmla="val 9007199254740992"/></a:avLst>`} {
			if geometry, err := parse(bad); err == nil || geometry != nil {
				t.Fatalf("invalid source override accepted: %s", bad)
			}
		}
	}
}

func TestNativePresetCatalogPublicSource(t *testing.T) {
	names, err := NativePPTXPresetNames()
	if err != nil {
		t.Fatal(err)
	}
	for _, strict := range []bool{false, true} {
		for _, name := range names {
			shape := nativeAutoShapeXMLWithGeometry(3, `<a:prstGeom prst="`+name+`"><a:avLst/></a:prstGeom>`, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
			shape = strings.Replace(shape, `cx="1000000" cy="500000"`, `cx="4000000" cy="3000000"`, 1)
			deck, err := ExtractNativePPTX(nativeShapeStyleFixture(t, strict, shape, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("%s: %v", name, err)
			}
			e := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if e.Compatibility.Status == NativeCompatibilityStatusRefused || (e.Geometry == nil && e.Preset == nil) {
				t.Fatalf("%s strict=%v: %+v", name, strict, e.Compatibility)
			}
			if e.Geometry != nil && (e.Preset != nil || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(e.Passthrough) != 1) {
				t.Fatalf("%s catalog authority", name)
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatalf("%s: %v", name, issues)
			}
			if name == "rect" && (e.Preset == nil || *e.Preset != NativeShapePresetRect || e.Compatibility.Status != NativeCompatibilityStatusEditable) {
				t.Fatal("legacy editable rectangle changed")
			}
		}
	}
}
func TestNativePresetAdjustedSourceAndNegative(t *testing.T) {
	for _, source := range []struct {
		xml   string
		valid bool
	}{
		{`<a:prstGeom prst="triangle"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>`, true},
		{`<a:prstGeom prst="triangle"><a:avLst><a:gd name="unknown" fmla="val 25000"/></a:avLst></a:prstGeom>`, false},
		{`<a:prstGeom prst="notAShape"><a:avLst/></a:prstGeom>`, false},
	} {
		shape := nativeAutoShapeXMLWithGeometry(3, source.xml, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
		deck, err := ExtractNativePPTX(nativeShapeStyleFixture(t, false, shape, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`), nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		e := nativeFixtureAutoShapes(deck.Slides[0])[0]
		if !source.valid {
			if e.Geometry != nil || e.Compatibility.Status != NativeCompatibilityStatusRefused {
				t.Fatal("invalid preset painted")
			}
			continue
		}
		if e.Geometry == nil || *e.Geometry.Paths[0].Commands[1].X != 250000 || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatal("adjusted source not evaluated")
		}
	}
}
