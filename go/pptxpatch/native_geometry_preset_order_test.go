package pptxpatch

import (
	"bytes"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"
)

func TestNativePresetLiteralAssignmentOrder(t *testing.T) {
	for _, ns := range []string{nativePresetDrawingNS, "http://purl.oclc.org/ooxml/drawingml/main"} {
		parse := func(entries string) (*NativeEvaluatedGeometry, error) {
			node, err := parseNativeXML([]byte(fmt.Sprintf(`<a:prstGeom xmlns:a="%s" prst="triangle"><a:avLst>%s</a:avLst></a:prstGeom>`, ns, entries)), "preset.xml")
			if err != nil {
				return nil, err
			}
			return evaluateNativePresetSource(node, ns, 4000000, 3000000)
		}
		repeated, err := parse(`<a:gd name="adj" fmla="val 75000"/><a:gd name="adj" fmla="val 25000"/>`)
		if err != nil {
			t.Fatal(err)
		}
		single, err := parse(`<a:gd name="adj" fmla="val 25000"/>`)
		if err != nil || !reflect.DeepEqual(repeated, single) || *repeated.Paths[0].Commands[1].X != 1000000 {
			t.Fatal("last source literal did not govern triangle apex")
		}
		for _, earlier := range []string{`<a:gd name="adj" fmla="val 9007199254740992"/>`, `<a:gd name="adj" fmla="*/ 1 2 3"/>`, `<a:gd name="adj" fmla="val missing"/>`, `<a:gd name="unknown" fmla="val 1"/>`, `<a:gd name="adj" fmla="val 1" extra="1"/>`} {
			if geometry, err := parse(earlier + `<a:gd name="adj" fmla="val 25000"/>`); err == nil || geometry != nil {
				t.Fatalf("invalid earlier assignment disappeared: %s", earlier)
			}
		}
		entry := `<a:gd name="adj" fmla="val 25000"/>`
		if _, err := parse(strings.Repeat(entry, nativeGeometryMaxGuides)); err != nil {
			t.Fatalf("bounded repeated names refused: %v", err)
		}
		if geometry, err := parse(strings.Repeat(entry, nativeGeometryMaxGuides+1)); err == nil || geometry != nil {
			t.Fatal("repeated names bypassed assignment budget")
		}
	}
}

func nativePresetOrderFixture(t *testing.T, strict, invalid bool) []byte {
	t.Helper()
	first := "75000"
	if invalid {
		first = "9007199254740992"
	}
	geometry := `<a:prstGeom prst="triangle"><a:avLst><a:gd name="adj" fmla="val ` + first + `"/><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>`
	shape := nativeAutoShapeXMLWithGeometry(3, geometry, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	return nativeShapeStyleFixture(t, strict, shape, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`)
}

func TestNativePresetLiteralOrderSource(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, invalid := range []bool{false, true} {
			input := nativePresetOrderFixture(t, strict, invalid)
			original := append([]byte(nil), input...)
			deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(input, original) {
				t.Fatal("source bytes changed")
			}
			shape := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if invalid {
				if shape.Geometry != nil || shape.Compatibility.Status != NativeCompatibilityStatusRefused {
					t.Fatal("invalid overwritten literal painted")
				}
				continue
			}
			if shape.Geometry == nil || shape.Preset != nil || shape.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(shape.Passthrough) != 1 || *shape.Geometry.Paths[0].Commands[1].X != 250000 {
				t.Fatal("source assignment order or authority changed")
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatal(issues)
			}
		}
	}
}

func TestNativePresetLiteralOrderBrowserFixtures(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_PRESET_ORDER_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional actual browser fixture")
	}
	if err := os.MkdirAll(dir, 0700); err != nil {
		t.Fatal(err)
	}
	for _, item := range []struct {
		name    string
		invalid bool
	}{{"preset-order", false}, {"preset-order-refused", true}} {
		input := nativePresetOrderFixture(t, false, item.invalid)
		deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		encoded, err := MarshalNativePPTXJSON(deck)
		if err != nil {
			t.Fatal(err)
		}
		if err = os.WriteFile(filepath.Join(dir, item.name+".pptx"), input, 0600); err != nil {
			t.Fatal(err)
		}
		if err = os.WriteFile(filepath.Join(dir, item.name+"-go.json"), encoded, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
