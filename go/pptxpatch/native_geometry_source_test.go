package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

const nativeGeometrySourceXML = `<a:custGeom><a:avLst><a:gd name="adj" fmla="val 1000"/></a:avLst><a:gdLst><a:gd name="inset" fmla="*/ adj 1 2"/></a:gdLst><a:ahLst/><a:cxnLst/><a:rect l="inset" t="inset" r="w" b="h"/><a:pathLst><a:path w="100" h="100" extrusionOk="true"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:quadBezTo><a:pt x="50" y="100"/><a:pt x="100" y="0"/></a:quadBezTo><a:cubicBezTo><a:pt x="100" y="25"/><a:pt x="50" y="75"/><a:pt x="0" y="100"/></a:cubicBezTo><a:close/></a:path><a:path fill="none" stroke="true"><a:moveTo><a:pt x="2000" y="0"/></a:moveTo><a:arcTo wR="2000" hR="1000" stAng="0" swAng="21600000"/><a:close/></a:path></a:pathLst></a:custGeom>`

func TestNativeGeometrySourceExtractionAndMutationRefusal(t *testing.T) {
	for _, strict := range []bool{false, true} {
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			shape := nativeAutoShapeXMLWithGeometry(3, nativeGeometrySourceXML, `<a:solidFill><a:srgbClr val="336699"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
			original := nativeShapeStyleFixture(t, strict, shape, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`)
			before := bytes.Clone(original)
			deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			e := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if e.Geometry == nil || e.Preset != nil || e.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(e.Passthrough) != 1 {
				t.Fatalf("custom source not preserve-only: %+v", e)
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatal(issues)
			}
			data, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatal(err)
			}
			if _, err := DecodeNativePPTXJSON(data); err != nil {
				t.Fatal(err)
			}
			update := NativePPTXAutoShapeMutation{Transform: e.Transform, Preset: NativeShapePresetRect, Fill: e.Fill, Stroke: e.Stroke}
			_, err = ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "shape", Kind: NativePPTXUpdateAutoShape, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, AutoShape: &update}}})
			if err == nil || !strings.Contains(err.Error(), "preview-only") {
				t.Fatalf("custom mutation not refused: %v", err)
			}
			if !bytes.Equal(original, before) {
				t.Fatal("source bytes changed")
			}
			for _, bad := range []string{strings.Replace(nativeGeometrySourceXML, `<a:close/>`, `<a:unknown/>`, 1), strings.Replace(nativeGeometrySourceXML, `<a:path w=`, `<a:path unsupported="1" w=`, 1)} {
				candidate := nativeAutoShapeXMLWithGeometry(3, bad, `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
				d, err := ExtractNativePPTX(nativeShapeStyleFixture(t, strict, candidate, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`), nativeTestExtractOptions())
				if err != nil {
					t.Fatal(err)
				}
				got := nativeFixtureAutoShapes(d.Slides[0])[0]
				if got.Geometry != nil || got.Compatibility.Status != NativeCompatibilityStatusRefused {
					t.Fatal("partial custom geometry accepted")
				}
			}
		})
	}
}
