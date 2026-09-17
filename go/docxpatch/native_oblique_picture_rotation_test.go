package docxpatch

import (
	"strings"
	"testing"
)

// obliqueAnchor is Word's own markup for a rotated floating picture: wp:extent
// stays the unrotated box, a:xfrm/@rot states an angle that is not a quarter
// turn, and wp:effectExtent carries the envelope the rotation reaches into.
const obliqueAnchorPrefix = `<wp:anchor simplePos="0" relativeHeight="7" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" distT="0" distB="0" distL="0" distR="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>1270000</wp:posOffset></wp:positionV>`

func obliquePictureDocument(t *testing.T, mutate func(string) string) *NativeDocumentV1 {
	t.Helper()
	parts := transitionalNativeParts()
	xml := strings.Replace(parts["Custom/Main.XML"], `<wp:inline distT="0" distB="0" distL="0" distR="0">`, obliqueAnchorPrefix, 1)
	xml = strings.Replace(xml, `<wp:effectExtent l="0" t="0" r="0" b="0"/>`, `<wp:effectExtent l="127000" t="88900" r="127000" b="88900"/><wp:wrapSquare wrapText="bothSides"/>`, 1)
	xml = strings.Replace(xml, `</wp:inline>`, `</wp:anchor>`, 1)
	xml = strings.Replace(xml, `<a:xfrm/>`, `<a:xfrm rot="641099"><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>`, 1)
	if mutate != nil {
		xml = mutate(xml)
	}
	parts["Custom/Main.XML"] = xml
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	return doc
}

func obliqueDrawing(doc *NativeDocumentV1) *NativeDrawingV1 {
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		if run.Drawing != nil {
			return run.Drawing
		}
	}
	return nil
}

func TestNativeObliquePictureRotationProjectsExactAngleAndEnvelope(t *testing.T) {
	drawing := obliqueDrawing(obliquePictureDocument(t, nil))
	if drawing == nil {
		t.Fatal("an obliquely rotated picture was refused")
	}
	if drawing.RotationDegrees != nil {
		t.Fatalf("an oblique angle must not claim a whole-degree projection: %v", *drawing.RotationDegrees)
	}
	if drawing.RotationAngle60000ths == nil || *drawing.RotationAngle60000ths != 641099 {
		t.Fatalf("oblique rotation lost its exact source angle: %#v", drawing.RotationAngle60000ths)
	}
	// 641099 is 10.685 degrees: the old whole-degree projection truncated to 10,
	// which is 0.685 degrees of drift across the whole picture.
	envelope := drawing.FloatingEffectExtentEMU
	if envelope == nil || *envelope.Left != 127000 || *envelope.Top != 88900 || *envelope.Right != 127000 || *envelope.Bottom != 88900 {
		t.Fatalf("floating rotation envelope was not projected: %#v", envelope)
	}
	if drawing.InlineEffectExtentEMU != nil {
		t.Fatal("a floating envelope must not be projected as an inline layout box")
	}
	if *drawing.WidthEMU != 914400 || *drawing.HeightEMU != 457200 {
		t.Fatal("the rotation envelope must not resize the painted box")
	}
}

func TestNativeQuarterTurnKeepsWholeDegreeProjection(t *testing.T) {
	drawing := obliqueDrawing(obliquePictureDocument(t, func(xml string) string {
		return strings.Replace(xml, `<a:xfrm rot="641099"><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>`, `<a:xfrm rot="5400000"><a:off x="0" y="0"/><a:ext cx="457200" cy="914400"/></a:xfrm>`, 1)
	}))
	if drawing == nil || drawing.RotationDegrees == nil || *drawing.RotationDegrees != 90 {
		t.Fatalf("quarter turn lost its whole-degree projection: %#v", drawing)
	}
	if drawing.RotationAngle60000ths == nil || *drawing.RotationAngle60000ths != 5400000 {
		t.Fatalf("quarter turn lost its exact angle: %#v", drawing)
	}
}

func TestNativeObliquePictureRotationRefusals(t *testing.T) {
	for _, tc := range []struct{ name, from, to string }{
		// wp:extent must stay the unrotated box for an oblique rotation; a
		// transposed shape extent describes a frame this projection cannot map.
		{"transposedExtent", `<a:ext cx="914400" cy="457200"/></a:xfrm>`, `<a:ext cx="457200" cy="914400"/></a:xfrm>`},
		{"missingExtent", `<a:xfrm rot="641099"><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>`, `<a:xfrm rot="641099"/>`},
		{"fullTurn", `rot="641099"`, `rot="21600000"`},
		{"negativeAngle", `rot="641099"`, `rot="-641099"`},
		{"paddedAngle", `rot="641099"`, `rot="0641099"`},
		{"lexicalAngle", `rot="641099"`, `rot=" 641099 "`},
		// The envelope is only inert because the picture shape itself carries no
		// effect: a recolouring mode, a fill, an outline or an effect list keeps
		// the whole drawing preserve-only.
		{"blackWhiteMode", `<pic:spPr>`, `<pic:spPr bwMode="gray">`},
		{"solidFill", `<a:prstGeom prst="rect"/>`, `<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>`},
		{"outline", `<a:prstGeom prst="rect"/>`, `<a:prstGeom prst="rect"/><a:ln w="12700"/>`},
		{"effectList", `<a:prstGeom prst="rect"/>`, `<a:prstGeom prst="rect"/><a:effectLst><a:outerShdw blurRad="50800"/></a:effectLst>`},
		{"nonEmptyNoFill", `<a:prstGeom prst="rect"/>`, `<a:prstGeom prst="rect"/><a:noFill><a:extra/></a:noFill>`},
		{"twoNoFills", `<a:prstGeom prst="rect"/>`, `<a:prstGeom prst="rect"/><a:noFill/><a:noFill/>`},
		{"unboundedEnvelope", `l="127000" t="88900" r="127000" b="88900"`, `l="91440001" t="88900" r="127000" b="88900"`},
		{"negativeEnvelope", `l="127000" t="88900" r="127000" b="88900"`, `l="-127000" t="88900" r="127000" b="88900"`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			doc := obliquePictureDocument(t, func(xml string) string {
				replaced := strings.Replace(xml, tc.from, tc.to, 1)
				if replaced == xml {
					t.Fatalf("fixture does not contain %q", tc.from)
				}
				return replaced
			})
			if drawing := obliqueDrawing(doc); drawing != nil {
				t.Fatalf("unsupported rotated picture accepted: %#v", drawing)
			}
			if len(doc.Unsupported) == 0 {
				t.Fatal("a refused rotated picture must state a preserved capability")
			}
		})
	}
}

func TestNativeInertPictureShapePropertiesAreAccepted(t *testing.T) {
	// Word writes bwMode="auto" (render normally) and a:noFill (no shape fill
	// behind a blip that already covers the frame) on ordinary pictures.
	drawing := obliqueDrawing(obliquePictureDocument(t, func(xml string) string {
		xml = strings.Replace(xml, `<pic:spPr>`, `<pic:spPr bwMode="auto">`, 1)
		return strings.Replace(xml, `<a:prstGeom prst="rect"/>`, `<a:prstGeom prst="rect"/><a:noFill/>`, 1)
	}))
	if drawing == nil || drawing.RotationAngle60000ths == nil || *drawing.RotationAngle60000ths != 641099 {
		t.Fatalf("inert picture shape properties were refused: %#v", drawing)
	}
}

// A drawing whose graphic payload is not a picture at all must be refused as
// that, not as an unqualified effect extent: PICTURE_GRAPHIC_REQUIRED names the
// content the preview omits, while DRAWING_EFFECTS_PRESERVED claimed the file
// carried an effect it does not have.
func TestNativeNonPictureAnchorWithEnvelopeStatesItsRealRefusal(t *testing.T) {
	doc := obliquePictureDocument(t, func(xml string) string {
		return strings.Replace(xml, `uri="http://schemas.openxmlformats.org/drawingml/2006/picture"`, `uri="`+nativeTextboxWPS+`"`, 1)
	})
	if obliqueDrawing(doc) != nil {
		t.Fatal("a wordprocessingShape payload must not project as a picture")
	}
	codes := map[string]bool{}
	for _, entry := range doc.Unsupported {
		codes[entry.Code] = true
	}
	if !codes["PICTURE_GRAPHIC_REQUIRED"] || codes["DRAWING_EFFECTS_PRESERVED"] {
		t.Fatalf("non-picture drawing reported the wrong capability: %#v", codes)
	}
}
