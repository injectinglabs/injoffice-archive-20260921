package pptxpatch

import (
	"strings"
	"testing"
)

func TestNativeDefaultPentagonRetainsQualifiedText(t *testing.T) {
	for _, strict := range []bool{false, true} {
		shape := nativeAutoShapeXML(3, "Pentagon", "pentagon", `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
		shape = strings.Replace(shape, `</p:sp>`, `<p:txBody><a:bodyPr wrap="square" lIns="0" rIns="0" tIns="0" bIns="0"><a:noAutofit/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>AB</a:t></a:r></a:p></p:txBody></p:sp>`, 1)
		for _, adjusted := range []bool{false, true} {
			candidate := shape
			if adjusted {
				candidate = strings.Replace(candidate, `<a:avLst/>`, `<a:avLst><a:gd name="hf" fmla="val 105146"/></a:avLst>`, 1)
			}
			deck, err := ExtractNativePPTX(nativeShapeStyleFixture(t, strict, candidate, `<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>`), nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			element := nativeFixtureAutoShapes(deck.Slides[0])[0]
			if adjusted {
				if element.Preset != nil {
					t.Fatal("adjusted preset qualified")
				}
			} else if element.Preset == nil || element.TextBody == nil || element.Paragraphs == nil || len(*element.Paragraphs) != 1 || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("qualified text lost: %+v", element)
			}
		}
	}
}
