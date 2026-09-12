package pptxpatch

import (
	"bytes"
	"testing"
)

func TestNativePictureRoundRectClipRequiresExactDefaultPreset(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			geometry string
			valid    bool
		}{
			{`<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>`, true},
			{`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 16667"/></a:avLst></a:prstGeom>`, false},
			{`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>`, false},
			{`<a:prstGeom prst="roundRect" extra="1"><a:avLst/></a:prstGeom>`, false},
			{`<a:prstGeom prst="roundRect"><a:avLst extra="1"/></a:prstGeom>`, false},
			{`<a:prstGeom prst="roundRect">text<a:avLst/></a:prstGeom>`, false},
			{`<a:prstGeom prst="roundRect"><a:avLst/><a:avLst/></a:prstGeom>`, false},
			{`<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`, false},
		} {
			data := nativePictureFixture(t, nativePictureFixtureOptions{strict: strict, pictureGeometry: tc.geometry, sourceRect: `<a:srcRect l="10000" t="20000" r="30000" b="0"/>`})
			before := bytes.Clone(data)
			deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
			if err != nil {
				if tc.valid {
					t.Fatal(err)
				}
				continue
			}
			picture := nativeFixturePicture(t, deck.Slides[0])
			if (picture.Clip != nil) != tc.valid {
				t.Fatalf("clip=%#v geometry%s", picture.Clip, tc.geometry)
			}
			if tc.valid && (*picture.Clip != "roundRect" || picture.Crop == nil || *picture.Crop.Left != 10000) {
				t.Fatal("lost source clip/crop")
			}
			if !tc.valid && picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatal("unsupported mask accepted")
			}
			if !bytes.Equal(data, before) {
				t.Fatal("source changed")
			}
			if tc.valid {
				bad := "ellipse"
				picture.Clip = &bad
				deck.Slides[0].Elements = []NativeElement{picture}
				if len(ValidateNativePPTX(deck)) == 0 {
					t.Fatal("unknown clip accepted")
				}
			}
		}
	}
}
