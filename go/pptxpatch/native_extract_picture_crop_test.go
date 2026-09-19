package pptxpatch

import (
	"strings"
	"testing"
)

// The projection is arithmetic, so it is measured directly rather than only
// through a package. croppedTo0.pptx is the first row.
func TestProjectNativeApproximatePictureCrop(t *testing.T) {
	t.Parallel()
	tests := []struct {
		name    string
		insets  [4]int64
		cx, cy  int64
		ok      bool
		painted *[4]int64
	}{
		{
			// croppedTo0.pptx "Picture 2": one thousandth of one percent of
			// outset, 35.7 EMU on a 2043991 EMU edge = 0.004 device pixels.
			name: "rounding crumb", insets: [4]int64{42751, 0, -1, 12700}, cx: 2043991, cy: 2352675,
			ok: true, painted: &[4]int64{42751, 0, 0, 12700},
		},
		{
			// croppedTo0.pptx "Picture 4": l+r is exactly the whole image, so
			// the crop samples nothing. There is no nearest representable crop.
			name: "degenerate axis", insets: [4]int64{60000, 70000, 40000, 30000}, cx: 0, cy: 0, ok: false,
		},
		{
			// A real outset frame: a tenth of the destination edge would be
			// blank, which this reader cannot paint.
			name: "visible outset", insets: [4]int64{0, 0, -10000, 0}, cx: 2043991, cy: 2352675, ok: false,
		},
		{
			// The bound is the painted edge, not the inset: the same outset on
			// a picture small enough to hide it is projected.
			// It clamps to nothing, so the picture is simply uncropped.
			name: "outset under a pixel on a small picture", insets: [4]int64{0, 0, -1000, 0}, cx: 9000, cy: 9000, ok: true,
		},
		{
			// Every authored inset was an outset crumb, so nothing is cropped.
			name: "all crumbs", insets: [4]int64{-1, -1, -1, -1}, cx: 2043991, cy: 2352675, ok: true,
		},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			crop, ok := projectNativeApproximatePictureCrop(test.insets, test.cx, test.cy)
			if ok != test.ok {
				t.Fatalf("projection = %v, want %v", ok, test.ok)
			}
			if !ok {
				return
			}
			if test.painted == nil {
				if crop != nil {
					t.Fatalf("an uncropped projection still carries a crop: %+v", crop)
				}
				return
			}
			if crop == nil {
				t.Fatal("projection dropped the crop it kept")
			}
			got := [4]int64{*crop.Left, *crop.Top, *crop.Right, *crop.Bottom}
			if got != *test.painted {
				t.Fatalf("painted insets %v, want %v", got, *test.painted)
			}
		})
	}
}

// End to end: the picture paints, and the diagnostic names both the authored
// and the painted insets so the move is auditable without the source.
func TestExtractNativePPTXPaintsACropClampedUnderOneDevicePixel(t *testing.T) {
	t.Parallel()
	const imagePart = "relocated/media/image1.png"
	picture := func(srcRect string) string {
		return `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Picture 2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
			`<p:blipFill><a:blip xmlns:r="` + nsOfficeRelsTransitional + `" r:embed="rIdImage"/>` + srcRect + `<a:stretch/></p:blipFill>` +
			`<p:spPr><a:xfrm><a:off x="7042858" y="66675"/><a:ext cx="2043991" cy="2352675"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
	}
	build := func(t *testing.T, srcRect string) NativeSlide {
		t.Helper()
		input := nativeExtractFixture(t, nativeExtractFixtureOptions{
			extraParts: []nativeExtractZipPart{{name: imagePart, data: "\x89PNG\r\n\x1a\nfixture"}},
			mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, picture(srcRect)+`</p:spTree>`, 1)
				parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`,
					`<Relationship Id="rIdImage" Type="`+relImageTransitional+`" Target="../media/image1.png"/></Relationships>`, 1)
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
					`<Override PartName="/`+imagePart+`" ContentType="image/png"/></Types>`, 1)
			},
		})
		deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		return deck.Slides[0]
	}

	painted := nativeFixturePicture(t, build(t, `<a:srcRect l="42751" r="-1" b="12700"/>`))
	if painted.Crop == nil || *painted.Crop.Left != 42751 || *painted.Crop.Right != 0 || *painted.Crop.Bottom != 12700 {
		t.Fatalf("the clamped crop was not painted: %+v", painted.Crop)
	}
	if nativeDiagnosticsContain(painted.Compatibility.Diagnostics, "pptx.picture-crop-unavailable") {
		t.Fatalf("a painted crop still claims it is unavailable: %+v", painted.Compatibility.Diagnostics)
	}
	if !nativeDiagnosticsContain(painted.Compatibility.Diagnostics, nativePictureCropCode) {
		t.Fatalf("the clamp was not disclosed: %+v", painted.Compatibility.Diagnostics)
	}
	for _, want := range []string{`l=42751 t=0 r=-1 b=12700`, `l=42751 t=0 r=0 b=12700`} {
		found := false
		for _, diagnostic := range painted.Compatibility.Diagnostics {
			if diagnostic.Code == nativePictureCropCode && strings.Contains(diagnostic.Message, want) {
				found = true
			}
		}
		if !found {
			t.Fatalf("the disclosure does not name %q: %+v", want, painted.Compatibility.Diagnostics)
		}
	}

	// An outset the viewer would see is still refused, with the message it had.
	refused := nativeFixturePicture(t, build(t, `<a:srcRect r="-10000"/>`))
	if refused.Crop != nil {
		t.Fatalf("a visible outset was painted as a crop: %+v", refused.Crop)
	}
	if !nativeDiagnosticsContain(refused.Compatibility.Diagnostics, "pptx.picture-crop-unavailable") {
		t.Fatalf("a visible outset lost its refusal: %+v", refused.Compatibility.Diagnostics)
	}
}
