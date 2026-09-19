package pptxpatch

import (
	"strings"
	"testing"
)

// customshape-bitmapfill-srcrect.pptx in miniature: the slide picture declares
// only p:ph and an empty p:spPr, and the layout placeholder carries the box and
// the outline PowerPoint paints it in.
func nativePictureInheritedBoxFixture(t *testing.T, layoutShapeProperties string) []byte {
	t.Helper()
	const picture = `<p:pic><p:nvPicPr><p:cNvPr id="3" name="Content Placeholder 5"/><p:cNvPicPr><a:picLocks noGrp="1"/></p:cNvPicPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>` +
		`<p:blipFill><a:blip xmlns:r="` + nsOfficeRelsTransitional + `" r:embed="rIdImage"/><a:srcRect l="4393" r="4393"/><a:stretch/></p:blipFill><p:spPr/></p:pic>`
	layoutPlaceholder := `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>` +
		layoutShapeProperties + `</p:sp>`
	return nativeExtractFixture(t, nativeExtractFixtureOptions{
		extraParts: []nativeExtractZipPart{{name: "relocated/media/image1.png", data: "\x89PNG\r\n\x1a\nfixture"}},
		mutate: func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, picture+`</p:spTree>`, 1)
			parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`,
				`<Relationship Id="rIdImage" Type="`+relImageTransitional+`" Target="../media/image1.png"/></Relationships>`, 1)
			parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], `</p:spTree>`,
				layoutPlaceholder+`</p:spTree>`, 1)
			parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], `<p:sldLayout xmlns:p="`+nsPresentationTransitional+`">`,
				`<p:sldLayout xmlns:p="`+nsPresentationTransitional+`" xmlns:a="`+nsDrawingTransitional+`">`, 1)
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
				`<Override PartName="/relocated/media/image1.png" ContentType="image/png"/></Types>`, 1)
		},
	})
}

// The read-only preview paints the picture in the box its placeholder chain
// declares. Before this lane customshape-bitmapfill-srcrect.pptx was a blank
// page against 26.2% LibreOffice ink, because the only shape on the slide was
// refused over a transform the layout states two parts away.
func TestNativePicturePlaceholderPaintsInheritedBox(t *testing.T) {
	t.Parallel()
	const layout = `<p:spPr><a:xfrm><a:off x="6192000" y="1332000"/><a:ext cx="5493600" cy="4012789"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 10813"/></a:avLst></a:prstGeom></p:spPr>`
	input := nativePictureInheritedBoxFixture(t, layout)

	options := nativeAtomicTestExtractOptions()
	strict, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	// Strict is unchanged by construction: the projection is never consulted.
	if len(strict.Slides[0].Elements) != 1 || strict.Slides[0].Elements[0].Kind != NativeElementKindText {
		t.Fatalf("the exact tier painted an inherited picture box: %#v", strict.Slides[0].Elements)
	}
	if !nativeDiagnosticsContain(strict.Slides[0].Compatibility.Diagnostics, "pptx.unsupported-picture") {
		t.Fatalf("the exact tier lost its refusal: %#v", strict.Slides[0].Compatibility)
	}

	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	if picture.Transform.X == nil || *picture.Transform.X != 6192000 || picture.Transform.Y == nil || *picture.Transform.Y != 1332000 {
		t.Fatalf("inherited offset was not projected: %#v", picture.Transform)
	}
	if picture.Transform.Cx == nil || *picture.Transform.Cx != 5493600 || picture.Transform.Cy == nil || *picture.Transform.Cy != 4012789 {
		t.Fatalf("inherited extent was not projected: %#v", picture.Transform)
	}
	// The picture keeps its own crop, and takes the ancestor's outline because
	// it declares no geometry of its own.
	if picture.Crop == nil || picture.Crop.Left == nil || *picture.Crop.Left != 4393 || picture.Crop.Right == nil || *picture.Crop.Right != 4393 {
		t.Fatalf("the picture's own crop was lost: %#v", picture.Crop)
	}
	if picture.Geometry == nil || len(picture.Geometry.Paths) == 0 {
		t.Fatalf("inherited a:prstGeom was not evaluated: %#v", picture.Geometry)
	}
	if picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("an approximated picture stayed editable: %q", picture.Compatibility.Status)
	}
	if !nativeDiagnosticsContain(picture.Compatibility.Diagnostics, nativePicturePlaceholderFrameCode) {
		t.Fatalf("the inherited box was not disclosed: %#v", picture.Compatibility.Diagnostics)
	}
}

// An ancestor that states no box leaves the refusal in place: the preview
// proposes only what the chain actually declares, and never invents a frame.
func TestNativePicturePlaceholderWithoutAnInheritedBoxStillRefuses(t *testing.T) {
	t.Parallel()
	input := nativePictureInheritedBoxFixture(t, `<p:spPr/>`)
	options := nativeAtomicTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 || deck.Slides[0].Elements[0].Kind != NativeElementKindText {
		t.Fatalf("a boxless placeholder chain invented a frame: %#v", deck.Slides[0].Elements)
	}
	if !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.unsupported-picture") {
		t.Fatalf("the picture refusal was lost: %#v", deck.Slides[0].Compatibility)
	}
}
