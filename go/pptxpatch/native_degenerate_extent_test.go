package pptxpatch

import (
	"strings"
	"testing"
)

// ECMA-376 ST_PositiveCoordinate is minInclusive 0, so a degenerate extent is
// conformant DrawingML: cy="0" is how a horizontal straight connector is
// authored, and a picture cropped to nothing keeps a zero box. Requiring 1
// turned one such shape into a package-level extraction failure, which removed
// every slide of the deck instead of the one degenerate region.
func TestExtractNativePPTXAcceptsDegenerateExtents(t *testing.T) {
	t.Parallel()
	line := nativeAutoShapeSolidLine("25400", "rnd", `<a:bevel/>`, "ABCDEF")

	t.Run("connector", func(t *testing.T) {
		t.Parallel()
		connector := nativeConnectorXML(3, "Horizontal rule", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", "", "", "")
		connector = strings.Replace(connector, `<a:ext cx="1000000" cy="500000"/>`, `<a:ext cx="6477000" cy="0"/>`, 1)
		deck, err := ExtractNativePPTX(nativeConnectorFixture(t, false, connector), nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatalf("extract zero-height connector: %v", err)
		}
		connectors := nativeFixtureConnectors(deck.Slides[0])
		if len(connectors) != 1 {
			t.Fatalf("zero-height connector disappeared: %#v", deck.Slides[0].Elements)
		}
		transform := connectors[0].Transform
		if transform.Cx == nil || *transform.Cx != 6477000 || transform.Cy == nil || *transform.Cy != 0 {
			t.Fatalf("connector extent changed: %#v", transform)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("zero-height connector deck is invalid: %#v", issues)
		}
	})

	t.Run("autoshape", func(t *testing.T) {
		t.Parallel()
		shape := nativeAutoShapeXML(3, "Zero height rule", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
		shape = strings.Replace(shape, `<a:ext cx="1000000" cy="500000"/>`, `<a:ext cx="2700000" cy="0"/>`, 1)
		deck, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, shape), nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatalf("extract zero-height AutoShape: %v", err)
		}
		shapes := nativeFixtureAutoShapes(deck.Slides[0])
		if len(shapes) != 1 {
			t.Fatalf("zero-height AutoShape disappeared: %#v", deck.Slides[0].Elements)
		}
		transform := shapes[0].Transform
		if transform.Cx == nil || *transform.Cx != 2700000 || transform.Cy == nil || *transform.Cy != 0 {
			t.Fatalf("AutoShape extent changed: %#v", transform)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("zero-height AutoShape deck is invalid: %#v", issues)
		}
	})

	t.Run("picture", func(t *testing.T) {
		t.Parallel()
		fixture := nativePictureFixture(t, nativePictureFixtureOptions{pictureExtent: `<a:ext cx="0" cy="0"/>`})
		deck, err := ExtractNativePPTX(fixture, nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatalf("extract zero-extent picture: %v", err)
		}
		pictures := []NativeElement{}
		for _, element := range deck.Slides[0].Elements {
			if element.Kind == NativeElementKindPicture {
				pictures = append(pictures, element)
			}
		}
		if len(pictures) != 1 {
			t.Fatalf("zero-extent picture disappeared: %#v", deck.Slides[0].Elements)
		}
		transform := pictures[0].Transform
		if transform.Cx == nil || *transform.Cx != 0 || transform.Cy == nil || *transform.Cy != 0 {
			t.Fatalf("picture extent changed: %#v", transform)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("zero-extent picture deck is invalid: %#v", issues)
		}
	})

	t.Run("negative extent still fails", func(t *testing.T) {
		t.Parallel()
		connector := nativeConnectorXML(3, "Negative rule", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", "", "", "")
		connector = strings.Replace(connector, `<a:ext cx="1000000" cy="500000"/>`, `<a:ext cx="6477000" cy="-1"/>`, 1)
		if _, err := ExtractNativePPTX(nativeConnectorFixture(t, false, connector), nativeAtomicTestExtractOptions()); err == nil {
			t.Fatal("a negative extent is outside ST_PositiveCoordinate and must stay refused")
		}
	})
}

// A group's child coordinate space is a divisor of the group affine, so it
// keeps the stricter positive bound that the shared transform schema no longer
// states on its own.
func TestValidateNativePPTXRejectsZeroGroupChildExtent(t *testing.T) {
	t.Parallel()
	shape := nativeAutoShapeXML(4, "Grouped sentinel", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Zero child group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="600" cy="800"/><a:chOff x="100" y="200"/><a:chExt cx="0" cy="400"/></a:xfrm></p:grpSpPr>` + shape + `</p:grpSp>`
	if _, err := ExtractNativePPTX(nativeAutoShapeFixture(t, false, group), nativeAtomicTestExtractOptions()); err == nil {
		t.Fatal("a zero group child extent has no exact scale and must stay refused")
	}
}
