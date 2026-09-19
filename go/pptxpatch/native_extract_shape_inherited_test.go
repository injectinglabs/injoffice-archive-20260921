package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// nativeInheritedAutoShapeFixture routes a slide placeholder through the
// AutoShape extractor. p:cNvSpPr@txBox="0" is what nativeShapeIsTextBox reads:
// an explicit "not a text box" sends the shape to extractAutoShape directly,
// which is also where a placeholder lands after the text extractor declines it
// (the ShapeLineProperties benchmark deck reaches it that way, through its
// p:style).
func nativeInheritedAutoShapeFixture(t *testing.T, slideProperties string) []byte {
	t.Helper()
	return nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, slideProperties,
		`<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Layout prompt", func(parts map[string]string) {
			slide := parts["relocated/slides/slide-a.xml"]
			parts["relocated/slides/slide-a.xml"] = strings.Replace(slide, `<p:cNvSpPr/>`, `<p:cNvSpPr txBox="0"/>`, 1)
		})
}

func TestNativeAutoShapePlaceholderInheritsItsFrame(t *testing.T) {
	local := `<a:xfrm><a:off x="100000" y="200000"/><a:ext cx="3000000" cy="1000000"/></a:xfrm>`
	for _, tc := range []struct {
		name               string
		slideProperties    string
		x, y, cx, cy       int64
		expectInheritedFit bool
	}{
		{"slide declares its own frame", local, 100000, 200000, 3000000, 1000000, false},
		{"slide declares nothing at all", "", 3048000, 1714500, 6096000, 3429000, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			input := nativeInheritedAutoShapeFixture(t, tc.slideProperties)
			before := bytes.Clone(input)
			options := nativeTestExtractOptions()
			options.AllowInheritedTextPreview = true
			deck, err := ExtractNativePPTX(input, options)
			if err != nil {
				t.Fatal(err)
			}
			if len(deck.Slides[0].Elements) != 1 {
				t.Fatalf("placeholder shape missing: %+v", deck.Slides[0].Compatibility)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("inherited placeholder frame was not painted: %+v", element.Compatibility)
			}
			if element.Preset == nil || *element.Preset != NativeShapePresetRect {
				t.Fatalf("inherited geometry was not the layout rectangle: %+v", element.Preset)
			}
			if element.Fill == nil || *element.Fill != "FBE4D5" {
				t.Fatalf("inherited fill was not painted: %+v", element.Fill)
			}
			if element.Stroke == nil || element.Stroke.Color != "C55A11" {
				t.Fatalf("inherited outline was not painted: %+v", element.Stroke)
			}
			if *element.Transform.X != tc.x || *element.Transform.Y != tc.y || *element.Transform.Cx != tc.cx || *element.Transform.Cy != tc.cy {
				t.Fatalf("frame precedence failed: %+v", element.Transform)
			}
			disclosed := false
			for _, diagnostic := range element.Compatibility.Diagnostics {
				if diagnostic.Code == nativeInheritedShapeFrameCode {
					disclosed = true
				}
				if diagnostic.Severity == NativeDiagnosticSeverityRefusal {
					t.Fatalf("inherited frame still refuses: %+v", diagnostic)
				}
			}
			if !disclosed {
				t.Fatalf("inherited frame was painted without disclosure: %+v", element.Compatibility.Diagnostics)
			}
			if !bytes.Equal(input, before) {
				t.Fatal("source mutated")
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatalf("invalid native: %+v", issues)
			}
		})
	}
}

// The exact tier claims to model only what a slide states itself, so nothing
// here may leak out of the opt-in preview.
func TestNativeAutoShapePlaceholderInheritanceIsPreviewOnly(t *testing.T) {
	input := nativeInheritedAutoShapeFixture(t, `<a:xfrm><a:off x="100000" y="200000"/><a:ext cx="3000000" cy="1000000"/></a:xfrm>`)
	deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatalf("expected the shape to survive as a refusal: %+v", deck.Slides[0].Compatibility)
	}
	element := deck.Slides[0].Elements[0]
	if element.Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatalf("the exact tier inherited a frame it does not model: %+v", element)
	}
	if element.Fill != nil || element.Stroke != nil {
		t.Fatalf("the exact tier painted inherited paint: fill=%+v stroke=%+v", element.Fill, element.Stroke)
	}
}

// A placeholder family the preview does not resolve keeps its refusal: the
// AutoShape path must not invent a rectangle for a chain it never qualified.
func TestNativeAutoShapeUnqualifiedPlaceholderStillRefuses(t *testing.T) {
	input := nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, `<a:xfrm><a:off x="100000" y="200000"/><a:ext cx="3000000" cy="1000000"/></a:xfrm>`,
		`<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Layout prompt", func(parts map[string]string) {
			slide := parts["relocated/slides/slide-a.xml"]
			slide = strings.Replace(slide, `<p:cNvSpPr/>`, `<p:cNvSpPr txBox="0"/>`, 1)
			parts["relocated/slides/slide-a.xml"] = slide
			layout := parts["relocated/layouts/layout.xml"]
			parts["relocated/layouts/layout.xml"] = strings.Replace(layout, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`, "", 1)
		})
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 || deck.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused {
		t.Fatalf("a chain that declares no geometry must keep refusing: %+v", deck.Slides[0].Elements)
	}
}
