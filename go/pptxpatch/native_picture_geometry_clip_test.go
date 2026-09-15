package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func nativePictureDiagnosticCodes(picture NativeElement) map[string]bool {
	codes := map[string]bool{}
	for _, diagnostic := range picture.Compatibility.Diagnostics {
		codes[diagnostic.Code] = true
	}
	return codes
}

func TestNativePictureNonRectangularPresetIsEvaluatedAsReadOnlyClip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			name     string
			geometry string
		}{
			{"ellipse", `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`},
			{"adjusted roundRect", `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom>`},
			{"heart", `<a:prstGeom prst="heart"><a:avLst/></a:prstGeom>`},
		} {
			data := nativePictureFixture(t, nativePictureFixtureOptions{strict: strict, pictureGeometry: tc.geometry, sourceRect: `<a:srcRect l="10000" t="20000" r="30000" b="0"/>`})
			before := bytes.Clone(data)
			deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("%s: %v", tc.name, err)
			}
			picture := nativeFixturePicture(t, deck.Slides[0])
			if picture.Geometry == nil || picture.Geometry.Profile != "drawingml-paths-v1" || len(picture.Geometry.Paths) == 0 {
				t.Fatalf("%s: missing evaluated picture geometry: %#v", tc.name, picture.Geometry)
			}
			if picture.Clip != nil {
				t.Fatalf("%s: evaluated geometry must not also claim the exact roundRect clip", tc.name)
			}
			if picture.Crop == nil || *picture.Crop.Left != 10000 {
				t.Fatalf("%s: lost exact source crop", tc.name)
			}
			if picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(picture.Passthrough) != 1 {
				t.Fatalf("%s: evaluated picture geometry must stay preserve-only with a passthrough capability", tc.name)
			}
			codes := nativePictureDiagnosticCodes(picture)
			if !codes[nativePictureGeometryPreviewCode] || codes["pptx.picture-geometry-unavailable"] {
				t.Fatalf("%s: unexpected diagnostics %v", tc.name, codes)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("%s: contract issues %v", tc.name, issues)
			}
			if !bytes.Equal(data, before) {
				t.Fatal("source changed")
			}
		}
	}
}

func TestNativePictureEllipseClipMatchesAutoShapeCatalogEvaluation(t *testing.T) {
	data := nativePictureFixture(t, nativePictureFixtureOptions{pictureGeometry: `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`})
	deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	expected, err := EvaluateNativePPTXPresetGeometry("ellipse", *picture.Transform.Cx, *picture.Transform.Cy, nil)
	if err != nil {
		t.Fatal(err)
	}
	if len(picture.Geometry.Paths) != len(expected.Paths) || picture.Geometry.TextRect != expected.TextRect {
		t.Fatalf("picture geometry %#v differs from catalog evaluation %#v", picture.Geometry, expected)
	}
	for i := range expected.Paths {
		if len(picture.Geometry.Paths[i].Commands) != len(expected.Paths[i].Commands) || picture.Geometry.Paths[i].FillMode != expected.Paths[i].FillMode {
			t.Fatalf("path %d differs from catalog evaluation", i)
		}
	}
	arcs := 0
	for _, command := range picture.Geometry.Paths[0].Commands {
		if command.Kind == "arcTo" {
			arcs++
		}
	}
	if arcs == 0 {
		t.Fatal("ellipse clip should be an arc outline")
	}
}

func TestNativePictureGeometryOutsideCatalogStaysUnavailable(t *testing.T) {
	for _, tc := range []struct {
		name     string
		geometry string
	}{
		{"unknown preset", `<a:prstGeom prst="notAShape"><a:avLst/></a:prstGeom>`},
		{"undeclared adjustment", `<a:prstGeom prst="ellipse"><a:avLst><a:gd name="adj" fmla="val 1"/></a:avLst></a:prstGeom>`},
		{"non-literal adjustment", `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="*/ w 1 2"/></a:avLst></a:prstGeom>`},
		{"foreign attribute", `<a:prstGeom prst="ellipse" extra="1"><a:avLst/></a:prstGeom>`},
		{"oversized preset name", `<a:prstGeom prst="` + strings.Repeat("x", 4000) + `"><a:avLst/></a:prstGeom>`},
	} {
		data := nativePictureFixture(t, nativePictureFixtureOptions{pictureGeometry: tc.geometry})
		before := bytes.Clone(data)
		deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("%s: %v", tc.name, err)
		}
		picture := nativeFixturePicture(t, deck.Slides[0])
		if picture.Geometry != nil || picture.Clip != nil {
			t.Fatalf("%s: refused geometry must not be evaluated: %#v", tc.name, picture.Geometry)
		}
		codes := nativePictureDiagnosticCodes(picture)
		if !codes["pptx.picture-geometry-unavailable"] || codes[nativePictureGeometryPreviewCode] || picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatalf("%s: unexpected diagnostics %v status %s", tc.name, codes, picture.Compatibility.Status)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("%s: diagnostics must stay within the contract: %v", tc.name, issues)
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source changed")
		}
	}
}

func TestNativePictureGeometryContractStaysReadOnlyAndExclusive(t *testing.T) {
	data := nativePictureFixture(t, nativePictureFixtureOptions{pictureGeometry: `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`})
	deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	for _, tc := range []struct {
		name   string
		mutate func(*NativeElement)
		code   string
	}{
		{"editable geometry", func(e *NativeElement) { e.Compatibility.Status = NativeCompatibilityStatusEditable }, "native.geometryAuthority"},
		{"clip and geometry", func(e *NativeElement) { e.Clip = stringPointer("roundRect") }, "native.geometry"},
		{"empty paths", func(e *NativeElement) { e.Geometry.Paths = nil }, "native.geometryBudget"},
	} {
		element := picture
		geometry := *picture.Geometry
		element.Geometry = &geometry
		element.Compatibility.Diagnostics = append([]NativeDiagnostic(nil), picture.Compatibility.Diagnostics...)
		tc.mutate(&element)
		copyDeck := deck
		copyDeck.Slides = []NativeSlide{deck.Slides[0]}
		copyDeck.Slides[0].Elements = []NativeElement{element}
		issues := ValidateNativePPTX(copyDeck)
		found := false
		for _, issue := range issues {
			if issue.Code == tc.code {
				found = true
			}
		}
		if !found {
			t.Fatalf("%s: expected %s in %v", tc.name, tc.code, issues)
		}
	}
	shape := NativeElement{Kind: NativeElementKindText, ID: "text-geometry", Provenance: NativeProvenanceAuthored, Transform: picture.Transform, Paragraphs: &[]NativeParagraph{}, Geometry: picture.Geometry, Passthrough: []NativePassthroughRef{}, Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusPreserveOnly, Diagnostics: []NativeDiagnostic{}}}
	copyDeck := deck
	copyDeck.Slides = []NativeSlide{deck.Slides[0]}
	copyDeck.Slides[0].Elements = []NativeElement{shape}
	rejected := false
	for _, issue := range ValidateNativePPTX(copyDeck) {
		if issue.Code == "native.elementUnion" && strings.HasSuffix(issue.Path, ".geometry") {
			rejected = true
		}
	}
	if !rejected {
		t.Fatal("text elements must still reject evaluated geometry")
	}
}

func TestNativePictureRotatedOrFlippedPresetStaysPlaceholder(t *testing.T) {
	for _, attrs := range []string{` rot="5400000"`, ` flipH="1"`, ` rot="5400000" flipH="1" flipV="1"`} {
		data := nativePictureFixture(t, nativePictureFixtureOptions{xfrmAttrs: attrs, pictureGeometry: `<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom>`})
		before := bytes.Clone(data)
		deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("%s: %v", attrs, err)
		}
		picture := nativeFixturePicture(t, deck.Slides[0])
		if picture.Geometry != nil || picture.Clip != nil {
			t.Fatalf("%s: rotated or flipped pictures must not carry an evaluated outline: %#v", attrs, picture.Geometry)
		}
		codes := nativePictureDiagnosticCodes(picture)
		if !codes["pptx.picture-transform-unavailable"] || !codes["pptx.picture-geometry-unavailable"] || codes[nativePictureGeometryPreviewCode] || picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatalf("%s: unexpected diagnostics %v status %s", attrs, codes, picture.Compatibility.Status)
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source changed")
		}
	}
	// Exact rect / roundRect keep their pre-existing rotated behavior: transform gap only.
	data := nativePictureFixture(t, nativePictureFixtureOptions{xfrmAttrs: ` rot="5400000"`, pictureGeometry: `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>`})
	deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	codes := nativePictureDiagnosticCodes(picture)
	if picture.Clip == nil || codes["pptx.picture-geometry-unavailable"] || !codes["pptx.picture-transform-unavailable"] {
		t.Fatalf("rotated roundRect contract changed: clip=%v codes=%v", picture.Clip, codes)
	}
}

func TestNativePictureGeometryTreatsOmittedAdjustmentListAsDefaults(t *testing.T) {
	for _, tc := range []struct {
		geometry string
		clip     bool
		outline  bool
	}{
		{`<a:prstGeom prst="rect"/>`, false, false},
		{`<a:prstGeom prst="roundRect"/>`, true, false},
		{`<a:prstGeom prst="ellipse"/>`, false, true},
	} {
		data := nativePictureFixture(t, nativePictureFixtureOptions{pictureGeometry: tc.geometry})
		deck, err := ExtractNativePPTX(data, nativeTestExtractOptions())
		if err != nil {
			t.Fatalf("%s: %v", tc.geometry, err)
		}
		picture := nativeFixturePicture(t, deck.Slides[0])
		codes := nativePictureDiagnosticCodes(picture)
		if (picture.Clip != nil) != tc.clip || (picture.Geometry != nil) != tc.outline || codes["pptx.picture-geometry-unavailable"] || codes[nativePictureGeometryPreviewCode] != tc.outline {
			t.Fatalf("%s: clip=%v geometry=%v codes=%v", tc.geometry, picture.Clip, picture.Geometry != nil, codes)
		}
		if tc.outline {
			withList, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{pictureGeometry: `<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>`}), nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if len(nativeFixturePicture(t, withList.Slides[0]).Geometry.Paths[0].Commands) != len(picture.Geometry.Paths[0].Commands) {
				t.Fatal("omitted avLst must evaluate exactly like an empty one")
			}
		}
	}
}
