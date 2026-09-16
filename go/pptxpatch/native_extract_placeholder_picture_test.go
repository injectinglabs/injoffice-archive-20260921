package pptxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// The layout frame is 6096000 x 3429000 EMU (nativePlaceholderFrameGeometry),
// so these paths are authored in the frame's own coordinate space.
const nativePicturePlaceholderCustomGeometry = `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>` +
	`<a:path w="6096000" h="3429000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="6096000" y="0"/></a:lnTo><a:lnTo><a:pt x="6096000" y="3429000"/></a:lnTo><a:lnTo><a:pt x="0" y="3429000"/></a:lnTo><a:close/></a:path>` +
	`<a:path w="6096000" h="3429000"><a:moveTo><a:pt x="1524000" y="857250"/></a:moveTo><a:quadBezTo><a:pt x="3048000" y="1714500"/><a:pt x="4572000" y="857250"/></a:quadBezTo><a:close/></a:path>` +
	`</a:pathLst></a:custGeom>`

const nativePicturePlaceholderFill = `<a:solidFill><a:srgbClr val="FFD966"/></a:solidFill><a:ln><a:noFill/></a:ln>`

// nativePicturePlaceholderFixture mirrors custgeom-placeholder.pptx: a slide
// picture placeholder with an empty p:spPr and no p:txBody at all, whose layout
// placeholder carries the geometry and fill PowerPoint paints for it.
func nativePicturePlaceholderFixture(t *testing.T, slidePlaceholder, layoutPlaceholder, layoutGeometry, layoutPaint string) []byte {
	t.Helper()
	return nativePlaceholderFrameFixture(t, "", "", layoutPlaceholder, layoutPaint, "Picture placeholder prompt", func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<p:ph type="obj" sz="quarter" idx="7"/>`, slidePlaceholder, 1)
		if layoutGeometry != "" {
			parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`, layoutGeometry, 1)
		}
	})
}

func nativePicturePlaceholderRefusal(t *testing.T, deck NativePPTXDeck) string {
	t.Helper()
	if len(deck.Slides[0].Elements) != 0 {
		t.Fatalf("expected the shape to be refused, got %d element(s)", len(deck.Slides[0].Elements))
	}
	for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
		if diagnostic.Code == "pptx.unsupported-shape" {
			return diagnostic.Message
		}
	}
	t.Fatalf("no pptx.unsupported-shape passthrough: %+v", deck.Slides[0].Compatibility.Diagnostics)
	return ""
}

// A picture placeholder inherits a painted frame, not a text chain. Before this
// lane the whole shape was refused, which left custgeom-placeholder.pptx as a
// completely blank page even though the layout declares everything needed.
func TestNativePicturePlaceholderPaintsInheritedCustomGeometry(t *testing.T) {
	input := nativePicturePlaceholderFixture(t,
		`<p:ph type="pic" sz="quarter" idx="7"/>`,
		`<p:ph type="pic" sz="quarter" idx="7"/>`,
		nativePicturePlaceholderCustomGeometry, nativePicturePlaceholderFill)
	before := bytes.Clone(input)
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatalf("picture placeholder frame missing: %+v", deck.Slides[0].Compatibility)
	}
	element := deck.Slides[0].Elements[0]
	if element.Kind != NativeElementKindShape || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("inherited picture frame was not a read-only shape: %+v", element)
	}
	if element.Geometry == nil || element.Preset != nil {
		t.Fatalf("inherited custom geometry was flattened to a preset: preset=%+v geometry=%+v", element.Preset, element.Geometry)
	}
	if len(element.Geometry.Paths) != 2 {
		t.Fatalf("inherited custom geometry lost a path: %+v", element.Geometry.Paths)
	}
	// A picture placeholder has no native v1 placeholder value, so the element
	// must not claim a title/body binding it does not have.
	if element.Placeholder != nil {
		t.Fatalf("frame placeholder claimed a text placeholder binding: %+v", element.Placeholder)
	}
	if element.Fill == nil || *element.Fill != "FFD966" || element.Stroke != nil {
		t.Fatalf("inherited paint did not cascade: fill=%+v stroke=%+v", element.Fill, element.Stroke)
	}
	if *element.Transform.X != 3048000 || *element.Transform.Y != 1714500 || *element.Transform.Cx != 6096000 || *element.Transform.Cy != 3429000 {
		t.Fatalf("layout geometry did not cascade: %+v", element.Transform)
	}
	message := nativePlaceholderPreviewMessage(t, element)
	for _, want := range []string{
		"pic placeholder resolved through the layout-by-index chain as a painted frame",
		"carries no native placeholder binding",
		"geometry a:custGeom evaluated from source (2 path(s))",
		"solid fill FFD966",
		"painted read-only as the inherited source geometry",
	} {
		if !strings.Contains(message, want) {
			t.Fatalf("frame disclosure missing %q: %s", want, message)
		}
	}
	// The text-placeholder caveat is about PowerPoint hiding empty text
	// placeholders; it must not be attached to a frame family.
	if strings.Contains(message, "hide placeholders without text") {
		t.Fatalf("text-placeholder caveat leaked onto a frame placeholder: %s", message)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid native: %+v", issues)
	}
	if !bytes.Equal(input, before) {
		t.Fatal("source mutated")
	}
}

func TestNativePicturePlaceholderPaintsInheritedAdjustedPreset(t *testing.T) {
	input := nativePicturePlaceholderFixture(t,
		`<p:ph type="pic" sz="quarter" idx="7"/>`,
		`<p:ph type="pic" sz="quarter" idx="7"/>`,
		`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom>`,
		nativePicturePlaceholderFill)
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	deck, err := ExtractNativePPTX(input, options)
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides[0].Elements) != 1 {
		t.Fatalf("picture placeholder frame missing: %+v", deck.Slides[0].Compatibility)
	}
	element := deck.Slides[0].Elements[0]
	// An adjusted preset is evaluated from source rather than flattened to the
	// unadjusted rectangle the text families are restricted to.
	if element.Geometry == nil || element.Preset != nil {
		t.Fatalf("adjusted preset was not evaluated: preset=%+v geometry=%+v", element.Preset, element.Geometry)
	}
	if message := nativePlaceholderPreviewMessage(t, element); !strings.Contains(message, "geometry a:prstGeom prst=roundRect evaluated from source") {
		t.Fatalf("adjusted preset was not disclosed by name: %s", message)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid native: %+v", issues)
	}
}

// What cannot be modeled exactly stays omitted with a disclosure that names the
// shape and the construct, never the generic layout message.
func TestNativePicturePlaceholderRefusalsNameTheShapeAndTheConstruct(t *testing.T) {
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	cases := []struct {
		name             string
		slidePlaceholder string
		layoutHolder     string
		layoutGeometry   string
		want             []string
	}{
		{
			name:             "geometry outside the evaluated profile",
			slidePlaceholder: `<p:ph type="pic" sz="quarter" idx="7"/>`,
			layoutHolder:     `<p:ph type="pic" sz="quarter" idx="7"/>`,
			layoutGeometry:   `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst><a:path w="6096000" h="3429000"><a:moveTo><a:pt x="not-a-coordinate" y="0"/></a:moveTo><a:close/></a:path></a:pathLst></a:custGeom>`,
			want: []string{
				`p:sp id=`, `name="Title"`,
				"inherited placeholder frame geometry a:custGeom is outside the evaluated geometry profile",
			},
		},
		{
			name:             "placeholder family with neither text nor frame inheritance",
			slidePlaceholder: `<p:ph type="tbl" sz="quarter" idx="7"/>`,
			layoutHolder:     `<p:ph type="tbl" sz="quarter" idx="7"/>`,
			want: []string{
				`p:sp id=`,
				"placeholder type tbl is outside the read-only title/body inheritance families and inherits no painted frame",
			},
		},
		{
			name:             "slide and layout placeholder types disagree",
			slidePlaceholder: `<p:ph type="pic" sz="quarter" idx="7"/>`,
			layoutHolder:     `<p:ph type="obj" sz="quarter" idx="7"/>`,
			want: []string{
				`p:sp id=`,
				"placeholder type pic conflicts with its layout placeholder type obj",
			},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			input := nativePicturePlaceholderFixture(t, tc.slidePlaceholder, tc.layoutHolder, tc.layoutGeometry, nativePicturePlaceholderFill)
			deck, err := ExtractNativePPTX(input, options)
			if err != nil {
				t.Fatal(err)
			}
			message := nativePicturePlaceholderRefusal(t, deck)
			for _, want := range tc.want {
				if !strings.Contains(message, want) {
					t.Fatalf("refusal missing %q: %s", want, message)
				}
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatalf("invalid native: %+v", issues)
			}
		})
	}
}

// nativePicturePlaceholderXML is the refused picture placeholder appended to a
// slide that also keeps its editable body placeholder.
const nativePicturePlaceholderXML = `<p:sp><p:nvSpPr><p:cNvPr id="9" name="Picture Placeholder 9"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="13"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>`

// The exact tier never sees the approximate frame preview: the picture
// placeholder stays refused there, and a mutation elsewhere on the slide
// rewrites the package with the refused shape's source bytes byte-identical.
func TestNativePicturePlaceholderStaysRefusedAndByteIdenticalOnTheExactTier(t *testing.T) {
	for _, strict := range []bool{false, true} {
		input := nativePlaceholderFixture(t, strict, func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, nativePicturePlaceholderXML+`</p:spTree>`, 1)
			layout := parts["relocated/layouts/layout.xml"]
			layout = strings.Replace(layout, `</p:spTree>`, `<p:sp><p:nvSpPr><p:cNvPr id="8" name="Picture Placeholder 8"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="pic" sz="quarter" idx="13"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="3048000" y="1714500"/><a:ext cx="6096000" cy="3429000"/></a:xfrm>`+nativePicturePlaceholderCustomGeometry+nativePicturePlaceholderFill+`</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp></p:spTree>`, 1)
			parts["relocated/layouts/layout.xml"] = layout
		})
		before := bytes.Clone(input)
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		// The body placeholder still resolves exactly; the picture placeholder
		// is the only refusal, and it is named.
		if len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("exact tier element count changed: %+v", deck.Slides[0].Elements)
		}
		message := nativePicturePlaceholderRefusalMessage(t, deck)
		if !strings.Contains(message, `p:sp id=9 name="Picture Placeholder 9"`) {
			t.Fatalf("exact-tier refusal did not name the shape: %s", message)
		}
		if !bytes.Equal(input, before) {
			t.Fatal("source mutated")
		}
		// Rewriting the editable body placeholder must leave the refused
		// picture placeholder's source bytes byte-identical in the output.
		element := deck.Slides[0].Elements[0]
		paragraphs := nativeMutationParagraphs("Replaced")
		rebuilt, err := ApplyNativePPTXMutations(input, NativePPTXMutationRequest{
			ExpectedSourceRevision: *deck.SourceRevision,
			Operations: []NativePPTXMutation{{
				OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: element.ID,
				ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &paragraphs,
			}},
		})
		if err != nil {
			t.Fatalf("exact-tier round trip failed: %v", err)
		}
		slide := chartZipEntry(t, rebuilt, "relocated/slides/slide-a.xml")
		if !bytes.Contains(slide, []byte(nativePicturePlaceholderXML)) {
			t.Fatalf("exact-tier output did not preserve the refused picture placeholder byte-identically: %s", slide)
		}
		// Everything the mutation did not target stays byte-identical too.
		for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
			if !bytes.Equal(chartZipEntry(t, rebuilt, part), chartZipEntry(t, before, part)) {
				t.Fatalf("exact-tier output rewrote %s", part)
			}
		}
	}
}

func nativePicturePlaceholderRefusalMessage(t *testing.T, deck NativePPTXDeck) string {
	t.Helper()
	for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
		if diagnostic.Code == "pptx.unsupported-shape" {
			return diagnostic.Message
		}
	}
	t.Fatalf("no pptx.unsupported-shape passthrough: %+v", deck.Slides[0].Compatibility.Diagnostics)
	return ""
}
