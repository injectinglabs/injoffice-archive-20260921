package pptxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

const nativePlaceholderFrameGeometry = `<a:xfrm><a:off x="3048000" y="1714500"/><a:ext cx="6096000" cy="3429000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`
const nativePlaceholderFramePaint = `<a:solidFill><a:srgbClr val="FBE4D5"/></a:solidFill><a:ln><a:solidFill><a:srgbClr val="C55A11"/></a:solidFill></a:ln>`

const nativePlaceholderFrameEmptyBody = `<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>`
const nativePlaceholderFrameTextBody = `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Real content, typed by the author</a:t></a:r></a:p></p:txBody>`

// nativePlaceholderFrameFixture mirrors the content-placeholder benchmark: an
// obj placeholder with an empty p:spPr on the slide whose layout placeholder
// carries geometry, a peach solid fill, an orange outline and prompt text.
func nativePlaceholderFrameFixture(t *testing.T, slideBody, slideProperties, layoutPlaceholder, layoutPaint, layoutPrompt string, customize func(map[string]string)) []byte {
	t.Helper()
	return nativePlaceholderFixture(t, false, func(parts map[string]string) {
		slide := parts["relocated/slides/slide-a.xml"]
		slide = strings.Replace(slide, `<p:ph idx="7"/>`, `<p:ph type="obj" sz="quarter" idx="7"/>`, 1)
		start, end := strings.Index(slide, "<p:txBody>"), strings.Index(slide, "</p:txBody>")+len("</p:txBody>")
		slide = slide[:start] + slideBody + slide[end:]
		if slideProperties != "" {
			slide = strings.Replace(slide, `<p:spPr/>`, `<p:spPr>`+slideProperties+`</p:spPr>`, 1)
		}
		parts["relocated/slides/slide-a.xml"] = slide
		layout := parts["relocated/layouts/layout.xml"]
		layout = strings.Replace(layout, `<p:ph type="body" idx="7"/>`, layoutPlaceholder, 1)
		layout = strings.Replace(layout, `<p:spPr></p:spPr>`, `<p:spPr>`+nativePlaceholderFrameGeometry+layoutPaint+`</p:spPr>`, 1)
		layout = strings.Replace(layout, `<a:p/>`, `<a:p><a:r><a:rPr lang="en-US"/><a:t>`+layoutPrompt+`</a:t></a:r></a:p>`, 1)
		parts["relocated/layouts/layout.xml"] = layout
		if customize != nil {
			customize(parts)
		}
	})
}

func nativePlaceholderPreviewMessage(t *testing.T, element NativeElement) string {
	t.Helper()
	for _, diagnostic := range element.Compatibility.Diagnostics {
		if diagnostic.Code == nativePlaceholderPreviewCode {
			return diagnostic.Message
		}
	}
	t.Fatalf("placeholder preview disclosure missing: %+v", element.Compatibility.Diagnostics)
	return ""
}

func TestNativePlaceholderInheritedFrameIsPaintedInPreviewLane(t *testing.T) {
	cases := []struct {
		name              string
		slideBody         string
		layoutPlaceholder string
		layoutPrompt      string
		hasText           bool
	}{
		{"no text body", "", `<p:ph type="obj" sz="quarter" idx="7"/>`, "Layout text, no custom prompt", false},
		{"empty text body", nativePlaceholderFrameEmptyBody, `<p:ph type="obj" sz="quarter" idx="7"/>`, "Layout text, no custom prompt", false},
		{"real text", nativePlaceholderFrameTextBody, `<p:ph type="obj" sz="quarter" idx="7"/>`, "Layout text, no custom prompt", true},
		{"no text body, custom prompt", "", `<p:ph type="obj" sz="quarter" idx="7" hasCustomPrompt="1"/>`, "Custom prompt to insert content", false},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			input := nativePlaceholderFrameFixture(t, tc.slideBody, "", tc.layoutPlaceholder, nativePlaceholderFramePaint, tc.layoutPrompt, nil)
			before := bytes.Clone(input)
			options := nativeTestExtractOptions()
			options.AllowInheritedTextPreview = true
			deck, err := ExtractNativePPTX(input, options)
			if err != nil {
				t.Fatal(err)
			}
			if len(deck.Slides[0].Elements) != 1 {
				t.Fatalf("placeholder preview missing: %+v", deck.Slides[0].Compatibility)
			}
			element := deck.Slides[0].Elements[0]
			if element.Kind != NativeElementKindShape || element.Preset == nil || *element.Preset != NativeShapePresetRect || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("inherited frame was not projected as a read-only rect shape: %+v", element)
			}
			if element.Placeholder == nil || *element.Placeholder != NativePlaceholderTypeBody || element.Paragraphs == nil || element.TextBody == nil {
				t.Fatalf("placeholder projection lost its family, paragraphs or body layout: %+v", element)
			}
			if *element.Transform.X != 3048000 || *element.Transform.Y != 1714500 || *element.Transform.Cx != 6096000 || *element.Transform.Cy != 3429000 {
				t.Fatalf("layout geometry did not cascade: %+v", element.Transform)
			}
			if element.Fill == nil || *element.Fill != "FBE4D5" {
				t.Fatalf("inherited fill was not painted: %+v", element.Fill)
			}
			if element.Stroke == nil || element.Stroke.Color != "C55A11" || element.Stroke.WidthEMU == nil || *element.Stroke.WidthEMU != nativePlaceholderPreviewHairlineWidthEmu || element.Stroke.Cap != nil || element.Stroke.Join != nil || element.Stroke.Dash != nil {
				t.Fatalf("inherited outline was not painted as a hairline: %+v", element.Stroke)
			}
			hasText := false
			for _, paragraph := range *element.Paragraphs {
				for _, run := range paragraph.Runs {
					if run.Text != nil && strings.Contains(*run.Text, tc.layoutPrompt) {
						t.Fatal("layout prompt leaked into slide content")
					}
					hasText = hasText || (run.Text != nil && strings.TrimSpace(*run.Text) != "")
				}
			}
			if hasText != tc.hasText {
				t.Fatalf("slide text presence mismatch: %+v", *element.Paragraphs)
			}
			message := nativePlaceholderPreviewMessage(t, element)
			if !strings.Contains(message, "Frame solid fill FBE4D5 and solid outline C55A11 (9525 EMU, a hairline default because the source declares no width) inherited from the slide/layout/master placeholder chain") || !strings.Contains(message, "ancestor prompt paragraphs") {
				t.Fatalf("inherited frame was not disclosed: %s", message)
			}
			if strings.Contains(message, "hide placeholders without text") == tc.hasText {
				t.Fatalf("empty-placeholder caveat mismatch: %s", message)
			}
			if strings.Contains(message, "no text body") == (tc.slideBody != "") {
				t.Fatalf("no-text-body disclosure mismatch: %s", message)
			}
			if issues := ValidateNativePPTX(deck); len(issues) > 0 {
				t.Fatalf("invalid native: %+v", issues)
			}
			if !bytes.Equal(input, before) {
				t.Fatal("source mutated")
			}
			replacement := nativeMutationParagraphs("changed")
			if _, err := resolveNativePPTXMutations(deck, []NativePPTXMutation{{OperationID: "edit", Kind: NativePPTXReplaceText, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, Paragraphs: &replacement}}); err == nil {
				t.Fatal("placeholder frame preview authorized text mutation")
			}
			fill := "112233"
			if _, err := resolveNativePPTXMutations(deck, []NativePPTXMutation{{OperationID: "paint", Kind: NativePPTXUpdateAutoShape, ElementID: element.ID, ExpectedFingerprintSHA256: element.Source.FingerprintSHA256, AutoShape: &NativePPTXAutoShapeMutation{Transform: element.Transform, Preset: NativeShapePresetRect, Fill: &fill}}}); err == nil {
				t.Fatal("placeholder frame preview authorized AutoShape mutation")
			}
		})
	}
}

func TestNativePlaceholderInheritedFrameStaysOutsideExactAndSourceFrameTiers(t *testing.T) {
	painted := nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, "", `<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Layout text, no custom prompt", nil)
	unpainted := nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, "", `<p:ph type="obj" sz="quarter" idx="7"/>`, "", "Layout text, no custom prompt", nil)
	for _, tier := range []struct {
		name    string
		options func() NativePPTXExtractOptions
	}{
		{"exact", nativeTestExtractOptions},
		{"source-frame", func() NativePPTXExtractOptions {
			options := nativeTestExtractOptions()
			options.AllowSourceFrameAutoFitPreview = true
			return options
		}},
	} {
		paintedDeck, err := ExtractNativePPTX(painted, tier.options())
		if err != nil {
			t.Fatal(err)
		}
		unpaintedDeck, err := ExtractNativePPTX(unpainted, tier.options())
		if err != nil {
			t.Fatal(err)
		}
		for _, element := range paintedDeck.Slides[0].Elements {
			if element.Compatibility.Status != NativeCompatibilityStatusRefused || element.Fill != nil || element.Stroke != nil || element.Kind == NativeElementKindShape {
				t.Fatalf("%s tier painted an inherited placeholder frame: %+v", tier.name, element)
			}
		}
		// Layout paint must not change the exact projection at all: the two
		// decks differ only by the ignored layout markup.
		paintedJSON, _ := json.Marshal(paintedDeck)
		unpaintedJSON, _ := json.Marshal(unpaintedDeck)
		if !bytes.Equal(nativeStripPackageIdentity(paintedJSON), nativeStripPackageIdentity(unpaintedJSON)) {
			t.Fatalf("%s tier output changed with layout paint present:\n%s\n%s", tier.name, paintedJSON, unpaintedJSON)
		}
	}
}

// nativeStripPackageIdentity blanks the digests and revision that legitimately
// differ between two packages with different bytes.
func nativeStripPackageIdentity(data []byte) []byte {
	var generic map[string]any
	if err := json.Unmarshal(data, &generic); err != nil {
		return data
	}
	var strip func(value any) any
	strip = func(value any) any {
		switch typed := value.(type) {
		case map[string]any:
			for key, item := range typed {
				if strings.Contains(strings.ToLower(key), "sha256") || key == "sourceRevision" || key == "revision" || key == "token" || key == "fingerprintSha256" {
					typed[key] = ""
					continue
				}
				typed[key] = strip(item)
			}
			return typed
		case []any:
			for i := range typed {
				typed[i] = strip(typed[i])
			}
			return typed
		}
		return value
	}
	stripped, _ := json.Marshal(strip(generic))
	return stripped
}

func TestNativePlaceholderInheritedFrameNearestLayerWins(t *testing.T) {
	options := nativeTestExtractOptions()
	options.AllowInheritedTextPreview = true
	extract := func(t *testing.T, input []byte) NativeElement {
		t.Helper()
		deck, err := ExtractNativePPTX(input, options)
		if err != nil {
			t.Fatal(err)
		}
		if len(deck.Slides[0].Elements) != 1 {
			t.Fatalf("placeholder preview missing: %+v", deck.Slides[0].Compatibility)
		}
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("invalid native: %+v", issues)
		}
		return deck.Slides[0].Elements[0]
	}
	t.Run("slide noFill wins over the layout fill", func(t *testing.T) {
		element := extract(t, nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, `<a:noFill/>`, `<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Prompt", nil))
		if element.Kind != NativeElementKindShape || element.Fill != nil || element.Stroke == nil || element.Stroke.Color != "C55A11" {
			t.Fatalf("slide a:noFill did not win: %+v", element)
		}
		if message := nativePlaceholderPreviewMessage(t, element); strings.Contains(message, "solid fill") || !strings.Contains(message, "solid outline C55A11") {
			t.Fatalf("disclosure did not follow the resolved paint: %s", message)
		}
	})
	t.Run("slide outline noFill wins over the layout outline", func(t *testing.T) {
		element := extract(t, nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, `<a:ln><a:noFill/></a:ln>`, `<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Prompt", nil))
		if element.Kind != NativeElementKindShape || element.Stroke != nil || element.Fill == nil || *element.Fill != "FBE4D5" {
			t.Fatalf("slide a:ln/a:noFill did not win: %+v", element)
		}
	})
	t.Run("no resolved paint keeps the text projection", func(t *testing.T) {
		element := extract(t, nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, `<a:noFill/><a:ln><a:noFill/></a:ln>`, `<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Prompt", nil))
		if element.Kind != NativeElementKindText || element.Fill != nil || element.Stroke != nil || element.Preset != nil {
			t.Fatalf("unpainted placeholder changed kind: %+v", element)
		}
		if message := nativePlaceholderPreviewMessage(t, element); strings.Contains(message, "Frame ") {
			t.Fatalf("unpainted frame was disclosed as painted: %s", message)
		}
	})
	t.Run("unmodeled nearer paint clears the farther solid paint", func(t *testing.T) {
		gradient := `<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst></a:gradFill><a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="C55A11"/></a:solidFill><a:prstDash val="dash"/></a:ln>`
		element := extract(t, nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, "", `<p:ph type="obj" sz="quarter" idx="7"/>`, gradient, "Prompt", func(parts map[string]string) {
			master := parts["relocated/masters/master.xml"]
			master = strings.Replace(master, `<a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr>`, `<a:ext cx="4572000" cy="914400"/></a:xfrm><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="12700"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln></p:spPr>`, 1)
			parts["relocated/masters/master.xml"] = master
		}))
		if element.Kind != NativeElementKindText || element.Fill != nil || element.Stroke != nil {
			t.Fatalf("farther master paint leaked through an unmodeled layout paint: %+v", element)
		}
		if message := nativePlaceholderPreviewMessage(t, element); !strings.Contains(message, "ancestor a:gradFill") || !strings.Contains(message, "ancestor a:ln (dash)") {
			t.Fatalf("unmodeled paint was not disclosed: %s", message)
		}
	})
	t.Run("declared width and theme colors resolve", func(t *testing.T) {
		paint := `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln w="38100"><a:solidFill><a:srgbClr val="C55A11"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
		element := extract(t, nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, "", `<p:ph type="obj" sz="quarter" idx="7"/>`, paint, "Prompt", nil))
		if element.Kind != NativeElementKindShape || element.Fill == nil || len(*element.Fill) != 6 || element.Stroke == nil || *element.Stroke.WidthEMU != 38100 {
			t.Fatalf("theme fill or declared outline width did not resolve: %+v %+v", element.Fill, element.Stroke)
		}
		if message := nativePlaceholderPreviewMessage(t, element); strings.Contains(message, "hairline") || !strings.Contains(message, "ancestor a:ln/a:round") {
			t.Fatalf("declared width disclosure mismatch: %s", message)
		}
	})
	t.Run("slide-level unmodeled paint still refuses", func(t *testing.T) {
		for _, properties := range []string{`<a:gradFill/>`, `<a:ln><a:solidFill><a:srgbClr val="C55A11"/></a:solidFill><a:prstDash val="dash"/></a:ln>`, `<a:solidFill><a:srgbClr val="FBE4D5"><a:alpha val="50000"/></a:srgbClr></a:solidFill>`} {
			deck, err := ExtractNativePPTX(nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, properties, `<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Prompt", nil), options)
			if err != nil {
				t.Fatal(err)
			}
			for _, element := range deck.Slides[0].Elements {
				if element.Compatibility.Status != NativeCompatibilityStatusRefused {
					t.Fatalf("slide-level %s was painted: %+v", properties, element)
				}
			}
		}
	})
	t.Run("slide-level solid paint resolves like an ancestor", func(t *testing.T) {
		element := extract(t, nativePlaceholderFrameFixture(t, nativePlaceholderFrameTextBody, `<a:solidFill><a:srgbClr val="00FF00"/></a:solidFill>`, `<p:ph type="obj" sz="quarter" idx="7"/>`, nativePlaceholderFramePaint, "Prompt", nil))
		if element.Kind != NativeElementKindShape || element.Fill == nil || *element.Fill != "00FF00" || element.Stroke == nil || element.Stroke.Color != "C55A11" {
			t.Fatalf("slide-level solid fill did not win: %+v", element)
		}
	})
}
