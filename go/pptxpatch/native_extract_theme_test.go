package pptxpatch

import (
	"encoding/xml"
	"strings"
	"testing"
)

func TestExtractNativePPTXResolvesThemeFontsAndUntransformedSchemeColors(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			drawingNS := nsDrawingTransitional
			presentationNS := nsPresentationTransitional
			if strict {
				drawingNS = nsDrawingStrict
				presentationNS = nsPresentationStrict
			}
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{
				strict: strict,
				mutate: func(parts map[string]string) {
					parts["relocated/themes/theme.xml"] = nativeExactThemeXML(drawingNS)
					parts["relocated/masters/master.xml"] = nativeExactMasterWithColorMapXML(presentationNS)
					slide := parts["relocated/slides/slide-a.xml"]
					slide = strings.Replace(slide, `typeface="Aptos"`, `typeface="+mj-lt"`, 1)
					slide = strings.Replace(slide, `typeface="Aptos"`, `typeface="+mn-lt"`, 1)
					slide = strings.Replace(slide, `<a:srgbClr val="112233"/>`, `<a:schemeClr val="tx1"/>`, 1)
					slide = strings.Replace(slide, `<a:srgbClr val="112233"/>`, `<a:schemeClr val="accent1"/>`, 1)
					parts["relocated/slides/slide-a.xml"] = slide
				},
			}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract theme-resolved text: %v", err)
			}
			element := deck.Slides[0].Elements[0]
			if element.Kind != NativeElementKindText || element.Compatibility.Status != NativeCompatibilityStatusEditable || element.Paragraphs == nil || len(*element.Paragraphs) != 1 || len((*element.Paragraphs)[0].Runs) != 2 {
				t.Fatalf("theme-resolved text was not exact/editable: %#v", element)
			}
			first, second := (*element.Paragraphs)[0].Runs[0], (*element.Paragraphs)[0].Runs[1]
			if first.FontFamily == nil || *first.FontFamily != "Calibri Light" || first.Color == nil || *first.Color != "000000" {
				t.Fatalf("major latin / tx1 were not materialized: %#v", first)
			}
			if second.FontFamily == nil || *second.FontFamily != "Calibri" || second.Color == nil || *second.Color != "2F6FED" {
				t.Fatalf("minor latin / accent1 were not materialized: %#v", second)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("theme-resolved deck invalid: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXThemeFontAndColorGapsRemainObjectRefusals(t *testing.T) {
	t.Parallel()
	cases := []struct {
		name   string
		mutate func(map[string]string)
	}{
		{
			name: "unknown-font-token",
			mutate: func(parts map[string]string) {
				parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `typeface="Aptos"`, `typeface="+mj-xx"`, 1)
			},
		},
		{
			name: "missing-font-scheme",
			mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `typeface="Aptos"`, `typeface="+mj-lt"`, 1)
			},
		},
		{
			name: "alpha-only-scheme-color",
			mutate: func(parts map[string]string) {
				parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
				parts["relocated/masters/master.xml"] = nativeExactMasterWithColorMapXML(nsPresentationTransitional)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:srgbClr val="112233"/>`, `<a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr>`, 1)
			},
		},
		{
			name: "unmodeled-hue-mod",
			mutate: func(parts map[string]string) {
				parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
				parts["relocated/masters/master.xml"] = nativeExactMasterWithColorMapXML(nsPresentationTransitional)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:srgbClr val="112233"/>`, `<a:schemeClr val="accent1"><a:hueMod val="50000"/></a:schemeClr>`, 1)
			},
		},
		{
			name: "placeholder-scheme-color",
			mutate: func(parts map[string]string) {
				parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:srgbClr val="112233"/>`, `<a:schemeClr val="phClr"/>`, 1)
			},
		},
	}
	for _, test := range cases {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: test.mutate}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract refused theme gap: %v", err)
			}
			element := deck.Slides[0].Elements[0]
			if element.Compatibility.Status != NativeCompatibilityStatusRefused || len(element.Passthrough) == 0 {
				t.Fatalf("unresolved theme token was approximated: %#v", element)
			}
		})
	}
}

func TestExtractNativePPTXResolvesThemePaintOnAutoShapeAndConnector(t *testing.T) {
	t.Parallel()
	fill := `<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>`
	line := `<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="dk2"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
	shape := nativeAutoShapeXML(3, "Themed shape", "rect", fill, line, "")
	connector := nativeConnectorXML(4, "Themed connector", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", "", "", "")
	deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
		parts["relocated/masters/master.xml"] = nativeExactMasterWithColorMapXML(nsPresentationTransitional)
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, shape+connector+`</p:spTree>`, 1)
	}}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract theme paint: %v", err)
	}
	autoShapes := nativeFixtureAutoShapes(deck.Slides[0])
	connectors := nativeFixtureConnectors(deck.Slides[0])
	if len(autoShapes) != 1 || autoShapes[0].Compatibility.Status != NativeCompatibilityStatusEditable || autoShapes[0].Fill == nil || *autoShapes[0].Fill != "2F6FED" || autoShapes[0].Stroke == nil || autoShapes[0].Stroke.Color != "1D2427" {
		t.Fatalf("autoshape theme paint was not exact: %#v", autoShapes)
	}
	if len(connectors) != 1 || connectors[0].Compatibility.Status != NativeCompatibilityStatusEditable || connectors[0].Stroke == nil || connectors[0].Stroke.Color != "1D2427" {
		t.Fatalf("connector theme paint was not exact: %#v", connectors)
	}
}

func nativeExactThemeXML(drawingNS string) string {
	return `<a:theme xmlns:a="` + drawingNS + `" name="InjOffice">` +
		`<a:themeElements><a:clrScheme name="InjOffice">` +
		`<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
		`<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="1D2427"/></a:dk2>` +
		`<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="2F6FED"/></a:accent1>` +
		`<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
		`<a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="4472C4"/></a:accent5>` +
		`<a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
		`<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme>` +
		`<a:fontScheme name="InjOffice">` +
		`<a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/></a:majorFont>` +
		`<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>` +
		`</a:fontScheme>` +
		`<a:fmtScheme name="InjOffice"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>` +
		`</a:themeElements></a:theme>`
}

func nativeExactMasterWithColorMapXML(presentationNS string) string {
	return `<p:sldMaster xmlns:p="` + presentationNS + `"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>` +
		`<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
		`</p:sldMaster>`
}

func TestNativeSchemeColorTransformsMatchDocumentedOfficeSRGB(t *testing.T) {
	t.Parallel()
	// ECMA-376 tint: 40% of 2F6FED mixed with 60% white.
	if got := applyNativeTint(mustParseNativeSRGB(t, "2F6FED"), 40000).hex(); got != "ACC5F8" {
		t.Fatalf("tint 40000 of 2F6FED = %s, want ACC5F8", got)
	}
	if got := applyNativeShade(mustParseNativeSRGB(t, "2F6FED"), 50000).hex(); got != "183877" {
		t.Fatalf("shade 50000 of 2F6FED = %s, want 183877", got)
	}
	// Published DrawingML lumMod/lumOff vector: Accent 1 40% lighter.
	got, err := applyNativeSchemeColorTransforms("5B9BD5", &nativeXMLNode{Children: []*nativeXMLNode{
		{Name: xml.Name{Local: "lumMod"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "60000"}}},
		{Name: xml.Name{Local: "lumOff"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "40000"}}},
	}}, nativeExtractDialect{drawing: nsDrawingTransitional})
	if err != nil {
		t.Fatalf("lumMod/lumOff: %v", err)
	}
	if got != "9CC3E6" {
		t.Fatalf("lumMod 60000 + lumOff 40000 of 5B9BD5 = %s, want 9CC3E6", got)
	}
}

func TestExtractNativePPTXMaterializesDocumentedSchemeColorTransforms(t *testing.T) {
	t.Parallel()
	wantTint := applyNativeTint(mustParseNativeSRGB(t, "2F6FED"), 40000).hex()
	wantShade := applyNativeShade(mustParseNativeSRGB(t, "2F6FED"), 75000).hex()
	wantLum, err := applyNativeSchemeColorTransforms("2F6FED", &nativeXMLNode{Children: []*nativeXMLNode{
		{Name: xml.Name{Local: "lumMod"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "60000"}}},
		{Name: xml.Name{Local: "lumOff"}, Attrs: []xml.Attr{{Name: xml.Name{Local: "val"}, Value: "40000"}}},
	}}, nativeExtractDialect{drawing: nsDrawingTransitional})
	if err != nil {
		t.Fatal(err)
	}
	deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/themes/theme.xml"] = nativeExactThemeXML(nsDrawingTransitional)
		parts["relocated/masters/master.xml"] = nativeExactMasterWithColorMapXML(nsPresentationTransitional)
		slide := parts["relocated/slides/slide-a.xml"]
		slide = strings.Replace(slide, `<a:srgbClr val="112233"/>`, `<a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr>`, 1)
		slide = strings.Replace(slide, `<a:srgbClr val="112233"/>`, `<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>`, 1)
		fill := `<a:solidFill><a:schemeClr val="accent1"><a:shade val="75000"/></a:schemeClr></a:solidFill>`
		line := `<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
		slide = strings.Replace(slide, `</p:spTree>`, nativeAutoShapeXML(3, "Tinted shape", "rect", fill, line, "")+`</p:spTree>`, 1)
		parts["relocated/slides/slide-a.xml"] = slide
	}}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract transformed scheme colors: %v", err)
	}
	text := deck.Slides[0].Elements[0]
	if text.Kind != NativeElementKindText || text.Compatibility.Status != NativeCompatibilityStatusEditable || text.Paragraphs == nil || len(*text.Paragraphs) != 1 || len((*text.Paragraphs)[0].Runs) != 2 {
		t.Fatalf("transformed run colors were refused: %#v", text)
	}
	first, second := (*text.Paragraphs)[0].Runs[0], (*text.Paragraphs)[0].Runs[1]
	if first.Color == nil || *first.Color != wantTint {
		t.Fatalf("tint run color = %#v, want %s", first.Color, wantTint)
	}
	if second.Color == nil || *second.Color != wantLum {
		t.Fatalf("lumMod/lumOff run color = %#v, want %s", second.Color, wantLum)
	}
	shapes := nativeFixtureAutoShapes(deck.Slides[0])
	if len(shapes) != 1 || shapes[0].Compatibility.Status != NativeCompatibilityStatusEditable || shapes[0].Fill == nil || *shapes[0].Fill != wantShade {
		t.Fatalf("shade fill was not exact: %#v", shapes)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("transformed color deck invalid: %#v", issues)
	}
}

func mustParseNativeSRGB(t *testing.T, value string) nativeSRGBColor {
	t.Helper()
	color, err := parseNativeSRGBHex(value)
	if err != nil {
		t.Fatal(err)
	}
	return color
}
