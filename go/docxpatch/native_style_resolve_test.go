package docxpatch

import (
	"bytes"
	"os"
	"reflect"
	"strconv"
	"strings"
	"testing"
)

func TestNativeLatentStyleBehaviorIsRenderNeutralButPreserved(t *testing.T) {
	metadata := `<w:latentStyles w:defLockedState="1" w:defUIPriority="99" w:count="2"><w:lsdException w:name="Heading 2" w:locked="false" w:semiHidden="0" w:unhideWhenUsed="true" w:qFormat="1" w:uiPriority="2"/></w:latentStyles>`
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `">%s<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:sz w:val="24"/></w:rPr></w:style></w:styles>`
	baseline, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(strings.Replace(styles, "%s", "", 1)))))
	if err != nil {
		t.Fatal(err)
	}
	parts := resolvedStylesTestParts(strings.Replace(styles, "%s", metadata, 1))
	data := buildNativeDOCX(t, nativeEntries(parts))
	before := append([]byte(nil), data...)
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(data, before) {
		t.Fatal("resolution changed source bytes")
	}
	if !reflect.DeepEqual(baseline.Runs[0].Properties, resolved.Runs[0].Properties) || !reflect.DeepEqual(baseline.Paragraphs[0].Properties, resolved.Paragraphs[0].Properties) {
		t.Fatal("latent metadata changed formatting")
	}
	if len(resolved.Diagnostics) != 1 || resolved.Diagnostics[0].Code != "LATENT_STYLE_BEHAVIOR_PRESERVED" || resolved.Diagnostics[0].Severity != "unsupported" || resolved.Diagnostics[0].Preservation != "preserve-verbatim" {
		t.Fatalf("lost preservation diagnostic: %#v", resolved.Diagnostics)
	}
	for name, markup := range map[string]string{
		"unknown attribute": `<w:latentStyles w:layout="1"/>`,
		"foreign attribute": `<w:latentStyles xmlns:x="urn:foreign" x:count="1"/>`,
		"invalid boolean":   `<w:latentStyles w:defLockedState="maybe"/>`,
		"invalid number":    `<w:latentStyles w:count="-1"/>`,
		"missing name":      `<w:latentStyles><w:lsdException w:locked="1"/></w:latentStyles>`,
		"nested formatting": `<w:latentStyles><w:lsdException w:name="Heading 2"><w:rPr/></w:lsdException></w:latentStyles>`,
		"foreign child":     `<w:latentStyles><x:lsdException xmlns:x="urn:foreign" x:name="Heading 2"/></w:latentStyles>`,
		"text":              `<w:latentStyles>unmodeled</w:latentStyles>`,
	} {
		t.Run(name, func(t *testing.T) {
			got, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(strings.Replace(styles, "%s", markup, 1)))))
			if err != nil {
				return
			} // Namespace spoofing may reject before classification.
			if len(got.Diagnostics) == 0 || got.Diagnostics[0].Code != "LATENT_STYLES_PRESERVED" {
				t.Fatalf("unqualified metadata became neutral: %#v", got.Diagnostics)
			}
		})
	}
	if _, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(strings.Replace(styles, "%s", metadata+metadata, 1))))); err == nil {
		t.Fatal("duplicate latentStyles accepted")
	}
}

func TestNativeStyleDiagnosticsFollowOnlyActiveCascadeConsumers(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="Unused"><w:pPr><w:tabs><w:tab w:val="right" w:pos="1000"/></w:tabs></w:pPr><w:rPr><w:szCs w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Unused"/></w:style></w:styles>`
	// Use an active unsupported effect; szCs is inactive for this Latin fixture.
	styles = strings.Replace(styles, `<w:szCs w:val="24"/>`, `<w:emboss/>`, 1)
	for _, active := range []bool{false, true} {
		parts := resolvedStylesTestParts(styles)
		if active {
			parts["word/document.xml"] = strings.Replace(parts["word/document.xml"], "<w:p>", `<w:p><w:pPr><w:pStyle w:val="Child"/></w:pPr>`, 1)
			// Tab stops now qualify for plain text; retain an actual active tab
			// consumer here so this test still checks source-scoped refusal.
			parts["word/document.xml"] = strings.Replace(parts["word/document.xml"], "<w:r>", "<w:r><w:tab/>", 1)
		}
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := append([]byte(nil), data...)
		resolved, err := ResolveNativeDocumentLayoutV1(data)
		if err != nil {
			t.Fatal(err)
		}
		if !bytes.Equal(before, data) {
			t.Fatal("style resolution mutated source")
		}
		if !active && len(resolved.Diagnostics) != 0 {
			t.Fatalf("unused styles blocked layout: %#v", resolved.Diagnostics)
		}
		if active {
			if len(resolved.Diagnostics) != 2 {
				t.Fatalf("active base style lost diagnostics: %#v", resolved.Diagnostics)
			}
			for _, diagnostic := range resolved.Diagnostics {
				if diagnostic.ScopeID != resolved.Paragraphs[0].ParagraphID || diagnostic.Path == nil || !strings.Contains(*diagnostic.Path, "/w:style[1]/") {
					t.Fatalf("incorrect cascade diagnostic provenance: %#v", diagnostic)
				}
			}
		}
	}
}

func TestDeferredStyleDiagnosticsRemainBounded(t *testing.T) {
	markup := strings.Repeat(`<w:unknown/>`, NativeDOCXMaxResolvedDiagnostics+1)
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="Unused"><w:rPr>` + markup + `</w:rPr></w:style></w:styles>`
	if _, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(styles)))); err == nil || !strings.Contains(err.Error(), "diagnostics exceed") {
		t.Fatalf("unbounded inactive style diagnostics: %v", err)
	}
}

func TestNativeBasicLatinFontSlotsRespectCascadeAndRefuseMixedScripts(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Base ASCII" w:hAnsi="Base High"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:rFonts w:ascii="Paragraph ASCII"/></w:rPr></w:style><w:style w:type="character" w:styleId="Em"><w:rPr><w:rFonts w:ascii="Character ASCII"/></w:rPr></w:style></w:styles>`
	for _, test := range []struct {
		name, text, rpr, font string
		refused               bool
	}{
		{"paragraph", "Basic 123", "", "Paragraph ASCII", false},
		{"character", "Basic 123", `<w:rStyle w:val="Em"/>`, "Character ASCII", false},
		{"direct", "Basic 123", `<w:rStyle w:val="Em"/><w:rFonts w:ascii="Direct ASCII"/>`, "Direct ASCII", false},
		{"non-latin", "Basic العربية", "", "", true},
		{"high-ansi", "café", "", "", true},
		{"rtl", "Basic", `<w:rtl/>`, "", true},
		{"equal-slots", "café", `<w:rFonts w:ascii="Equal" w:hAnsi="Equal"/>`, "Equal", false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = strings.Replace(parts["word/document.xml"], `<w:r><w:t>test</w:t></w:r>`, `<w:r><w:rPr>`+test.rpr+`</w:rPr><w:t>`+test.text+`</w:t></w:r>`, 1)
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			font := resolved.Runs[0].Properties.FontFamily
			if test.refused {
				if font != nil {
					t.Fatalf("guessed font %q", *font)
				}
				if len(resolved.Diagnostics) == 0 || resolved.Diagnostics[0].Code != "SCRIPT_DEPENDENT_LATIN_FONT" {
					t.Fatalf("missing slot refusal: %#v", resolved.Diagnostics)
				}
			} else if font == nil || *font != test.font || len(resolved.Diagnostics) != 0 {
				t.Fatalf("font=%v diagnostics=%#v", font, resolved.Diagnostics)
			}
			if mark := resolved.Paragraphs[0].ParagraphMarkProperties.FontFamily; mark == nil || *mark != "Paragraph ASCII" {
				t.Fatalf("paragraph mark lost its own ASCII cascade: %v", mark)
			}
		})
	}
}

func TestResolveNativeDocumentLayoutV1CascadePrecedenceAndToggles(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/lists/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="WORD/document.xml"/></Relationships>`,
		"word/document.xml":            `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/><w:numPr><w:ilvl w:val="0"/></w:numPr><w:jc w:val="right"/><w:spacing w:before="360" w:beforeAutospacing="false"/><w:ind w:left="720" w:hanging="360"/></w:pPr><w:r><w:rPr><w:rStyle w:val="Em"/><w:b w:val="false"/><w:sz w:val="28"/><w:color w:val="FF0000"/></w:rPr><w:t>Cascade</w:t></w:r><w:r><w:t>Default character</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="styles" Type="` + relBaseTransitional + `styles" Target="STYLES.xml"/><Relationship Id="numbering" Type="` + relBaseTransitional + `numbering" Target="lists/NUMBERING.xml"/></Relationships>`,
		"word/Styles.XML":              `<w:styles xmlns:w="` + wordMLTransitional + `" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/><w:color w:val="111111"/><w:b/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="left"/><w:spacing w:before="120"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:pPr><w:jc w:val="center"/><w:spacing w:after="240"/></w:pPr><w:rPr><w:i/><w:b w:val="false"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/><w:pPr><w:numPr><w:numId w:val="7"/></w:numPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:i/></w:rPr></w:style><w:style w:type="character" w:styleId="Em"><w:rPr><w:b/><w:i/></w:rPr></w:style><w:style w:type="character" w:styleId="DefaultChar" w:default="1"><w:rPr><w:u w:val="double"/><w:vanish/></w:rPr></w:style><w:style w:type="numbering" w:styleId="ListLink"/><mc:AlternateContent/></w:styles>`,
		"word/Lists/Numbering.XML":     `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="3"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="space"/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier" w:hAnsi="Courier"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="7"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="4"/><w:lvl w:ilvl="0"><w:start w:val="2"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%1)"/><w:suff w:val="tab"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="480" w:hanging="240"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/><w:color w:val="112233"/></w:rPr></w:lvl></w:lvlOverride></w:num></w:numbering>`,
	}
	data := buildNativeDOCX(t, nativeEntries(parts))
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if resolved.SourceParts.StylesPart == nil || *resolved.SourceParts.StylesPart != "word/Styles.XML" || resolved.SourceParts.NumberingPart == nil || *resolved.SourceParts.NumberingPart != "word/Lists/Numbering.XML" {
		t.Fatalf("case-equivalent related parts were not resolved: %#v", resolved.SourceParts)
	}
	if len(resolved.Paragraphs) != 1 || len(resolved.Runs) != 2 {
		t.Fatalf("unexpected resolved objects: paragraphs=%d runs=%d", len(resolved.Paragraphs), len(resolved.Runs))
	}
	paragraph := resolved.Paragraphs[0]
	if got := strings.Join(paragraph.AppliedStyles, ","); got != "Base,Child" {
		t.Fatalf("paragraph style cascade = %q", got)
	}
	if paragraph.Properties.Alignment == nil || *paragraph.Properties.Alignment != "right" || paragraph.Properties.SpacingBeforeTwips == nil || *paragraph.Properties.SpacingBeforeTwips != 360 || paragraph.Properties.SpacingAfterTwips == nil || *paragraph.Properties.SpacingAfterTwips != 240 || paragraph.Properties.IndentLeftTwips == nil || *paragraph.Properties.IndentLeftTwips != 720 || paragraph.Properties.HangingTwips == nil || *paragraph.Properties.HangingTwips != 360 || paragraph.Properties.KeepNext == nil || !*paragraph.Properties.KeepNext {
		t.Fatalf("paragraph precedence was not resolved: %#v", paragraph.Properties)
	}
	if paragraph.Numbering == nil || paragraph.Numbering.NumID != "7" || paragraph.Numbering.AbstractNumID != "3" || paragraph.Numbering.Start != 4 || paragraph.Numbering.Format != "upperRoman" || paragraph.Numbering.Text != "%1)" || paragraph.Numbering.Suffix != "tab" || paragraph.Numbering.ResolvedText != "IV)" || paragraph.Numbering.Marker.FontFamily == nil || *paragraph.Numbering.Marker.FontFamily != "Courier New" {
		t.Fatalf("numbering level/override was not resolved: %#v", paragraph.Numbering)
	}
	if paragraph.ParagraphMarkProperties.FontFamily == nil || *paragraph.ParagraphMarkProperties.FontFamily != "Calibri" || paragraph.ParagraphMarkProperties.FontSizeHalfPoint == nil || *paragraph.ParagraphMarkProperties.FontSizeHalfPoint != 22 {
		t.Fatalf("paragraph-mark defaults/style properties were not resolved: %#v", paragraph.ParagraphMarkProperties)
	}
	run := resolved.Runs[0]
	if paragraphStyles, characterStyles := strings.Join(run.AppliedParagraphStyles, ","), strings.Join(run.AppliedCharacterStyles, ","); paragraphStyles != "Base,Child" || characterStyles != "Em" {
		t.Fatalf("run style cascade = paragraph %q, character %q", paragraphStyles, characterStyles)
	}
	if run.Properties.FontFamily == nil || *run.Properties.FontFamily != "Calibri" || run.Properties.FontSizeHalfPoint == nil || *run.Properties.FontSizeHalfPoint != 28 || run.Properties.Color == nil || *run.Properties.Color != "FF0000" {
		t.Fatalf("run scalar precedence was not resolved: %#v", run.Properties)
	}
	if run.Properties.Bold == nil || *run.Properties.Bold || run.Properties.Italic == nil || !*run.Properties.Italic {
		t.Fatalf("toggle/direct precedence was not resolved: %#v", run.Properties)
	}
	defaultRun := resolved.Runs[1]
	if defaultRun.CharacterStyle == nil || *defaultRun.CharacterStyle != "DefaultChar" || strings.Join(defaultRun.AppliedCharacterStyles, ",") != "DefaultChar" || defaultRun.Properties.Underline == nil || *defaultRun.Properties.Underline != "double" || defaultRun.Properties.Hidden == nil || !*defaultRun.Properties.Hidden {
		t.Fatalf("default character style was not applied: %#v", defaultRun)
	}
	if !hasResolutionDiagnostic(resolved, "NUMBERING_STYLE_PRESERVED") || !hasResolutionDiagnostic(resolved, "FOREIGN_STYLES_MARKUP") {
		t.Fatalf("preserved style metadata was not diagnosed: %#v", resolved.Diagnostics)
	}
	first, err := EncodeNativeResolvedLayoutInputV1(resolved)
	if err != nil {
		t.Fatal(err)
	}
	again, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	second, _ := EncodeNativeResolvedLayoutInputV1(again)
	if !bytes.Equal(first, second) {
		t.Fatal("resolved layout input is not deterministic")
	}
}

func TestResolveNativeDocumentLayoutV1ExportsDirectParagraphMarkProperties(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="20"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:rPr><w:rFonts w:ascii="Aptos" w:hAnsi="Aptos"/><w:sz w:val="30"/><w:i/></w:rPr></w:pPr></w:p><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Paragraphs) != 1 || len(resolved.Runs) != 0 {
		t.Fatalf("unexpected blank paragraph projection: %#v", resolved)
	}
	mark := resolved.Paragraphs[0].ParagraphMarkProperties
	if mark.FontFamily == nil || *mark.FontFamily != "Aptos" || mark.FontSizeHalfPoint == nil || *mark.FontSizeHalfPoint != 30 || mark.Italic == nil || !*mark.Italic {
		t.Fatalf("direct paragraph mark formatting was not resolved: %#v", mark)
	}
	encoded, err := EncodeNativeResolvedLayoutInputV1(resolved)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(encoded, []byte(`"paragraph_mark_properties":{"font_family":"Aptos","font_size_half_points":30,"italic":true}`)) {
		t.Fatalf("paragraph mark binding absent from deterministic JSON: %s", encoded)
	}
}

func hasResolutionDiagnostic(input *NativeResolvedLayoutInputV1, code string) bool {
	for _, diagnostic := range input.Diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

func TestResolveNativeDocumentLayoutV1BrokenBasedOnRetainsSafeDescendants(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="MissingChild"><w:basedOn w:val="Absent"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="CycleA"><w:basedOn w:val="CycleB"/><w:pPr><w:keepNext/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="CycleB"><w:basedOn w:val="CycleA"/><w:pPr><w:keepLines/></w:pPr></w:style><w:style w:type="paragraph" w:styleId="Outer"><w:basedOn w:val="CycleA"/><w:pPr><w:pageBreakBefore/></w:pPr></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="MissingChild"/></w:pPr><w:r><w:t>missing</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="Outer"/></w:pPr><w:r><w:t>cycle</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if got := strings.Join(resolved.Paragraphs[0].AppliedStyles, ","); got != "MissingChild" || resolved.Paragraphs[0].Properties.Alignment == nil || *resolved.Paragraphs[0].Properties.Alignment != "right" {
		t.Fatalf("missing ancestor discarded safe descendant: %#v", resolved.Paragraphs[0])
	}
	if got := strings.Join(resolved.Paragraphs[1].AppliedStyles, ","); got != "Outer" || resolved.Paragraphs[1].Properties.PageBreakBefore == nil || !*resolved.Paragraphs[1].Properties.PageBreakBefore || resolved.Paragraphs[1].Properties.KeepNext != nil || resolved.Paragraphs[1].Properties.KeepLines != nil {
		t.Fatalf("cycle did not retain only safe descendant: %#v", resolved.Paragraphs[1])
	}
	if !hasResolutionDiagnostic(resolved, "MISSING_STYLE_REFERENCE") || !hasResolutionDiagnostic(resolved, "STYLE_BASED_ON_CYCLE") {
		t.Fatalf("broken basedOn chains were not diagnosed: %#v", resolved.Diagnostics)
	}
}

func TestResolveNativeDocumentLayoutV1StrictRelocatedThemeAndFontTable(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml":     `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/odd/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/odd/assets/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/odd/assets/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/odd/assets/fonts.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>`,
		"_RELS/.RELS":             `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseStrict + `officeDocument" Target="ODD/MAIN.xml"/></Relationships>`,
		"Odd/Main.XML":            `<w:document xmlns:w="` + wordMLStrict + `"><w:body><w:p><w:pPr><w:bidi/><w:ind w:start="900" w:end="-120"/></w:pPr><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:eastAsia="MS Mincho"/><w:color w:themeColor="accent1"/><w:lang w:val="en-US" w:eastAsia="ja-JP"/><w:rtl/><w:vanish/><w:highlight w:val="chartreuse"/></w:rPr><w:t>Strict</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"Odd/_RELS/Main.XML.RELS": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="styles" Type="` + relBaseStrict + `styles" Target="assets/STYLES.xml"/><Relationship Id="theme" Type="` + relBaseStrict + `theme" Target="assets/THEME.xml"/><Relationship Id="fonts" Type="` + relBaseStrict + `fontTable" Target="assets/FONTS.xml"/></Relationships>`,
		"Odd/Assets/Styles.XML":   `<w:styles xmlns:w="` + wordMLStrict + `"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
		"Odd/Assets/Theme.XML":    `<a:theme xmlns:a="` + drawingMLStrict + `" name="Strict theme"><a:themeElements/></a:theme>`,
		"Odd/Assets/Fonts.XML":    `<w:fonts xmlns:w="` + wordMLStrict + `" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:font w:name="Calibri"><w:altName w:val="Carlito"/><w:panose1 w:val="020F0502020204030204"/></w:font><mc:AlternateContent/></w:fonts>`,
	}
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if resolved.SourceParts.StylesPart == nil || *resolved.SourceParts.StylesPart != "Odd/Assets/Styles.XML" || resolved.SourceParts.ThemePart == nil || *resolved.SourceParts.ThemePart != "Odd/Assets/Theme.XML" || resolved.SourceParts.FontTablePart == nil || *resolved.SourceParts.FontTablePart != "Odd/Assets/Fonts.XML" {
		t.Fatalf("strict relocated relationships were not resolved: %#v", resolved.SourceParts)
	}
	if len(resolved.Fonts) != 1 || resolved.Fonts[0].Name != "Calibri" || resolved.Fonts[0].AltName == nil || *resolved.Fonts[0].AltName != "Carlito" {
		t.Fatalf("font table was not projected: %#v", resolved.Fonts)
	}
	run := resolved.Runs[0].Properties
	if run.FontFamily != nil || run.Color != nil || run.Highlight != nil || run.FontSizeHalfPoint == nil || *run.FontSizeHalfPoint != 24 || run.Language == nil || *run.Language != "en-US" || run.RTL == nil || !*run.RTL || run.Hidden == nil || !*run.Hidden {
		t.Fatalf("strict explicit/theme properties were not conservative: %#v", run)
	}
	paragraph := resolved.Paragraphs[0].Properties
	if paragraph.Bidi == nil || !*paragraph.Bidi || paragraph.IndentStartTwips == nil || *paragraph.IndentStartTwips != 900 || paragraph.IndentEndTwips == nil || *paragraph.IndentEndTwips != -120 {
		t.Fatalf("paragraph bidi/logical indents were dropped: %#v", paragraph)
	}
	for _, code := range []string{"THEME_FONT_PRESERVED", "THEME_COLOR_PRESERVED", "SCRIPT_FONT_PRESERVED", "SCRIPT_LANGUAGE_PRESERVED", "UNSUPPORTED_HIGHLIGHT", "FOREIGN_FONT_TABLE_MARKUP", "FONT_MATCHING_METADATA_PRESERVED"} {
		if !hasResolutionDiagnostic(resolved, code) {
			t.Fatalf("missing %s diagnostic: %#v", code, resolved.Diagnostics)
		}
	}
}

func nativeThemeSrgbFixtureXML(ns string) string {
	return `<a:theme xmlns:a="` + ns + `" name="Office Theme"><a:themeElements><a:clrScheme name="Office">` +
		`<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
		`<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="44546A"/></a:dk2>` +
		`<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
		`<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>` +
		`<a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>` +
		`<a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink>` +
		`<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"/></a:themeElements></a:theme>`
}

func TestResolveNativeDocumentLayoutV1ResolvesExactThemeSrgbColors(t *testing.T) {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relBaseTransitional+`theme" Target="theme/theme1.xml"/></Relationships>`, 1)
	parts["word/theme/theme1.xml"] = nativeThemeSrgbFixtureXML(drawingMLTransitional)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` +
		`<w:p><w:r><w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr><w:t>stale cache</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:color w:themeColor="hyperlink"/></w:rPr><w:t>theme only</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:color w:themeColor="text1"/></w:rPr><w:t>sysClr</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:color w:val="112233" w:themeColor="accent1" w:themeTint="99"/></w:rPr><w:t>tint</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:color w:val="00AAFF"/></w:rPr><w:t>explicit</w:t></w:r></w:p>` +
		`<w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Runs) != 5 {
		t.Fatalf("unexpected run count: %#v", resolved.Runs)
	}
	if resolved.Runs[0].Properties.Color == nil || *resolved.Runs[0].Properties.Color != "4472C4" {
		t.Fatalf("stale themeColor cache was not replaced by exact srgb: %#v", resolved.Runs[0].Properties)
	}
	if resolved.Runs[1].Properties.Color == nil || *resolved.Runs[1].Properties.Color != "0563C1" {
		t.Fatalf("themeColor-only hyperlink was not resolved from srgb: %#v", resolved.Runs[1].Properties)
	}
	if resolved.Runs[2].Properties.Color == nil || *resolved.Runs[2].Properties.Color != "000000" {
		t.Fatalf("sysClr lastClr was not projected from the exact srgb snapshot: %#v", resolved.Runs[2].Properties)
	}
	if resolved.Runs[3].Properties.Color != nil {
		t.Fatalf("themeTint was guessed: %#v", resolved.Runs[3].Properties)
	}
	if resolved.Runs[4].Properties.Color == nil || *resolved.Runs[4].Properties.Color != "00AAFF" {
		t.Fatalf("explicit RGB color was lost: %#v", resolved.Runs[4].Properties)
	}
	for _, run := range resolved.Runs[:3] {
		if hasRunThemeColorDiagnostic(resolved, run.RunID) {
			t.Fatalf("exact theme srgb still diagnosed as preserve-only: %#v", resolved.Diagnostics)
		}
	}
	if !hasResolutionDiagnostic(resolved, "THEME_COLOR_PRESERVED") {
		t.Fatalf("themeTint was not fail-closed: %#v", resolved.Diagnostics)
	}
}

func TestResolveNativeDocumentLayoutV1ResolvesExactThemeLatinFonts(t *testing.T) {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relBaseTransitional+`theme" Target="theme/theme1.xml"/></Relationships>`, 1)
	parts["word/theme/theme1.xml"] = nativeThemeSrgbFixtureXML(drawingMLTransitional)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` +
		`<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/></w:rPr><w:t>theme</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="majorHAnsi"/></w:rPr><w:t>major</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorEastAsia"/></w:rPr><w:t>script</w:t></w:r></w:p>` +
		`<w:p><w:r><w:rPr><w:rFonts w:ascii="Fixture Sans"/></w:rPr><w:t>explicit</w:t></w:r></w:p>` +
		`<w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Runs) != 4 {
		t.Fatalf("unexpected run count: %#v", resolved.Runs)
	}
	if resolved.Runs[0].Properties.FontFamily == nil || *resolved.Runs[0].Properties.FontFamily != "Calibri" {
		t.Fatalf("stale ascii cache was not replaced by exact theme latin typeface: %#v", resolved.Runs[0].Properties)
	}
	if resolved.Runs[1].Properties.FontFamily == nil || *resolved.Runs[1].Properties.FontFamily != "Calibri Light" {
		t.Fatalf("majorHAnsi was not resolved from the theme latin typeface: %#v", resolved.Runs[1].Properties)
	}
	if resolved.Runs[2].Properties.FontFamily != nil || !hasResolutionDiagnostic(resolved, "THEME_FONT_PRESERVED") {
		t.Fatalf("script theme font slot was guessed: %#v", resolved.Runs[2].Properties)
	}
	if resolved.Runs[3].Properties.FontFamily == nil || *resolved.Runs[3].Properties.FontFamily != "Fixture Sans" {
		t.Fatalf("explicit latin font was lost: %#v", resolved.Runs[3].Properties)
	}
}

func TestResolveNativeDocumentLayoutV1ProjectsSimpleTableStyleBorders(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="table" w:styleId="PlainBorders"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:themeColor="accent1"/><w:left w:val="single" w:sz="8" w:space="0" w:color="112233"/></w:tblBorders></w:tblPr><w:tcPr><w:shd w:val="clear" w:color="auto" w:fill="DDEEFF"/></w:tcPr></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relBaseTransitional+`theme" Target="theme/theme1.xml"/></Relationships>`, 1)
	parts["word/theme/theme1.xml"] = nativeThemeSrgbFixtureXML(drawingMLTransitional)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="PlainBorders"/><w:tblW w:w="2400" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>styled</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Tables) != 1 || resolved.Tables[0].StyleID == nil || *resolved.Tables[0].StyleID != "PlainBorders" {
		t.Fatalf("simple table style identity was lost: %#v", resolved.Tables)
	}
	if resolved.Tables[0].Borders == nil || resolved.Tables[0].Borders.Top == nil || resolved.Tables[0].Borders.Top.ColorRGB == nil || *resolved.Tables[0].Borders.Top.ColorRGB != "4472C4" {
		t.Fatalf("simple table style theme border was not projected: %#v", resolved.Tables[0].Borders)
	}
	if resolved.Tables[0].CellShadingRGB == nil || *resolved.Tables[0].CellShadingRGB != "DDEEFF" {
		t.Fatalf("simple table style cell fill was not projected: %#v", resolved.Tables[0])
	}
	if hasResolutionDiagnostic(resolved, "TABLE_STYLE_EFFECTS_PRESERVED") || hasResolutionDiagnostic(resolved, "CONDITIONAL_TABLE_STYLE_PRESERVED") {
		t.Fatalf("simple table style was treated as unmodeled: %#v", resolved.Diagnostics)
	}
}

func hasRunThemeColorDiagnostic(input *NativeResolvedLayoutInputV1, runID string) bool {
	for _, diagnostic := range input.Diagnostics {
		if diagnostic.Code == "THEME_COLOR_PRESERVED" && diagnostic.ScopeID == runID {
			return true
		}
	}
	return false
}

func TestResolveNativeDocumentLayoutV1DelegatesEmbeddedFontResourceBindings(t *testing.T) {
	data := nativeFontTestPackage(t, false, nativeFontTestSFNT(0x0008), "false", nil)
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Fonts) != 1 || resolved.Fonts[0].Name != "Fixture Sans" || len(resolved.Runs) != 1 || resolved.Runs[0].Properties.FontFamily == nil || *resolved.Runs[0].Properties.FontFamily != "Fixture Sans" {
		t.Fatalf("embedded font family/style cascade was not preserved: %#v", resolved)
	}
	if hasResolutionDiagnostic(resolved, "UNMODELED_FONT_METADATA") {
		t.Fatalf("exact font resource binding incorrectly blocked shaping: %#v", resolved.Diagnostics)
	}
	inventory, err := ExtractNativeDOCXFontInventoryV1(data)
	if err != nil || inventory.NativeTextManifest == nil || len(inventory.NativeTextManifest.Faces) != 1 {
		t.Fatalf("delegated exact font inventory is unavailable: inventory=%#v err=%v", inventory, err)
	}

	unattested := nativeFontTestPackage(t, false, nativeFontTestSFNT(0x0008), "false", func(parts map[string]string) {
		parts["Word/FontTable.XML"] = strings.Replace(parts["Word/FontTable.XML"], nativeFontTestKey, strings.ToLower(nativeFontTestKey), 1)
	})
	resolved, err = ResolveNativeDocumentLayoutV1(unattested)
	if err != nil {
		t.Fatal(err)
	}
	if !hasResolutionDiagnostic(resolved, "UNATTESTED_EMBEDDED_FONT_BINDING") {
		t.Fatalf("unproven embed metadata retained blanket diagnostic suppression: %#v", resolved.Diagnostics)
	}
}

func TestResolveNativeDocumentLayoutV1MissingOptionalPartsAndNumbering(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":   `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>defaults</w:t></w:r></w:p><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="99"/></w:numPr></w:pPr><w:r><w:t>missing list</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	}
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if resolved.SourceParts.StylesPart != nil || resolved.SourceParts.NumberingPart != nil || resolved.SourceParts.ThemePart != nil || resolved.SourceParts.FontTablePart != nil {
		t.Fatalf("missing optional parts were fabricated: %#v", resolved.SourceParts)
	}
	if resolved.Paragraphs[0].StyleID != nil || resolved.Paragraphs[0].Numbering != nil || resolved.Paragraphs[1].Numbering != nil || !hasResolutionDiagnostic(resolved, "MISSING_NUMBERING_INSTANCE") {
		t.Fatalf("missing defaults/numbering was guessed or silent: %#v", resolved)
	}
}

func TestResolveNativeDocumentLayoutV1MakesEveryTableStyleEffectExplicit(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="table" w:styleId="TableGrid"><w:tblPr><w:tblBorders/></w:tblPr><w:tblStylePr w:type="firstRow"><w:rPr><w:b/></w:rPr></w:tblStylePr></w:style><w:style w:type="table" w:styleId="DefaultTable" w:default="1"><w:tblPr><w:tblCellMar/></w:tblPr></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>explicit</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:tbl><w:tblPr/><w:tr><w:tc><w:p><w:r><w:t>default</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Tables) != 2 || resolved.Tables[0].StyleID == nil || *resolved.Tables[0].StyleID != "TableGrid" || resolved.Tables[1].StyleID != nil {
		t.Fatalf("tblStyle omission must not apply the authoring default: %#v", resolved.Tables)
	}
	if !hasResolutionDiagnostic(resolved, "TABLE_STYLE_EFFECTS_PRESERVED") || !hasResolutionDiagnostic(resolved, "CONDITIONAL_TABLE_STYLE_PRESERVED") {
		t.Fatalf("table-style effects were silent: %#v", resolved.Diagnostics)
	}
}

func TestResolveNativeDocumentLayoutV1RejectsAmbiguousStyleParts(t *testing.T) {
	tests := []struct {
		name        string
		styles      string
		contentType string
		want        string
	}{
		{"duplicate style", `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="A"/><w:style w:type="paragraph" w:styleId="A"/></w:styles>`, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", "duplicate paragraph style"},
		{"namespace spoof", `<w:styles xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:spoof"><x:style w:type="paragraph" w:styleId="A"/></w:styles>`, "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", "namespace spoofing"},
		{"wrong content type", `<w:styles xmlns:w="` + wordMLTransitional + `"/>`, "application/xml", "has content type"},
		{"Unicode-folded content type", `<w:styles xmlns:w="` + wordMLTransitional + `"/>`, "application/vnd.openxmlformats-officedocument.wordprocessingml.style\u017f+xml", "has content type"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(test.styles)
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", test.contentType, 1)
			_, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("error = %v, want %q", err, test.want)
			}
		})
	}
}

func TestResolveNativeDocumentLayoutV1RejectsAmbiguousNumberingParts(t *testing.T) {
	tests := []struct {
		name      string
		numbering string
		want      string
	}{
		{"duplicate abstract", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"/><w:abstractNum w:abstractNumId="1"/></w:numbering>`, "duplicate abstractNum"},
		{"duplicate level", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"/><w:lvl w:ilvl="0"/></w:abstractNum></w:numbering>`, "duplicate level"},
		{"duplicate num", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`, "duplicate num id"},
		{"duplicate level override", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"/><w:lvlOverride w:ilvl="0"/></w:num></w:numbering>`, "duplicate level override"},
		{"missing multilevel value", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:multiLevelType/></w:abstractNum></w:numbering>`, "invalid multiLevelType"},
		{"unknown multilevel value", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="outline"/></w:abstractNum></w:numbering>`, "invalid multiLevelType"},
		{"duplicate multilevel metadata", `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="singleLevel"/><w:multiLevelType w:val="multilevel"/></w:abstractNum></w:numbering>`, "duplicate multiLevelType"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedNumberingTestParts(test.numbering))))
			if err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("error = %v, want %q", err, test.want)
			}
		})
	}
}

func TestResolveNativeDocumentLayoutV1LeavesDuplicateSingletonsUnresolved(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:pPrDefault><w:pPr><w:jc w:val="left"/><w:spacing w:before="100"/></w:pPr></w:pPrDefault><w:rPrDefault><w:rPr><w:b/><w:color w:val="111111"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/fonts.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="fonts" Type="`+relBaseTransitional+`fontTable" Target="fonts.xml"/></Relationships>`, 1)
	parts["word/fonts.xml"] = `<w:fonts xmlns:w="` + wordMLTransitional + `"><w:font w:name="Calibri"><w:altName w:val="Carlito"/><w:altName w:val="MetricCompatible"/></w:font></w:fonts>`
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:jc w:val="center"/><w:jc w:val="right"/><w:spacing w:before="200"/><w:spacing w:before="300"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:rPr><w:b w:val="false"/><w:b/><w:color w:val="AAAAAA"/><w:color w:val="BBBBBB"/></w:rPr><w:t>duplicates</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	paragraph := resolved.Paragraphs[0]
	if paragraph.Properties.Alignment == nil || *paragraph.Properties.Alignment != "left" || paragraph.Properties.SpacingBeforeTwips == nil || *paragraph.Properties.SpacingBeforeTwips != 100 || paragraph.Numbering != nil {
		t.Fatalf("duplicate paragraph/numPr singleton became last-wins: %#v", paragraph)
	}
	run := resolved.Runs[0].Properties
	if run.Bold == nil || !*run.Bold || run.Color == nil || *run.Color != "111111" {
		t.Fatalf("duplicate run singleton became last-wins: %#v", run)
	}
	if len(resolved.Fonts) != 1 || resolved.Fonts[0].AltName != nil {
		t.Fatalf("duplicate font alias became first-wins: %#v", resolved.Fonts)
	}
	for _, code := range []string{"DUPLICATE_PARAGRAPH_PROPERTY", "DUPLICATE_RUN_PROPERTY", "DUPLICATE_NUMBERING_REFERENCE", "DUPLICATE_FONT_ALT_NAME"} {
		if !hasResolutionDiagnostic(resolved, code) {
			t.Fatalf("missing %s diagnostic: %#v", code, resolved.Diagnostics)
		}
	}
}

func TestResolveNativeDocumentLayoutV1PreservesPictureBullets(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:numPicBullet w:numPicBulletId="1"><w:pict/></w:numPicBullet><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlPicBulletId w:val="1"/></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	parts := resolvedNumberingTestParts(numbering)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:numPr><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>picture bullet</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if resolved.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(resolved, "PICTURE_BULLET_PRESERVED") {
		t.Fatalf("picture bullet was guessed or silent: %#v", resolved)
	}
}

func TestResolveNativeDocumentLayoutV1BoundsStylesAndAvoidsDOMDependencies(t *testing.T) {
	var styles strings.Builder
	styles.WriteString(`<w:styles xmlns:w="` + wordMLTransitional + `">`)
	for index := 0; index <= NativeDOCXMaxCollectionItems; index++ {
		styles.WriteString(`<w:style w:type="paragraph" w:styleId="S`)
		styles.WriteString(strconv.Itoa(index))
		styles.WriteString(`"/>`)
	}
	styles.WriteString(`</w:styles>`)
	parts := resolvedStylesTestParts(styles.String())
	_, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err == nil || !strings.Contains(err.Error(), "exceeds") {
		t.Fatalf("style resource bound error = %v", err)
	}
	source, err := os.ReadFile("native_style_resolve.go")
	if err != nil {
		t.Fatal(err)
	}
	for _, dependency := range []string{"golang.org/x/net/html", "goquery", "chromedp", "playwright", "mammoth"} {
		if bytes.Contains(bytes.ToLower(source), []byte(dependency)) {
			t.Fatalf("style resolver must remain HTML/DOM-free; found %q", dependency)
		}
	}
}

func TestValidateNativeResolvedLayoutInputV1RejectsUnsafeMutations(t *testing.T) {
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedStylesTestParts(`<w:styles xmlns:w="`+wordMLTransitional+`"/>`))))
	if err != nil {
		t.Fatal(err)
	}
	tests := []struct {
		name   string
		mutate func(*NativeResolvedLayoutInputV1)
	}{
		{"unsafe source part", func(input *NativeResolvedLayoutInputV1) { input.SourceParts.MainPart = "word/../evil.xml" }},
		{"duplicate paragraph", func(input *NativeResolvedLayoutInputV1) {
			input.Paragraphs = append(input.Paragraphs, input.Paragraphs[0])
		}},
		{"dangling run", func(input *NativeResolvedLayoutInputV1) { input.Runs[0].ParagraphID = "paragraph:missing" }},
		{"invalid font size", func(input *NativeResolvedLayoutInputV1) { input.Runs[0].Properties.FontSizeHalfPoint = nativeInt(0) }},
		{"invalid paragraph mark font size", func(input *NativeResolvedLayoutInputV1) {
			input.Paragraphs[0].ParagraphMarkProperties.FontSizeHalfPoint = nativeInt(0)
		}},
		{"duplicate style trace", func(input *NativeResolvedLayoutInputV1) { input.Paragraphs[0].AppliedStyles = []string{"A", "A"} }},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			copyInput := *resolved
			copyInput.Paragraphs = append([]NativeResolvedParagraphV1{}, resolved.Paragraphs...)
			copyInput.Runs = append([]NativeResolvedRunV1{}, resolved.Runs...)
			test.mutate(&copyInput)
			if err := ValidateNativeResolvedLayoutInputV1(&copyInput); err == nil {
				t.Fatal("unsafe resolved-layout mutation was accepted")
			}
		})
	}
}

func resolvedStylesTestParts(styles string) map[string]string {
	return map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":            `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>test</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="styles" Type="` + relBaseTransitional + `styles" Target="styles.xml"/></Relationships>`,
		"word/styles.xml":              styles,
	}
}

func resolvedNumberingTestParts(numbering string) map[string]string {
	return map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":            `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>test</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="numbering" Type="` + relBaseTransitional + `numbering" Target="numbering.xml"/></Relationships>`,
		"word/numbering.xml":           numbering,
	}
}
