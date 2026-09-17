package docxpatch

import (
	"strings"
	"testing"
)

func nativeMode15Compat() string {
	return `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>`
}

func nativeAttestedMathPr() string {
	return `<m:mathPr><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>`
}

func nativeAttestedShapeDefaults() string {
	return `<w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults>`
}

func nativeImageCropNeutralExtras() string {
	// Recreates the ImageCrop.docx settings extras that are attested-neutral,
	// with explicit compatibilityMode=15 in place of the original empty compat.
	return `<w:zoom w:percent="100"/><w:proofState w:spelling="clean" w:grammar="clean"/><w:stylePaneFormatFilter w:val="3F01"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>` +
		nativeMode15Compat() +
		`<w:rsids><w:rsidRoot w:val="001E052D"/><w:rsid w:val="001E052D"/></w:rsids>` +
		nativeAttestedMathPr() +
		`<w:themeFontLang w:val="en-US"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/><w:doNotIncludeSubdocsInStats/>` +
		nativeAttestedShapeDefaults() +
		`<w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>`
}

func nativeTwoColHeaderNeutralExtras() string {
	// Recreates the 2col-header.docx extras that are attested-neutral, with
	// compatibilityMode=15 instead of 14 and without the layout-affecting
	// extra compatSetting flags.
	return `<w:zoom w:percent="150"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr>` +
		nativeMode15Compat() +
		`<w:rsids><w:rsidRoot w:val="002D662D"/><w:rsid w:val="002D662D"/></w:rsids>` +
		strings.Replace(nativeAttestedMathPr(), `m:val="0"`, `m:val="off"`, 1) +
		`<w:themeFontLang w:val="en-US" w:eastAsia="x-none" w:bidi="x-none"/><w:clrSchemeMapping w:bg1="light1" w:t1="dark1" w:bg2="light2" w:t2="dark2" w:accent1="accent1" w:accent2="accent2" w:accent3="accent3" w:accent4="accent4" w:accent5="accent5" w:accent6="accent6" w:hyperlink="hyperlink" w:followedHyperlink="followedHyperlink"/>` +
		nativeAttestedShapeDefaults() +
		`<w:decimalSymbol w:val=","/><w:listSeparator w:val=","/>`
}

func extractNativePaginationSettingsMarkup(t *testing.T, wordNS, inner string) *NativePaginationSettingsV1 {
	t.Helper()
	settingsXML := `<w:settings xmlns:w="` + wordNS + `" xmlns:m="` + nativeMathNamespace + `" xmlns:o="` + nativeOfficeNamespace + `" xmlns:v="` + nativeVMLNamespace + `" xmlns:w14="` + nativeWord14Namespace + `" xmlns:w15="` + nativeWord15Namespace + `">` + inner + `</w:settings>`
	parts := nativePaginationSettingsParts(settingsXML)
	if wordNS == wordMLStrict {
		parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
		parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
		parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict, 1)
	}
	settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	return settings
}

func paginationSettingsHasCode(settings *NativePaginationSettingsV1, code string) bool {
	for _, diagnostic := range settings.Diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

func TestNativeExtractOmitsUnsafeAttachedTemplate(t *testing.T) {
	settingsXML := `<w:settings xmlns:w="` + wordMLTransitional + `"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
	parts := nativePaginationSettingsParts(settingsXML)
	parts["Word/_rels/Settings.XML.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate" Target="file:///f:\dsbuildroot\global.doc.dotx" TargetMode="External"/></Relationships>`
	if _, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts))); err != nil {
		t.Fatal(err)
	}
	hyper := nativePaginationSettingsParts(settingsXML)
	hyper["Word/_RELS/Document.XML.RELS"] = strings.Replace(hyper["Word/_RELS/Document.XML.RELS"], `Target="SETTINGS.xml"`, `Target="file:///f:\unsafe\settings.xml" TargetMode="External"`, 1)
	if _, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(hyper))); err == nil {
		t.Fatal("unsafe non-template external relationship must still fail extract")
	}
}

func TestExtractNativePaginationSettingsV1StylePaneFilterBitsAreNotInvalid(t *testing.T) {
	inner := `<w:stylePaneFormatFilter w:val="3F01" w:allStyles="1" w:customStyles="0"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>` + nativeMode15Compat()
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, inner)
	if paginationSettingsHasCode(settings, "INVALID_SETTINGS_STRUCTURE") {
		t.Fatalf("style pane filter bits must not be invalid structure: %#v", settings.Diagnostics)
	}
	if settings.Profile != "word-modern-default" || len(settings.Diagnostics) != 0 {
		t.Fatalf("style pane task-pane bits are chrome and must be admitted: %#v", settings)
	}
	nested := extractNativePaginationSettingsMarkup(t, wordMLTransitional, `<w:stylePaneFormatFilter w:val="3F01" w:allStyles="1"><w:foo/></w:stylePaneFormatFilter><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/>`+nativeMode15Compat())
	if !paginationSettingsHasCode(nested, "INVALID_SETTINGS_STRUCTURE") {
		t.Fatal("nested style pane filter markup must stay invalid")
	}
}

func TestExtractNativePaginationSettingsV1AdmitsImageCropNeutralSubset(t *testing.T) {
	for _, wordNS := range []string{wordMLTransitional, wordMLStrict} {
		t.Run(wordNS, func(t *testing.T) {
			settings := extractNativePaginationSettingsMarkup(t, wordNS, nativeImageCropNeutralExtras())
			if settings.Profile != "word-modern-default" || settings.CompatibilityMode == nil || *settings.CompatibilityMode != 15 || len(settings.Diagnostics) != 0 {
				t.Fatalf("ImageCrop-neutral subset with mode 15: %#v", settings)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1AdmitsTwoColHeaderNeutralSubset(t *testing.T) {
	for _, wordNS := range []string{wordMLTransitional, wordMLStrict} {
		t.Run(wordNS, func(t *testing.T) {
			settings := extractNativePaginationSettingsMarkup(t, wordNS, nativeTwoColHeaderNeutralExtras())
			if settings.Profile != "word-modern-default" || settings.CompatibilityMode == nil || *settings.CompatibilityMode != 15 || len(settings.Diagnostics) != 0 {
				t.Fatalf("2col-header-neutral subset with mode 15: %#v", settings)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1NeutralExtrasDoNotMaskOmittedMode(t *testing.T) {
	inner := strings.Replace(nativeImageCropNeutralExtras(), nativeMode15Compat(), `<w:compat/>`, 1)
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, inner)
	if settings.Profile != "unsupported" || settings.CompatibilityMode != nil || !paginationSettingsHasCode(settings, "COMPATIBILITY_SETTING_UNSUPPORTED") {
		t.Fatalf("omitted mode must stay refused: %#v", settings)
	}
	for _, code := range []string{"PAGINATION_SETTING_UNSUPPORTED", "UNKNOWN_SETTINGS_ELEMENT"} {
		if paginationSettingsHasCode(settings, code) {
			t.Fatalf("attested-neutral extras should not add %s when only mode is omitted: %#v", code, settings.Diagnostics)
		}
	}
}

func TestExtractNativePaginationSettingsV1NeutralExtrasDoNotAdmitMode14(t *testing.T) {
	inner := strings.Replace(nativeTwoColHeaderNeutralExtras(), `w:val="15"`, `w:val="14"`, 1)
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, inner)
	if settings.Profile != "unsupported" || settings.CompatibilityMode != nil || !paginationSettingsHasCode(settings, "COMPATIBILITY_SETTING_UNSUPPORTED") {
		t.Fatalf("mode 14 must stay refused: %#v", settings)
	}
	if paginationSettingsHasCode(settings, "PAGINATION_SETTING_UNSUPPORTED") || paginationSettingsHasCode(settings, "UNKNOWN_SETTINGS_ELEMENT") {
		t.Fatalf("mode 14 extras should not add unattested-element codes: %#v", settings.Diagnostics)
	}
}

func TestExtractNativePaginationSettingsV1RefusesLayoutAffectingCompatFlagsWithMode15(t *testing.T) {
	flags := `<w:compat>` +
		`<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>` +
		`<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>` +
		`<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>` +
		`<w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>` +
		`</w:compat>`
	inner := strings.Replace(nativeTwoColHeaderNeutralExtras(), nativeMode15Compat(), flags, 1)
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, inner)
	if settings.Profile != "unsupported" || !paginationSettingsHasCode(settings, "COMPATIBILITY_SETTING_UNSUPPORTED") {
		t.Fatalf("layout-affecting compat flags must stay refused: %#v", settings)
	}
	count := 0
	for _, diagnostic := range settings.Diagnostics {
		if diagnostic.Code == "COMPATIBILITY_SETTING_UNSUPPORTED" && strings.Contains(diagnostic.Path, "compatSetting") {
			count++
		}
	}
	if count != 3 {
		t.Fatalf("expected 3 extra-flag refusals, got %d: %#v", count, settings.Diagnostics)
	}
}

func TestExtractNativePaginationSettingsV1NeutralExtrasFailClosed(t *testing.T) {
	for _, test := range []struct {
		name, markup, code string
	}{
		{"east-asian theme language", `<w:themeFontLang w:val="ja-JP"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"arabic theme language", `<w:themeFontLang w:val="ar-SA"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"theme eastAsia language", `<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"empty shape defaults", `<w:shapeDefaults/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"shape fill defaults", `<w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout><o:shapedefaults v:ext="edit" spidmax="1"/></w:shapeDefaults>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"unknown decimal", `<w:decimalSymbol w:val="/"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"unknown list separator", `<w:listSeparator w:val=":"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"multi-char decimal", `<w:decimalSymbol w:val=".,"/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"incomplete math", `<m:mathPr><m:mathFont m:val="Cambria Math"/><m:wrapIndent m:val="1440"/></m:mathPr>`, "UNKNOWN_SETTINGS_ELEMENT"},
		{"non-default math wrap", strings.Replace(nativeAttestedMathPr(), `m:val="1440"`, `m:val="720"`, 1), "UNKNOWN_SETTINGS_ELEMENT"},
		{"smallFrac true", strings.Replace(nativeAttestedMathPr(), `<m:smallFrac m:val="0"/>`, `<m:smallFrac m:val="true"/>`, 1), "UNKNOWN_SETTINGS_ELEMENT"},
		{"smallFrac implicit on", strings.Replace(nativeAttestedMathPr(), `<m:smallFrac m:val="0"/>`, `<m:smallFrac/>`, 1), "UNKNOWN_SETTINGS_ELEMENT"},
		{"dispDef false", strings.Replace(nativeAttestedMathPr(), `<m:dispDef/>`, `<m:dispDef m:val="false"/>`, 1), "UNKNOWN_SETTINGS_ELEMENT"},
		{"unknown math child", `<m:mathPr><m:unknown m:val="1"/></m:mathPr>`, "UNKNOWN_SETTINGS_ELEMENT"},
		{"foreign markup", `<x:unknown xmlns:x="urn:foreign"/>`, "UNKNOWN_SETTINGS_ELEMENT"},
	} {
		t.Run(test.name, func(t *testing.T) {
			settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, test.markup+nativeMode15Compat())
			if settings.Profile != "unsupported" || !paginationSettingsHasCode(settings, test.code) {
				t.Fatalf("want %s, got %#v", test.code, settings)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1NeutralExtrasIndividualAdmission(t *testing.T) {
	for _, test := range []struct {
		name, markup string
	}{
		{"theme en-US", `<w:themeFontLang w:val="en-US"/>`},
		{"theme fr-FR x-none", `<w:themeFontLang w:val="fr-FR" w:eastAsia="x-none" w:bidi="x-none"/>`},
		{"decimal comma", `<w:decimalSymbol w:val=","/>`},
		{"decimal period", `<w:decimalSymbol w:val="."/>`},
		{"list comma", `<w:listSeparator w:val=","/>`},
		{"list semicolon", `<w:listSeparator w:val=";"/>`},
		{"shape defaults", nativeAttestedShapeDefaults()},
		{"math defaults", nativeAttestedMathPr()},
		{"math smallFrac off", strings.Replace(nativeAttestedMathPr(), `m:val="0"`, `m:val="off"`, 1)},
		{"math smallFrac false", strings.Replace(nativeAttestedMathPr(), `<m:smallFrac m:val="0"/>`, `<m:smallFrac m:val="false"/>`, 1)},
		{"math dispDef true", strings.Replace(nativeAttestedMathPr(), `<m:dispDef/>`, `<m:dispDef m:val="true"/>`, 1)},
	} {
		t.Run(test.name, func(t *testing.T) {
			settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, test.markup+nativeMode15Compat())
			if settings.Profile != "word-modern-default" || len(settings.Diagnostics) != 0 {
				t.Fatalf("attested extra refused: %#v", settings)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1RejectsSpoofedNeutralExtras(t *testing.T) {
	for _, markup := range []string{
		`<w:settings xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:foreign"><x:decimalSymbol w:val="."/>` + nativeMode15Compat() + `</w:settings>`,
		`<w:settings xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:foreign"><x:themeFontLang w:val="en-US"/>` + nativeMode15Compat() + `</w:settings>`,
		`<w:settings xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:foreign"><x:shapeDefaults/>` + nativeMode15Compat() + `</w:settings>`,
	} {
		if _, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(markup)))); err == nil || !strings.Contains(err.Error(), "namespace spoofing") {
			t.Fatalf("spoof error = %v for %s", err, markup)
		}
	}
}

// The Word/LibreOffice new-shape template carries the fill, style and colour
// history the next drawn shape inherits. It is outside the attested identity
// subset, so it must refuse — but as "not proven neutral" for the one
// w:shapeDefaults element, never as a structural invalidity of settings.xml,
// which is what keeps the whole package out of the approximate tier.
func TestExtractNativePaginationSettingsV1NewShapeTemplateIsNotStructurallyInvalid(t *testing.T) {
	markup := `<w:shapeDefaults><o:shapedefaults v:ext="edit" spidmax="8193" style="mso-height-percent:900" fillcolor="white"><v:fill color="white"/><o:colormru v:ext="edit" colors="#40a6be,#b4dce6"/></o:shapedefaults><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults>`
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, markup+nativeMode15Compat())
	if paginationSettingsHasCode(settings, "INVALID_SETTINGS_STRUCTURE") {
		t.Fatalf("new-shape defaults must not report structural invalidity: %#v", settings.Diagnostics)
	}
	count := 0
	for _, diagnostic := range settings.Diagnostics {
		if diagnostic.Path != "/w:settings[1]/w:shapeDefaults[1]" {
			continue
		}
		if diagnostic.Code != "PAGINATION_SETTING_UNSUPPORTED" {
			t.Fatalf("want PAGINATION_SETTING_UNSUPPORTED at w:shapeDefaults, got %s", diagnostic.Code)
		}
		count++
	}
	if count != 1 {
		t.Fatalf("expected one w:shapeDefaults refusal, got %d: %#v", count, settings.Diagnostics)
	}
	eligibility, err := ExtractNativeDocxApproximationEligibilityV1(nativeApproximationTestDOCX(t, markup+nativeMode15Compat()))
	if err != nil {
		t.Fatal(err)
	}
	if eligibility.Status != "eligible" {
		t.Fatalf("new-shape defaults must stay approximate-eligible, got %s: %#v", eligibility.Status, eligibility.Reasons)
	}
}

// Markup that is genuinely unrepresentable still reports INVALID_SETTINGS_STRUCTURE
// through the same element, so the quiet subset probe never hides a malformed part.
func TestExtractNativePaginationSettingsV1ShapeDefaultsTextStaysStructurallyInvalid(t *testing.T) {
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, `<w:shapeDefaults>shape</w:shapeDefaults>`+nativeMode15Compat())
	if !paginationSettingsHasCode(settings, "INVALID_SETTINGS_STRUCTURE") {
		t.Fatalf("non-whitespace settings text must stay structurally invalid: %#v", settings.Diagnostics)
	}
}

func nativeHdrShapeDefaults() string {
	return `<w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="2049"/></w:hdrShapeDefaults>`
}

// Each entry is one Word-emitted settings.xml child whose whole effect is
// outside the shaped body box, so admitting it cannot move a glyph, a line
// break, a margin or a page boundary. The per-setting arguments live beside the
// implementation in native_pagination_settings_neutral.go.
func TestExtractNativePaginationSettingsV1AdmitsLayoutNeutralChrome(t *testing.T) {
	for _, test := range []struct{ name, markup string }{
		{"page border excludes header", `<w:bordersDoNotSurroundHeader/>`},
		{"page border excludes footer", `<w:bordersDoNotSurroundFooter w:val="1"/>`},
		{"drawing grid origin", `<w:doNotUseMarginsForDrawingGridOrigin/>`},
		{"annotation date removal", `<w:removeDateAndTime w:val="true"/>`},
		{"legacy ui compatibility", `<w:uiCompat97To2003/>`},
		{"grid display interval", `<w:displayHorizontalDrawingGridEvery w:val="0"/><w:displayVerticalDrawingGridEvery w:val="2"/>`},
		{"grid spacing", `<w:drawingGridHorizontalSpacing w:val="120"/><w:drawingGridVerticalSpacing w:val="120"/>`},
		{"grid origin", `<w:drawingGridHorizontalOrigin w:val="0"/><w:drawingGridVerticalOrigin w:val="0"/>`},
		{"header shape defaults", nativeHdrShapeDefaults()},
		{"header shape defaults with layout", `<w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="2055"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="2"/></o:shapelayout></w:hdrShapeDefaults>`},
		{"proofing registration", `<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="64" w:dllVersion="131078" w:nlCheck="1" w:checkStyle="1"/>`},
		{"repeated proofing registrations", `<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="64" w:dllVersion="6" w:nlCheck="1" w:checkStyle="1"/><w:activeWritingStyle w:appName="MSWord" w:lang="fr-FR" w:vendorID="64" w:dllVersion="6" w:nlCheck="1" w:checkStyle="0"/>`},
		{"style pane filter bits", `<w:stylePaneFormatFilter w:val="3804" w:allStyles="0" w:customStyles="0" w:latentStyles="1" w:stylesInUse="0" w:headingStyles="0" w:numberingStyles="0" w:tableStyles="0" w:directFormattingOnRuns="0" w:directFormattingOnParagraphs="0" w:directFormattingOnNumbering="0" w:directFormattingOnTables="1" w:clearFormatting="1" w:top3HeadingStyles="1" w:visibleStyles="0" w:alternateStyleNames="0"/>`},
		{"word 2010 document identity", `<w14:docId w14:val="03CBB905"/>`},
		{"word 2013 document identity", `<w15:docId w15:val="{003352EF-7FFF-409D-8CBC-B91DDDD484AE}"/>`},
		{"default image dpi", `<w14:defaultImageDpi w14:val="330"/>`},
		{"disabled image compression", `<w14:defaultImageDpi w14:val="0"/>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			for _, wordNS := range []string{wordMLTransitional, wordMLStrict} {
				settings := extractNativePaginationSettingsMarkup(t, wordNS, test.markup+nativeMode15Compat())
				if settings.Profile != "word-modern-default" || len(settings.Diagnostics) != 0 {
					t.Fatalf("layout-neutral chrome refused: %#v", settings)
				}
			}
		})
	}
}

// The admitted arm is a bounded subset, not a waiver: markup outside its exact
// shape, and every setting that can still move layout, keeps refusing.
func TestExtractNativePaginationSettingsV1NeutralChromeFailsClosed(t *testing.T) {
	for _, test := range []struct{ name, markup, code string }{
		{"non-lexical border switch", `<w:bordersDoNotSurroundHeader w:val="maybe"/>`, "INVALID_SETTINGS_ON_OFF"},
		{"non-lexical date removal", `<w:removeDateAndTime w:val="yes please"/>`, "INVALID_SETTINGS_ON_OFF"},
		{"grid spacing without value", `<w:drawingGridHorizontalSpacing/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"non-numeric grid spacing", `<w:drawingGridVerticalSpacing w:val="wide"/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"negative grid origin", `<w:drawingGridHorizontalOrigin w:val="-20"/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"header shape paint defaults", `<w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="8193" fillcolor="white"/></w:hdrShapeDefaults>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"empty header shape defaults", `<w:hdrShapeDefaults/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"header shape zero identity", `<w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="0"/></w:hdrShapeDefaults>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"writing style without language", `<w:activeWritingStyle w:appName="MSWord" w:vendorID="64"/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"writing style with nested markup", `<w:activeWritingStyle w:lang="en-US"><w:foo/></w:activeWritingStyle>`, "INVALID_SETTINGS_STRUCTURE"},
		{"style pane filter nested markup", `<w:stylePaneFormatFilter w:val="3F01" w:allStyles="1"><w:foo/></w:stylePaneFormatFilter>`, "INVALID_SETTINGS_STRUCTURE"},
		{"style pane filter unknown bit", `<w:stylePaneFormatFilter w:val="3F01" w:allStyles="1" w:unknownBit="1"/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"document identity without value", `<w15:docId/>`, "UNKNOWN_SETTINGS_ELEMENT"},
		{"document identity with nested markup", `<w15:docId w15:val="{003352EF-7FFF-409D-8CBC-B91DDDD484AE}"><w15:foo/></w15:docId>`, "INVALID_SETTINGS_STRUCTURE"},
		{"non-numeric image dpi", `<w14:defaultImageDpi w14:val="high"/>`, "INVALID_SETTINGS_STRUCTURE"},
		{"unbounded image dpi", `<w14:defaultImageDpi w14:val="99000000"/>`, "INVALID_SETTINGS_STRUCTURE"},
		// Settings that can still move layout are untouched by this admission.
		{"mirror margins", `<w:mirrorMargins/>`, "MIRROR_MARGINS_UNSUPPORTED"},
		{"gutter at top", `<w:gutterAtTop/>`, "GUTTER_AT_TOP_UNSUPPORTED"},
		{"active hyphenation", `<w:autoHyphenation/><w:hyphenationZone w:val="360"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"punctuation compression", `<w:characterSpacingControl w:val="compressPunctuation"/>`, "CHARACTER_SPACING_CONTROL_UNSUPPORTED"},
		{"east asian theme language", `<w:themeFontLang w:val="ja-JP"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"document grid snapping", `<w:doNotSnapToGridInCell/>`, "PAGINATION_SETTING_UNSUPPORTED"},
		{"update fields on open", `<w:updateFields w:val="true"/>`, "PAGINATION_SETTING_UNSUPPORTED"},
	} {
		t.Run(test.name, func(t *testing.T) {
			settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, test.markup+nativeMode15Compat())
			if settings.Profile != "unsupported" || !paginationSettingsHasCode(settings, test.code) {
				t.Fatalf("want %s, got %#v", test.code, settings)
			}
		})
	}
}

// A repeated w:activeWritingStyle is schema-legal; every other repeat is not.
func TestExtractNativePaginationSettingsV1DuplicateExemptionIsNarrow(t *testing.T) {
	repeated := `<w:removeDateAndTime/><w:removeDateAndTime/>`
	settings := extractNativePaginationSettingsMarkup(t, wordMLTransitional, repeated+nativeMode15Compat())
	if settings.Profile != "unsupported" || !paginationSettingsHasCode(settings, "DUPLICATE_SETTINGS_PROPERTY") {
		t.Fatalf("only activeWritingStyle is schema-repeatable: %#v", settings)
	}
}

// A foreign-namespace element borrowing an admitted local name is still a spoof.
func TestExtractNativePaginationSettingsV1RejectsSpoofedNeutralChrome(t *testing.T) {
	for _, local := range []string{"hdrShapeDefaults", "activeWritingStyle", "uiCompat97To2003", "drawingGridHorizontalSpacing", "stylePaneFormatFilter"} {
		markup := `<w:settings xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:foreign"><x:` + local + ` w:val="1"/>` + nativeMode15Compat() + `</w:settings>`
		if _, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(markup)))); err == nil || !strings.Contains(err.Error(), "namespace spoofing") {
			t.Fatalf("spoof error = %v for %s", err, local)
		}
	}
}
