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
	settingsXML := `<w:settings xmlns:w="` + wordNS + `" xmlns:m="` + nativeMathNamespace + `" xmlns:o="` + nativeOfficeNamespace + `" xmlns:v="` + nativeVMLNamespace + `">` + inner + `</w:settings>`
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
	if !paginationSettingsHasCode(settings, "PAGINATION_SETTING_UNSUPPORTED") {
		t.Fatal("style pane filter must remain pagination-unsupported")
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
