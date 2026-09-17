package docxpatch

import (
	"strings"
	"testing"
)

// The exact chain Word follows for an East-Asian theme slot, built as a package:
//
//	docDefaults w:eastAsiaTheme -> theme <a:ea typeface=""/> -> settings
//	w:themeFontLang/@w:eastAsia -> script tag -> <a:font script=.../>
func eastAsianThemeParts(docDefaults, themeEA, scriptRows, settings, body string) map[string]string {
	parts := map[string]string{
		"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
			`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
			`<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
			`<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`,
		"_rels/.rels":       `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml": `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` + body + `<w:sectPr/></w:body></w:document>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `">` +
			`<Relationship Id="styles" Type="` + relBaseTransitional + `styles" Target="styles.xml"/>` +
			`<Relationship Id="theme" Type="` + relBaseTransitional + `theme" Target="theme/theme1.xml"/>` +
			`<Relationship Id="settings" Type="` + relBaseTransitional + `settings" Target="settings.xml"/></Relationships>`,
		"word/styles.xml":       `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` + docDefaults + `<w:sz w:val="24"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`,
		"word/settings.xml":     `<w:settings xmlns:w="` + wordMLTransitional + `">` + settings + `</w:settings>`,
		"word/theme/theme1.xml": `<a:theme xmlns:a="` + drawingMLTransitional + `" name="Office"><a:themeElements><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/>` + themeEA + `<a:cs typeface=""/>` + scriptRows + `</a:minorFont></a:fontScheme></a:themeElements></a:theme>`,
	}
	return parts
}

const eastAsianOfficeScriptRows = `<a:font script="Jpan" typeface="ＭＳ 明朝"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="宋体"/><a:font script="Hant" typeface="新細明體"/>`

func resolveEastAsianTheme(t *testing.T, parts map[string]string) *NativeResolvedLayoutInputV1 {
	t.Helper()
	layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	return layout
}

func TestNativeEastAsianThemeSlotResolvesThroughThemeFontLang(t *testing.T) {
	parts := eastAsianThemeParts(
		`<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi" w:cs="Times New Roman"/><w:szCs w:val="24"/><w:lang w:val="en-US" w:eastAsia="zh-TW" w:bidi="ar-SA"/>`,
		`<a:ea typeface=""/>`, eastAsianOfficeScriptRows,
		`<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`,
		`<w:p><w:r><w:t>Latin</w:t></w:r></w:p><w:p><w:r><w:t>甲.</w:t></w:r></w:p>`)
	layout := resolveEastAsianTheme(t, parts)
	if layout.SourceParts.SettingsPart == nil || *layout.SourceParts.SettingsPart != "word/settings.xml" {
		t.Fatalf("the settings part that selected the font was not attested: %#v", layout.SourceParts)
	}
	latin, eastAsian := layout.Runs[0].Properties, layout.Runs[1].Properties
	if latin.FontFamily == nil || *latin.FontFamily != "Calibri" || latin.EastAsiaFontFamily != nil || latin.EastAsiaLanguage != nil {
		t.Fatalf("latin-only text asked for an East-Asian face: %#v", latin)
	}
	if eastAsian.FontFamily == nil || *eastAsian.FontFamily != "Calibri" {
		t.Fatalf("the East-Asian slot displaced the latin slot: %#v", eastAsian)
	}
	if eastAsian.EastAsiaFontFamily == nil || *eastAsian.EastAsiaFontFamily != "新細明體" {
		t.Fatalf("minorEastAsia did not resolve through themeFontLang zh-TW: %#v", eastAsian)
	}
	if eastAsian.EastAsiaLanguage == nil || *eastAsian.EastAsiaLanguage != "zh-TW" {
		t.Fatalf("the East-Asian language was not carried with its slot: %#v", eastAsian)
	}
	for _, code := range []string{"SCRIPT_FONT_PRESERVED", "SCRIPT_LANGUAGE_PRESERVED", "COMPLEX_SCRIPT_SIZE_PRESERVED", "FONT_HINT_PRESERVED"} {
		if hasResolutionDiagnostic(layout, code) {
			t.Fatalf("a resolved East-Asian slot still deferred %s: %#v", code, layout.Diagnostics)
		}
	}
}

func TestNativeEastAsianThemeSlotRefusesOutsideTheExactChain(t *testing.T) {
	for _, test := range []struct {
		name, themeEA, scriptRows, settings string
	}{
		{"no settings themeFontLang", `<a:ea typeface=""/>`, eastAsianOfficeScriptRows, ``},
		{"themeFontLang states no east-asian language", `<a:ea typeface=""/>`, eastAsianOfficeScriptRows, `<w:themeFontLang w:val="en-US"/>`},
		{"east-asian language outside the bounded map", `<a:ea typeface=""/>`, eastAsianOfficeScriptRows, `<w:themeFontLang w:val="en-US" w:eastAsia="th-TH"/>`},
		{"script table has no row for the language", `<a:ea typeface=""/>`, `<a:font script="Hans" typeface="宋体"/>`, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`},
		{"duplicated script row has no single answer", `<a:ea typeface=""/>`, eastAsianOfficeScriptRows + `<a:font script="Hant" typeface="細明體"/>`, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`},
		{"script row states an empty typeface", `<a:ea typeface=""/>`, `<a:font script="Hant" typeface=""/>`, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`},
		{"duplicated themeFontLang has no single answer", `<a:ea typeface=""/>`, eastAsianOfficeScriptRows, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/><w:themeFontLang w:val="en-US" w:eastAsia="zh-CN"/>`},
		{"theme states no east-asian slot at all", ``, eastAsianOfficeScriptRows, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			layout := resolveEastAsianTheme(t, eastAsianThemeParts(
				`<w:rFonts w:asciiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:hAnsiTheme="minorHAnsi"/>`,
				test.themeEA, test.scriptRows, test.settings,
				`<w:p><w:r><w:t>甲.</w:t></w:r></w:p>`))
			if !hasResolutionDiagnostic(layout, "SCRIPT_FONT_PRESERVED") {
				t.Fatalf("an unresolved East-Asian slot stopped refusing: %#v", layout.Diagnostics)
			}
			if layout.Runs[0].Properties.EastAsiaFontFamily != nil {
				t.Fatalf("an unresolved East-Asian slot invented a face: %#v", layout.Runs[0].Properties)
			}
		})
	}
}

func TestNativeEastAsianSlotStatedDirectlyIsUnaffectedByTheme(t *testing.T) {
	// A package that names its own East-Asian family never consults the theme
	// script table, so themeFontLang cannot move it.
	for _, settings := range []string{``, `<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>`} {
		layout := resolveEastAsianTheme(t, eastAsianThemeParts(
			`<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="SimSun"/>`,
			`<a:ea typeface=""/>`, eastAsianOfficeScriptRows, settings,
			`<w:p><w:r><w:t>甲.</w:t></w:r></w:p>`))
		run := layout.Runs[0].Properties
		if run.EastAsiaFontFamily == nil || *run.EastAsiaFontFamily != "SimSun" || run.FontFamily == nil || *run.FontFamily != "Arial" {
			t.Fatalf("a directly stated East-Asian family was not preserved: %#v", run)
		}
		if hasResolutionDiagnostic(layout, "SCRIPT_FONT_PRESERVED") {
			t.Fatalf("a directly stated East-Asian family still refused: %#v", layout.Diagnostics)
		}
	}
	// A non-empty <a:ea> answers the slot itself and outranks the script table.
	layout := resolveEastAsianTheme(t, eastAsianThemeParts(
		`<w:rFonts w:eastAsiaTheme="minorEastAsia"/>`,
		`<a:ea typeface="ＭＳ 明朝"/>`, eastAsianOfficeScriptRows, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`,
		`<w:p><w:r><w:t>甲.</w:t></w:r></w:p>`))
	if run := layout.Runs[0].Properties; run.EastAsiaFontFamily == nil || *run.EastAsiaFontFamily != "ＭＳ 明朝" {
		t.Fatalf("an authored <a:ea> typeface was overruled by the script table: %#v", run)
	}
}

func TestNativeEastAsianTextWithoutAnySlotStillRefuses(t *testing.T) {
	// No w:eastAsia, no w:eastAsiaTheme: the ascii face is not a guess for 甲.
	layout := resolveEastAsianTheme(t, eastAsianThemeParts(
		`<w:rFonts w:ascii="Arial" w:hAnsi="Arial"/>`,
		`<a:ea typeface=""/>`, eastAsianOfficeScriptRows, `<w:themeFontLang w:val="en-US" w:eastAsia="zh-TW"/>`,
		`<w:p><w:r><w:t>Latin</w:t></w:r></w:p><w:p><w:r><w:t>甲.</w:t></w:r></w:p>`))
	if !hasResolutionDiagnostic(layout, "SCRIPT_FONT_PRESERVED") {
		t.Fatalf("East-Asian text with no slot at all was painted in the ascii face: %#v", layout.Diagnostics)
	}
	if layout.Runs[0].Properties.FontFamily == nil || *layout.Runs[0].Properties.FontFamily != "Arial" {
		t.Fatalf("the latin paragraph lost its own font: %#v", layout.Runs[0].Properties)
	}
}

func TestNativeEastAsianThemeScriptMap(t *testing.T) {
	for tag, want := range map[string]string{
		"zh-TW": "Hant", "zh-HK": "Hant", "zh-MO": "Hant", "zh-Hant": "Hant",
		"zh-CN": "Hans", "zh-SG": "Hans", "zh-Hans": "Hans", "zh": "Hans",
		"ja-JP": "Jpan", "ja": "Jpan", "ko-KR": "Hang", "ko": "Hang",
	} {
		if got := nativeEastAsianThemeScript(tag); got != want {
			t.Fatalf("%s resolved to %q, want %q", tag, got, want)
		}
	}
	for _, tag := range []string{"", "en-US", "ar-SA", "th-TH", "x-none", "zhh", strings.Repeat("z", 40)} {
		if got := nativeEastAsianThemeScript(tag); got != "" {
			t.Fatalf("%q was guessed as script %q", tag, got)
		}
	}
}
