package docxpatch

import (
	"reflect"
	"strings"
	"testing"
)

func TestNativeApproximationKnownSettingsRetainStrictRefusal(t *testing.T) {
	math := `<m:mathPr xmlns:m="` + nativeMathNamespace + `"><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>`
	shape := `<w:shapeDefaults xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults>`
	for _, test := range []struct {
		name, markup string
		eligible     bool
	}{
		{"language", `<w:themeFontLang w:val="en-US" w:eastAsia="x-none"/>`, true},
		{"locale", `<w:decimalSymbol w:val=","/><w:listSeparator w:val=";"/>`, true},
		{"math", math, true}, {"shape", shape, true},
		{"combined", math + shape + `<w:themeFontLang w:val="fr-FR"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>`, true},
		{"compatibility flags", `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`, true},
		{"mode 15 extras", `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="differentiateMultirowTableHeaders" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat><w:themeFontLang w:val="hu-HU"/>`, true},
		{"malformed language", `<w:themeFontLang w:val="en_US"/>`, true},
		{"unknown language", `<w:themeFontLang w:val="ar-SA"/>`, true},
		{"duplicate locale", `<w:decimalSymbol w:val="."/><w:decimalSymbol w:val="."/>`, false},
		{"duplicate writing style", `<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="64" w:dllVersion="1" w:checkStyle="1"/><w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="8" w:dllVersion="1" w:checkStyle="1"/><w:compat><w:applyBreakingRules/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`, true},
		{"unknown decimal", `<w:decimalSymbol w:val="unknown"/>`, false},
		{"nested locale", `<w:decimalSymbol w:val="."><w:foo/></w:decimalSymbol>`, false},
		{"unknown math", strings.Replace(math, `m:val="1440"`, `m:val="999"`, 1), true},
		{"duplicate math", strings.Replace(math, `</m:mathPr>`, `<m:intLim m:val="subSup"/></m:mathPr>`, 1), true},
		{"unknown math attribute", strings.Replace(math, `<m:mathPr `, `<m:mathPr foo="1" `, 1), false},
		{"malformed shape id", strings.Replace(shape, `spidmax="1026"`, `spidmax="01026"`, 1), true},
		{"shape content", strings.Replace(shape, `data="1"/>`, `data="1"><o:unknown/></o:idmap>`, 1), false},
		{"unknown flag value", `<w:compat><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="oops"/></w:compat>`, false},
		{"flag first without mode", `<w:compat><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`, true},
		{"flag before mode", `<w:compat><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`, true},
		{"empty script slots", `<w:themeFontLang w:val="en-CA" w:eastAsia="" w:bidi=""/>`, true},
		{"east asian theme language", `<w:themeFontLang w:val="en-US" w:eastAsia="ja-JP"/>`, true},
		{"word 2013 flags", `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/><w:compatSetting w:name="allowHyphenationAtTrackBottom" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="allowTextAfterFloatingTableBreak" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/></w:compat>`, true},
		{"unknown compat setting", `<w:compat><w:compatSetting w:name="notAKnownFlag" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`">`+test.markup+`</w:settings>`)))
			strict, err := ExtractNativePaginationSettingsV1(data)
			if err != nil {
				t.Fatal(err)
			}
			approx, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if (approx.Status == "eligible") != test.eligible {
				t.Fatalf("wrong eligibility %#v", approx)
			}
			if test.name == "mode 15 extras" && (approx.LegacyCompatibilityMode == nil || *approx.LegacyCompatibilityMode != 15 || strict.CompatibilityMode == nil || *strict.CompatibilityMode != 15) {
				t.Fatalf("mode 15 extras must name current-layout mode 15 while strict stays unsupported: %#v %#v", strict, approx)
			}
			if test.eligible && len(approx.Reasons) == 0 {
				t.Fatalf("eligible approximation must retain a reason: %#v", approx)
			}
			for _, fact := range approx.ApproximatedSettings {
				matched := false
				for _, diagnostic := range strict.Diagnostics {
					if diagnostic.Path == fact.Path {
						matched = true
						break
					}
				}
				if !matched {
					t.Fatalf("approximated fact %s %s has no matching strict diagnostic; join would 422", fact.Kind, fact.Path)
				}
			}
			after, err := ExtractNativePaginationSettingsV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if strict.Profile != "unsupported" || !reflect.DeepEqual(strict, after) {
				t.Fatal("strict settings changed")
			}
		})
	}
}

func TestNativeApproximationEligibleWithStylePaneFilterBits(t *testing.T) {
	markup := `<w:stylePaneFormatFilter w:val="3F01" w:allStyles="1"/><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`
	data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`">`+markup+`</w:settings>`)))
	approx, err := ExtractNativeDocxApproximationEligibilityV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if approx.Status != "eligible" || approx.LegacyCompatibilityMode == nil || *approx.LegacyCompatibilityMode != 14 {
		t.Fatalf("style pane filter bits plus mode 14 must stay approximately eligible: %#v", approx)
	}
}

func TestNativeApproximationAllowsUncoveredPaginationExtras(t *testing.T) {
	data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`"><w:autoHyphenation/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat></w:settings>`)))
	approx, err := ExtractNativeDocxApproximationEligibilityV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if approx.Status != "eligible" || approx.LegacyCompatibilityMode == nil || *approx.LegacyCompatibilityMode != 14 {
		t.Fatalf("autoHyphenation plus mode 14 must stay approximately eligible: %#v", approx)
	}
}

func TestNativeApproximationOmitsSilentNeutralExtrasFromFacts(t *testing.T) {
	math := `<m:mathPr xmlns:m="` + nativeMathNamespace + `"><m:mathFont m:val="Cambria Math"/><m:brkBin m:val="before"/><m:brkBinSub m:val="--"/><m:smallFrac m:val="0"/><m:dispDef/><m:lMargin m:val="0"/><m:rMargin m:val="0"/><m:defJc m:val="centerGroup"/><m:wrapIndent m:val="1440"/><m:intLim m:val="subSup"/><m:naryLim m:val="undOvr"/></m:mathPr>`
	shape := `<w:shapeDefaults xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml"><o:shapedefaults v:ext="edit" spidmax="1026"/><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout></w:shapeDefaults>`
	markup := math + shape + `<w:themeFontLang w:val="en-US"/><w:decimalSymbol w:val="."/><w:listSeparator w:val=","/>` +
		`<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`
	data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`">`+markup+`</w:settings>`)))
	strict, err := ExtractNativePaginationSettingsV1(data)
	if err != nil {
		t.Fatal(err)
	}
	approx, err := ExtractNativeDocxApproximationEligibilityV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if approx.Status != "eligible" || approx.LegacyCompatibilityMode == nil || *approx.LegacyCompatibilityMode != 14 {
		t.Fatalf("mode 14 plus Word extras must stay approximately eligible: %#v", approx)
	}
	if len(approx.ApproximatedSettings) != 1 || approx.ApproximatedSettings[0].Kind != "enableOpenTypeFeatures" {
		t.Fatalf("silent-neutral extras must not be facts; extra compat flags must: %#v", approx.ApproximatedSettings)
	}
	for _, fact := range approx.ApproximatedSettings {
		matched := false
		for _, diagnostic := range strict.Diagnostics {
			if diagnostic.Path == fact.Path {
				matched = true
				break
			}
		}
		if !matched {
			t.Fatalf("fact %s missing diagnostic", fact.Path)
		}
	}
}
