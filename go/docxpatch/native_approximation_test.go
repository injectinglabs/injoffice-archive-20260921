package docxpatch

import (
	"bytes"
	"reflect"
	"strings"
	"testing"
)

func TestNativeApproximationEligibilityRejectsAmbiguousLegacySettings(t *testing.T) {
	mode := func(value string) string {
		return `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="` + value + `"/>`
	}
	flag := func(name string) string {
		return `<w:compatSetting w:name="` + name + `" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`
	}
	extraFlags := flag("overrideTableStyleFontSizeAndJustification") + flag("enableOpenTypeFeatures") + flag("doNotFlipMirrorIndents") + flag("differentiateMultirowTableHeaders")
	for _, strict := range []bool{false, true} {
		for _, test := range []struct {
			name, markup string
			eligible     bool
			mode         int
		}{
			{"omitted", "", true, 12}, {"empty compat", "<w:compat/>", true, 12},
			{"mode12", "<w:compat>" + mode("12") + "</w:compat>", true, 12}, {"mode14", "<w:compat>" + mode("14") + "</w:compat>", true, 14},
			{"modern", "<w:compat>" + mode("15") + "</w:compat>", false, 0},
			{"mode15 extras", "<w:compat>" + mode("15") + extraFlags + "</w:compat>", true, 15},
			{"mode14 four flags", "<w:compat>" + mode("14") + extraFlags + "</w:compat>", true, 14},
			{"duplicate mode", "<w:compat>" + mode("12") + mode("12") + "</w:compat>", false, 0},
			{"duplicate compat", "<w:compat/><w:compat/>", false, 0},
			{"invalid number", "<w:compat>" + mode("oops") + "</w:compat>", false, 0},
			{"noncanonical number", "<w:compat>" + mode("012") + "</w:compat>", false, 0},
			{"unknown flag", "<w:compat><w:useWord2002TableStyleRules/></w:compat>", false, 0},
			{"foreign mode", `<w:compat><x:compatSetting xmlns:x="urn:foreign" w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="12"/></w:compat>`, false, 0},
			{"enabled hyphenation", "<w:autoHyphenation/>", true, 12},
			{"unknown attribute", `<w:compat w:unknown="1"/>`, false, 0},
			{"applyBreakingRules mode14", "<w:compat><w:applyBreakingRules/>" + mode("14") + flag("overrideTableStyleFontSizeAndJustification") + flag("enableOpenTypeFeatures") + flag("doNotFlipMirrorIndents") + "</w:compat>", true, 14},
			{"duplicate writing style", `<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="64" w:dllVersion="1" w:checkStyle="1"/><w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="8" w:dllVersion="1" w:checkStyle="1"/><w:compat>` + mode("14") + extraFlags + "</w:compat>", true, 14},
		} {
			t.Run(test.name+map[bool]string{false: " transitional", true: " strict"}[strict], func(t *testing.T) {
				parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordMLTransitional + `">` + test.markup + `</w:settings>`)
				if strict {
					for key, value := range parts {
						parts[key] = strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := append([]byte(nil), data...)
				settings, err := ExtractNativePaginationSettingsV1(data)
				if test.name == "foreign mode" {
					if err == nil || !strings.Contains(err.Error(), "namespace spoofing") {
						t.Fatalf("foreign settings must hard-refuse: %v", err)
					}
					if _, approximationErr := ExtractNativeDocxApproximationEligibilityV1(data); approximationErr == nil {
						t.Fatal("approximation accepted foreign settings")
					}
					return
				}
				if err != nil {
					t.Fatal(err)
				}
				eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if (eligibility.Status == "eligible") != test.eligible {
					t.Fatalf("unexpected eligibility %#v", eligibility)
				}
				if test.eligible && (eligibility.LegacyCompatibilityMode == nil || *eligibility.LegacyCompatibilityMode != test.mode || settings.Profile != "unsupported") {
					t.Fatalf("lost original refusal or exact mode: %#v %#v", settings, eligibility)
				}
				if test.name == "mode15 extras" && (settings.CompatibilityMode == nil || *settings.CompatibilityMode != 15) {
					t.Fatalf("mode 15 extras must keep attested strict mode 15: %#v", settings)
				}
				if strings.Contains(test.name, "four flags") || test.name == "mode15 extras" {
					if len(eligibility.ApproximatedSettings) != 3 {
						t.Fatalf("typed extra flags must be facts; differentiateMultirowTableHeaders is uncovered: %#v", eligibility.ApproximatedSettings)
					}
				}
				after, err := ExtractNativePaginationSettingsV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if !bytes.Equal(before, data) || !reflect.DeepEqual(settings, after) {
					t.Fatal("approximation changed source/strict attestation")
				}
			})
		}
	}
}
