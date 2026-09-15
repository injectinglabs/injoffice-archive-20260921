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
			{"duplicate mode", "<w:compat>" + mode("12") + mode("12") + "</w:compat>", true, 12},
			{"disagreeing modes", "<w:compat>" + mode("12") + mode("14") + "</w:compat>", false, 0},
			{"mode after flag", "<w:compat>" + flag("enableOpenTypeFeatures") + mode("14") + "</w:compat>", true, 14},
			{"duplicate compat", "<w:compat/><w:compat/>", false, 0},
			{"invalid number", "<w:compat>" + mode("oops") + "</w:compat>", false, 0},
			{"noncanonical number", "<w:compat>" + mode("012") + "</w:compat>", false, 0},
			{"legacy option", "<w:compat><w:useWord2002TableStyleRules/></w:compat>", true, 12},
			{"repeated legacy option", "<w:compat><w:useFELayout/><w:useFELayout/></w:compat>", false, 0},
			{"unknown compat markup", "<w:compat><w:notACompatOption/></w:compat>", false, 0},
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
					if len(eligibility.ApproximatedSettings) != 4 {
						t.Fatalf("all four typed extra flags must be facts: %#v", eligibility.ApproximatedSettings)
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

func nativeApproximationTestDOCX(t *testing.T, markup string) []byte {
	t.Helper()
	return buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`" xmlns:r="`+relNSTransitional+`" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:v="urn:schemas-microsoft-com:vml">`+markup+`</w:settings>`)))
}

func nativeApproximationFact(facts []NativeDocxApproximatedSettingV1, kind string) *NativeDocxApproximatedSettingV1 {
	for index := range facts {
		if facts[index].Kind == kind {
			return &facts[index]
		}
	}
	return nil
}

// Every admitted extra is a typed fact whose path joins one strict diagnostic
// and whose disclosed reason is present; strict refusal and source bytes stay put.
func requireNativeApproximationDisclosure(t *testing.T, data []byte) (*NativePaginationSettingsV1, *NativeDocxApproximationEligibilityV1) {
	t.Helper()
	before := append([]byte(nil), data...)
	strict, err := ExtractNativePaginationSettingsV1(data)
	if err != nil {
		t.Fatal(err)
	}
	approx, err := ExtractNativeDocxApproximationEligibilityV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if strict.Profile != "unsupported" || !bytes.Equal(before, data) {
		t.Fatalf("strict refusal or source bytes changed: %#v", strict)
	}
	if len(approx.ApproximatedSettings) > 8 {
		t.Fatalf("facts exceed the decoder bound: %#v", approx.ApproximatedSettings)
	}
	kinds, paths := map[string]bool{}, map[string]bool{}
	for _, fact := range approx.ApproximatedSettings {
		if kinds[fact.Kind] || paths[fact.Path] {
			t.Fatalf("fact kinds and paths must be unique: %#v", approx.ApproximatedSettings)
		}
		kinds[fact.Kind], paths[fact.Path] = true, true
		joined := false
		for _, diagnostic := range strict.Diagnostics {
			joined = joined || diagnostic.Path == fact.Path
		}
		if !joined {
			t.Fatalf("fact %s at %s has no strict diagnostic to join", fact.Kind, fact.Path)
		}
		disclosed := false
		for _, reason := range approx.Reasons {
			disclosed = disclosed || reason == nativeApproximationSettingReason(fact)
		}
		if !disclosed {
			t.Fatalf("fact %s at %s has no disclosed reason", fact.Kind, fact.Path)
		}
	}
	return strict, approx
}

func TestNativeApproximationRecordsActiveHyphenationAsNotPerformed(t *testing.T) {
	data := nativeApproximationTestDOCX(t, `<w:autoHyphenation w:val="true"/><w:hyphenationZone w:val="360"/><w:consecutiveHyphenLimit w:val="2"/><w:doNotHyphenateCaps/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>`)
	strict, approx := requireNativeApproximationDisclosure(t, data)
	if approx.Status != "eligible" || approx.LegacyCompatibilityMode == nil || *approx.LegacyCompatibilityMode != 15 || strict.CompatibilityMode == nil {
		t.Fatalf("active hyphenation with attested mode 15 must be approximately eligible: %#v", approx)
	}
	fact := nativeApproximationFact(approx.ApproximatedSettings, "autoHyphenation")
	if fact == nil || fact.Path != "/w:settings[1]/w:autoHyphenation[1]" || !reflect.DeepEqual(fact.Values, map[string]string{"val": "true", "hyphenationZone": "360", "consecutiveHyphenLimit": "2"}) {
		t.Fatalf("hyphenation fact must retain the switch and its active options: %#v", approx.ApproximatedSettings)
	}
	if !strings.Contains(nativeApproximationSettingReason(*fact), "automatic hyphenation is not performed") {
		t.Fatalf("hyphenation disclosure must say it is not performed: %q", nativeApproximationSettingReason(*fact))
	}
	inactive := nativeApproximationTestDOCX(t, `<w:autoHyphenation w:val="0"/><w:hyphenationZone w:val="360"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`)
	_, approx = requireNativeApproximationDisclosure(t, inactive)
	if approx.Status != "eligible" || nativeApproximationFact(approx.ApproximatedSettings, "autoHyphenation") != nil {
		t.Fatalf("disabled hyphenation is strict-neutral and must not become a fact: %#v", approx)
	}
}

func TestNativeApproximationRecordsAuthoringOnlySettingsAsOneGroupedFact(t *testing.T) {
	markup := `<w:removePersonalInformation/><w:removeDateAndTime/><w:displayBackgroundShape/>` +
		`<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="64" w:dllVersion="131078" w:nlCheck="1" w:checkStyle="1"/>` +
		`<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="8" w:dllVersion="513" w:checkStyle="1"/>` +
		`<w:attachedTemplate r:id="rId1"/><w:linkStyles/><w:stylePaneFormatFilter w:val="3804" w:allStyles="0" w:latentStyles="1"/>` +
		`<w:displayHorizontalDrawingGridEvery w:val="0"/><w:displayVerticalDrawingGridEvery w:val="0"/><w:doNotUseMarginsForDrawingGridOrigin/>` +
		`<w:noPunctuationKerning/><w:hdrShapeDefaults><o:shapedefaults v:ext="edit" spidmax="18433"/></w:hdrShapeDefaults>` +
		`<w:compat><w:applyBreakingRules/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/><w:compatSetting w:name="doNotFlipMirrorIndents" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`
	_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, markup))
	if approx.Status != "eligible" || approx.LegacyCompatibilityMode == nil || *approx.LegacyCompatibilityMode != 14 {
		t.Fatalf("Word 2010 authoring extras must be approximately eligible: %#v", approx)
	}
	group := nativeApproximationFact(approx.ApproximatedSettings, "authoringSettings")
	if group == nil || group.Path != "/w:settings[1]/w:removePersonalInformation[1]" || len(group.Values) != 12 {
		t.Fatalf("authoring settings must be one grouped fact anchored at the first diagnosed member: %#v", approx.ApproximatedSettings)
	}
	for path, expected := range map[string]string{
		"/w:settings[1]/w:attachedTemplate[1]":                  `r:id="rId1"`,
		"/w:settings[1]/w:activeWritingStyle[1]":                `w:appName="MSWord" w:checkStyle="1" w:dllVersion="131078" w:lang="en-US" w:nlCheck="1" w:vendorID="64"`,
		"/w:settings[1]/w:hdrShapeDefaults[1]":                  ``,
		"/w:settings[1]/w:displayHorizontalDrawingGridEvery[1]": `w:val="0"`,
	} {
		if group.Values[path] != expected {
			t.Fatalf("member %s retained %q, want %q", path, group.Values[path], expected)
		}
	}
	if _, present := group.Values["/w:settings[1]/w:activeWritingStyle[2]"]; present {
		t.Fatal("the duplicate writing style belongs to the duplicate group, not the authoring group")
	}
	duplicates := nativeApproximationFact(approx.ApproximatedSettings, "duplicateSettings")
	if duplicates == nil || duplicates.Path != "/w:settings[1]/w:activeWritingStyle[2]" || len(duplicates.Values) != 1 {
		t.Fatalf("authoring-only duplicates must be one grouped fact: %#v", approx.ApproximatedSettings)
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "applyBreakingRules"); fact == nil || fact.Path != "/w:settings[1]/w:compat[1]/w:applyBreakingRules[1]" || len(fact.Values) != 0 {
		t.Fatalf("legacy compat options must be typed not-applied facts: %#v", approx.ApproximatedSettings)
	}
	if len(approx.ApproximatedSettings) != 6 {
		t.Fatalf("expected authoring, duplicate, applyBreakingRules and three flag facts: %#v", approx.ApproximatedSettings)
	}
}

func TestNativeApproximationDuplicateSettingsMustAgreeOrBeAuthoringOnly(t *testing.T) {
	compat := `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`
	for _, test := range []struct {
		name, markup string
		eligible     bool
	}{
		{"agreeing tab stop", `<w:defaultTabStop w:val="360"/><w:defaultTabStop w:val="360"/>`, true},
		{"disagreeing tab stop", `<w:defaultTabStop w:val="360"/><w:defaultTabStop w:val="720"/>`, false},
		{"agreeing rsidRoot", `<w:rsids><w:rsidRoot w:val="00C1654B"/><w:rsidRoot w:val="00C1654B"/></w:rsids>`, true},
		{"disagreeing rsidRoot", `<w:rsids><w:rsidRoot w:val="00C1654B"/><w:rsidRoot w:val="00C1654C"/></w:rsids>`, false},
		{"agreeing note sentinel", `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/><w:footnote w:id="0"/></w:footnotePr>`, true},
		{"disagreeing writing style", `<w:activeWritingStyle w:appName="MSWord" w:lang="en-US" w:vendorID="64" w:dllVersion="1" w:checkStyle="1"/><w:activeWritingStyle w:appName="MSWord" w:lang="fr-FR" w:vendorID="8" w:dllVersion="1" w:checkStyle="1"/>`, true},
		{"disagreeing container", `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:footnotePr><w:footnote w:id="-1"/></w:footnotePr>`, false},
		{"repeated compat", `<w:compat/>`, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, test.markup+compat))
			if (approx.Status == "eligible") != test.eligible {
				t.Fatalf("wrong eligibility: %#v", approx)
			}
			if test.eligible {
				group := nativeApproximationFact(approx.ApproximatedSettings, "duplicateSettings")
				if group == nil || len(group.Values) != 1 || !strings.HasSuffix(group.Path, "[2]") && !strings.HasSuffix(group.Path, "[3]") {
					t.Fatalf("admitted duplicates must be disclosed as a grouped fact anchored at the duplicate: %#v", approx.ApproximatedSettings)
				}
			} else {
				found := false
				for _, reason := range approx.Reasons {
					found = found || strings.HasPrefix(reason, "Approximate eligibility refused: ")
				}
				if !found {
					t.Fatalf("refusal must disclose its cause: %#v", approx.Reasons)
				}
			}
		})
	}
}

func TestNativeApproximationCompatibilityModeFactsAndLegacyOptions(t *testing.T) {
	mode := func(value string) string {
		return `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="` + value + `"/>`
	}
	flag := `<w:compatSetting w:name="enableOpenTypeFeatures" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>`
	_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:compat><w:useFELayout/><w:doNotExpandShiftReturn w:val="0"/>`+flag+mode("14")+`</w:compat>`))
	if approx.Status != "eligible" || *approx.LegacyCompatibilityMode != 14 {
		t.Fatalf("legacy options plus a non-leading mode must be eligible: %#v", approx)
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "compatibilityMode"); fact == nil || fact.Path != "/w:settings[1]/w:compat[1]/w:compatSetting[2]" || fact.Values["val"] != "14" {
		t.Fatalf("non-leading mode must be a typed fact: %#v", approx.ApproximatedSettings)
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "doNotExpandShiftReturn"); fact == nil || fact.Values["val"] != "0" {
		t.Fatalf("legacy option value must be retained: %#v", approx.ApproximatedSettings)
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "enableOpenTypeFeatures"); fact == nil || fact.Path != "/w:settings[1]/w:compat[1]/w:compatSetting[1]" {
		t.Fatalf("a leading flag is still disclosed at its own path: %#v", approx.ApproximatedSettings)
	}
	_, approx = requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:compat>`+mode("15")+mode("15")+`</w:compat><w:autoHyphenation/>`))
	if approx.Status != "eligible" || *approx.LegacyCompatibilityMode != 15 || nativeApproximationFact(approx.ApproximatedSettings, "compatibilityMode") == nil {
		t.Fatalf("a repeated agreeing mode 15 attestation must be a fact under current-layout mode 15: %#v", approx)
	}
	for name, markup := range map[string]string{
		"two extra mode facts": `<w:compat>` + flag + mode("14") + mode("14") + `</w:compat>`,
		"invalid option value": `<w:compat><w:useFELayout w:val="maybe"/></w:compat>`,
		"option with content":  `<w:compat><w:useFELayout><w:x/></w:useFELayout></w:compat>`,
		"nine facts":           `<w:compat><w:useFELayout/><w:noLeading/><w:noTabHangInd/><w:spaceForUL/><w:ulTrailSpace/><w:wpJustification/><w:growAutofit/><w:useWord97LineBreakRules/><w:mwSmallCaps/>` + mode("14") + `</w:compat>`,
	} {
		_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, markup))
		if approx.Status != "ineligible" {
			t.Fatalf("%s must fail closed: %#v", name, approx)
		}
	}
}

func TestNativeApproximationKeepsGenuinelyUnsupportedSettingsRefused(t *testing.T) {
	for name, markup := range map[string]string{
		"mirror margins":    `<w:mirrorMargins/>`,
		"invalid tab stop":  `<w:defaultTabStop w:val="-1"/>`,
		"structure":         `<w:compat w:unknown="1"/>`,
		"character spacing": `<w:characterSpacingControl w:val="compressPunctuation"/>`,
	} {
		approx, err := ExtractNativeDocxApproximationEligibilityV1(nativeApproximationTestDOCX(t, markup+`<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`))
		if err != nil {
			t.Fatal(err)
		}
		if approx.Status != "ineligible" {
			t.Fatalf("%s must stay ineligible: %#v", name, approx)
		}
	}
}
