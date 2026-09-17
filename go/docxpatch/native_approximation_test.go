package docxpatch

import (
	"bytes"
	"reflect"
	"slices"
	"sort"
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
	if len(approx.ApproximatedSettings) > nativeApproximationMaxFacts {
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
	if fact == nil || fact.Path != "/w:settings[1]/w:autoHyphenation[1]" || !reflect.DeepEqual(fact.Values, map[string]string{"val": "true", "hyphenationZone": "360", "consecutiveHyphenLimit": "2", "doNotHyphenateCaps": "true"}) {
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
	// The drawing-grid, header-shape, proofing and style-pane members of this
	// markup are strict-neutral now, so they carry no diagnostic to join and are
	// not disregarded facts. Only settings strict still refuses stay in the group.
	if group == nil || group.Path != "/w:settings[1]/w:removePersonalInformation[1]" || len(group.Values) != 5 {
		t.Fatalf("authoring settings must be one grouped fact anchored at the first diagnosed member: %#v", approx.ApproximatedSettings)
	}
	for _, neutral := range []string{"activeWritingStyle", "hdrShapeDefaults", "displayHorizontalDrawingGridEvery", "removeDateAndTime", "stylePaneFormatFilter"} {
		if _, present := group.Values["/w:settings[1]/w:"+neutral+"[1]"]; present {
			t.Fatalf("strict-neutral %s must not be disclosed as disregarded: %#v", neutral, group.Values)
		}
	}
	for path, expected := range map[string]string{
		"/w:settings[1]/w:attachedTemplate[1]":          `r:id="rId1"`,
		"/w:settings[1]/w:removePersonalInformation[1]": ``,
		"/w:settings[1]/w:noPunctuationKerning[1]":      ``,
	} {
		if group.Values[path] != expected {
			t.Fatalf("member %s retained %q, want %q", path, group.Values[path], expected)
		}
	}
	// The repeated w:activeWritingStyle is the schema's own per-language shape,
	// so strict no longer calls it a duplicate and there is nothing to group.
	if nativeApproximationFact(approx.ApproximatedSettings, "duplicateSettings") != nil {
		t.Fatalf("schema-repeatable proofing registrations are not duplicates: %#v", approx.ApproximatedSettings)
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "applyBreakingRules"); fact == nil || fact.Path != "/w:settings[1]/w:compat[1]/w:applyBreakingRules[1]" || len(fact.Values) != 0 {
		t.Fatalf("legacy compat options must be typed not-applied facts: %#v", approx.ApproximatedSettings)
	}
	if len(approx.ApproximatedSettings) != 5 {
		t.Fatalf("expected authoring, applyBreakingRules and three flag facts: %#v", approx.ApproximatedSettings)
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
		{"disagreeing authoring-only setting", `<w:noPunctuationKerning/><w:noPunctuationKerning w:val="0"/>`, true},
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
		"two extra mode facts":      `<w:compat>` + flag + mode("14") + mode("14") + `</w:compat>`,
		"invalid option value":      `<w:compat><w:useFELayout w:val="maybe"/></w:compat>`,
		"option with content":       `<w:compat><w:useFELayout><w:x/></w:useFELayout></w:compat>`,
		"more facts than the bound": `<w:compat>` + nativeApproximationEveryLegacyCompatLeaf() + mode("14") + `</w:compat>`,
	} {
		_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, markup))
		if approx.Status != "ineligible" {
			t.Fatalf("%s must fail closed: %#v", name, approx)
		}
	}
}

// nativeApproximationEveryLegacyCompatLeaf emits one leaf per ECMA-376 legacy
// w:compat option, in deterministic order: 65 leaves, one past the fact bound.
func nativeApproximationEveryLegacyCompatLeaf() string {
	names := make([]string, 0, len(nativeApproximateLegacyCompatFlags))
	for name := range nativeApproximateLegacyCompatFlags {
		names = append(names, name)
	}
	sort.Strings(names)
	var out strings.Builder
	for _, name := range names {
		out.WriteString("<w:" + name + "/>")
	}
	return out.String()
}

// A full legacy w:compat block is ordinary Word 97-2003 output. Every option is
// recorded as not applied, so the bound must admit the whole block rather than
// refusing the document over the size of its disclosure vector.
func TestNativeApproximationAdmitsAWholeLegacyCompatBlock(t *testing.T) {
	names := make([]string, 0, len(nativeApproximateLegacyCompatFlags))
	for name := range nativeApproximateLegacyCompatFlags {
		names = append(names, name)
	}
	sort.Strings(names)
	var markup strings.Builder
	for _, name := range names[:len(names)-1] {
		markup.WriteString("<w:" + name + "/>")
	}
	_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:compat>`+markup.String()+`<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`))
	if approx.Status != "eligible" || *approx.LegacyCompatibilityMode != 14 {
		t.Fatalf("a whole legacy compat block must stay eligible: %#v", approx)
	}
	if len(approx.ApproximatedSettings) != len(names)-1 {
		t.Fatalf("every legacy option must be disclosed as its own fact: %d of %d", len(approx.ApproximatedSettings), len(names)-1)
	}
	if len(approx.ApproximatedSettings) > nativeApproximationMaxFacts {
		t.Fatalf("facts exceed the decoder bound: %d", len(approx.ApproximatedSettings))
	}
}

// East Asian punctuation compression is recorded, never performed: the
// approximate tier shapes natural advances, so the value cannot change what it
// paints and is disclosed instead of refusing every East-Asian-locale save.
func TestNativeApproximationRecordsCharacterSpacingControl(t *testing.T) {
	compat := `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`
	for _, value := range []string{"compressPunctuation", "compressPunctuationAndJapaneseKana"} {
		_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:characterSpacingControl w:val="`+value+`"/>`+compat))
		if approx.Status != "eligible" {
			t.Fatalf("%s must be recorded, not refused: %#v", value, approx)
		}
		fact := nativeApproximationFact(approx.ApproximatedSettings, "characterSpacingControl")
		if fact == nil || fact.Path != "/w:settings[1]/w:characterSpacingControl[1]" || fact.Values["val"] != value {
			t.Fatalf("%s must be a typed not-applied fact: %#v", value, approx.ApproximatedSettings)
		}
		if !slices.Contains(approx.Reasons, nativeApproximationSettingReason(*fact)) {
			t.Fatalf("%s must disclose its not-applied reason: %#v", value, approx.Reasons)
		}
	}
	// A value outside ECMA-376 17.15.1.20 is not a known compression mode and
	// still fails closed.
	approx, err := ExtractNativeDocxApproximationEligibilityV1(nativeApproximationTestDOCX(t, `<w:characterSpacingControl w:val="squashEverything"/>`+compat))
	if err != nil {
		t.Fatal(err)
	}
	if approx.Status != "ineligible" {
		t.Fatalf("an unknown character spacing value must stay ineligible: %#v", approx)
	}
}

// Both attested values of a recorded compatSetting flag select Word behaviour
// this tier never emulates, and repeats cannot disagree about anything it uses.
func TestNativeApproximationRecordsRepeatedAndZeroValuedCompatSettings(t *testing.T) {
	setting := func(value string) string {
		return `<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="` + value + `"/>`
	}
	mode := `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>`
	_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:compat>`+mode+setting("0")+`</w:compat>`))
	if approx.Status != "eligible" {
		t.Fatalf("a zero-valued recorded flag must not refuse: %#v", approx)
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "overrideTableStyleFontSizeAndJustification"); fact == nil || fact.Values["val"] != "0" {
		t.Fatalf("the attested value must be retained verbatim: %#v", approx.ApproximatedSettings)
	}
	_, approx = requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:compat>`+mode+setting("1")+setting("1")+setting("0")+`</w:compat>`))
	if approx.Status != "eligible" {
		t.Fatalf("disagreeing repeats of a not-applied flag must not refuse: %#v", approx)
	}
	group := nativeApproximationFact(approx.ApproximatedSettings, "repeatedCompatSettings")
	if group == nil || len(group.Values) != 2 {
		t.Fatalf("every repeat must be disclosed in one grouped fact: %#v", approx.ApproximatedSettings)
	}
	if group.Values["/w:settings[1]/w:compat[1]/w:compatSetting[3]"] != "overrideTableStyleFontSizeAndJustification=1" || group.Values["/w:settings[1]/w:compat[1]/w:compatSetting[4]"] != "overrideTableStyleFontSizeAndJustification=0" {
		t.Fatalf("repeat values must be retained verbatim by path: %#v", group.Values)
	}
	// compatibilityMode is consumed, not recorded, so disagreement still refuses.
	bad, err := ExtractNativeDocxApproximationEligibilityV1(nativeApproximationTestDOCX(t, `<w:compat>`+mode+`<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`))
	if err != nil {
		t.Fatal(err)
	}
	if bad.Status != "ineligible" {
		t.Fatalf("disagreeing compatibilityMode attestations must stay ineligible: %#v", bad)
	}
}

func TestNativeApproximationKeepsGenuinelyUnsupportedSettingsRefused(t *testing.T) {
	for name, markup := range map[string]string{
		"mirror margins":         `<w:mirrorMargins/>`,
		"invalid tab stop":       `<w:defaultTabStop w:val="-1"/>`,
		"structure":              `<w:compat w:unknown="1"/>`,
		"unknown compat setting": `<w:compat><w:compatSetting w:name="madeUpFlag" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/></w:compat>`,
		"note sentinels":         `<w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>`,
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

// A duplicated known extra (themeFontLang, decimalSymbol, ...) is disclosed only
// through the duplicate group. The TS mirror requires index [1] for those kinds,
// so a fact at [2] would turn a disclosed refusal into a decode error.
func TestNativeApproximationDuplicatedKnownExtrasNeverBecomeIndexedFacts(t *testing.T) {
	compat := `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`
	markup := `<w:themeFontLang w:val="en-CA" w:eastAsia=""/><w:themeFontLang w:val="en-CA" w:eastAsia=""/><w:decimalSymbol w:val="."/><w:decimalSymbol w:val="."/>` + compat
	_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, markup))
	if approx.Status != "eligible" {
		t.Fatalf("agreeing duplicates of known extras must stay eligible: %#v", approx)
	}
	for _, fact := range approx.ApproximatedSettings {
		if fact.Kind != "authoringSettings" && fact.Kind != "duplicateSettings" && !strings.HasSuffix(fact.Path, "[1]") {
			t.Fatalf("known extra emitted at a duplicate index: %#v", fact)
		}
	}
	if fact := nativeApproximationFact(approx.ApproximatedSettings, "themeFontLang"); fact == nil || fact.Path != "/w:settings[1]/w:themeFontLang[1]" {
		t.Fatalf("the diagnosed first themeFontLang stays a [1] fact: %#v", approx.ApproximatedSettings)
	}
	if nativeApproximationFact(approx.ApproximatedSettings, "decimalSymbol") != nil {
		t.Fatalf("a strict-neutral first decimalSymbol has no diagnostic and no fact: %#v", approx.ApproximatedSettings)
	}
	group := nativeApproximationFact(approx.ApproximatedSettings, "duplicateSettings")
	if group == nil || len(group.Values) != 2 {
		t.Fatalf("both duplicates belong to the duplicate group: %#v", approx.ApproximatedSettings)
	}
	for _, path := range []string{"/w:settings[1]/w:themeFontLang[2]", "/w:settings[1]/w:decimalSymbol[2]"} {
		if _, present := group.Values[path]; !present {
			t.Fatalf("duplicate %s missing from the group: %#v", path, group.Values)
		}
	}
	// Refusal drops every accumulated fact and its reason.
	refused, err := ExtractNativeDocxApproximationEligibilityV1(nativeApproximationTestDOCX(t, `<w:themeFontLang w:val="en-CA"/><w:themeFontLang w:val="fr-CA"/>`+compat))
	if err != nil {
		t.Fatal(err)
	}
	if refused.Status != "ineligible" || len(refused.ApproximatedSettings) != 0 {
		t.Fatalf("disagreeing known-extra duplicates must refuse without partial facts: %#v", refused)
	}
	for _, reason := range refused.Reasons {
		if strings.HasPrefix(reason, "Current-layout approximation") {
			t.Fatalf("refused attestation must not disclose dropped facts: %q", reason)
		}
	}
}

func TestNativeApproximationBareHyphenationLeavesRecordTrue(t *testing.T) {
	_, approx := requireNativeApproximationDisclosure(t, nativeApproximationTestDOCX(t, `<w:autoHyphenation/><w:doNotHyphenateCaps/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`))
	fact := nativeApproximationFact(approx.ApproximatedSettings, "autoHyphenation")
	if approx.Status != "eligible" || fact == nil || !reflect.DeepEqual(fact.Values, map[string]string{"val": "true", "doNotHyphenateCaps": "true"}) {
		t.Fatalf("omitted w:val must be recorded as the schema default true: %#v", approx.ApproximatedSettings)
	}
}

// The approximate eligibility policy is a separate read-only attestation: every
// newly admitted settings shape must leave strict extraction, the strict
// pagination settings projection and its diagnostics byte-identical.
func TestNativeApproximationNeverChangesStrictSettingsProjection(t *testing.T) {
	compat := `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`
	flag := func(value string) string {
		return `<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="` + value + `"/>`
	}
	for name, markup := range map[string]string{
		"character spacing":         `<w:characterSpacingControl w:val="compressPunctuation"/>` + compat,
		"whole legacy compat block": `<w:compat>` + nativeApproximationEveryLegacyCompatLeaf() + `<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="14"/></w:compat>`,
		"repeated compat settings":  `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>` + flag("1") + flag("0") + `</w:compat>`,
		"note sentinels":            `<w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>` + compat,
	} {
		data := nativeApproximationTestDOCX(t, markup)
		before := append([]byte(nil), data...)
		strict, err := ExtractNativePaginationSettingsV1(data)
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		strictDigest := nativeDOCXCanonicalWireSHA256(strict)
		document, err := ExtractNativeDocumentV1(data)
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		documentDigest := nativeDOCXCanonicalWireSHA256(document)
		if _, err := ExtractNativeDocxApproximationEligibilityV1(data); err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		again, err := ExtractNativePaginationSettingsV1(data)
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		repeated, err := ExtractNativeDocumentV1(data)
		if err != nil {
			t.Fatalf("%s: %v", name, err)
		}
		if !bytes.Equal(before, data) {
			t.Fatalf("%s: source package bytes changed", name)
		}
		if again.Profile != "unsupported" || nativeDOCXCanonicalWireSHA256(again) != strictDigest {
			t.Fatalf("%s: strict settings projection changed: %#v", name, again)
		}
		if nativeDOCXCanonicalWireSHA256(repeated) != documentDigest {
			t.Fatalf("%s: strict document extraction changed", name)
		}
	}
}

// TestNativeApproximationEligibilityNamesTheDiagnosticThatBlockedIt covers the
// two refusals that returned an unnamed or misnamed cause: a strict settings
// diagnostic outside the admitted set returned "ineligible" with nothing in
// Reasons distinguishing it from the admitted disclosures beside it, and the
// bounded-fact refusal still quoted the bound of 8 it was raised away from.
func TestNativeApproximationEligibilityNamesTheDiagnosticThatBlockedIt(t *testing.T) {
	legacyFlags := make([]string, 0, len(nativeApproximateLegacyCompatFlags))
	for name := range nativeApproximateLegacyCompatFlags {
		legacyFlags = append(legacyFlags, "<w:"+name+"/>")
	}
	sort.Strings(legacyFlags)
	if len(legacyFlags) <= nativeApproximationMaxFacts {
		t.Fatalf("fixture needs more than %d recorded legacy compat flags, have %d", nativeApproximationMaxFacts, len(legacyFlags))
	}
	for _, test := range []struct {
		name, markup, want string
	}{
		{
			// A LibreOffice save registers the reserved separator stories as ids 0 and
			// 1 instead of Word's -1 and 0, which strict diagnoses as
			// INVALID_SETTINGS_STRUCTURE: a code the current-layout policy does not
			// admit. The attestation must say so instead of leaving the caller to
			// guess among the PAGINATION_SETTING_UNSUPPORTED disclosures.
			"non-admitted diagnostic",
			`<w:footnotePr><w:footnote w:id="0"/><w:footnote w:id="1"/></w:footnotePr>`,
			"Approximate eligibility refused: strict settings diagnostic INVALID_SETTINGS_STRUCTURE at /w:settings[1]/w:footnotePr[1] is outside the current-layout admitted set",
		},
		{
			"bounded typed facts",
			"<w:compat>" + strings.Join(legacyFlags, "") + "</w:compat>",
			"Approximate eligibility refused: more than 64 typed settings facts would be required",
		},
	} {
		t.Run(test.name, func(t *testing.T) {
			data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(`<w:settings xmlns:w="`+wordMLTransitional+`">`+test.markup+`</w:settings>`)))
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if eligibility.Status != "ineligible" || len(eligibility.ApproximatedSettings) != 0 {
				t.Fatalf("expected an ineligible attestation carrying no typed facts: %#v", eligibility)
			}
			if !slices.Contains(eligibility.Reasons, test.want) {
				t.Fatalf("attestation does not name its blocking cause %q: %#v", test.want, eligibility.Reasons)
			}
			for _, reason := range eligibility.Reasons {
				if strings.Contains(reason, "more than 8 typed settings facts") {
					t.Fatalf("refusal quotes a bound the policy no longer applies: %q", reason)
				}
			}
			named := 0
			for _, reason := range eligibility.Reasons {
				if strings.HasPrefix(reason, "Approximate eligibility refused: ") {
					named++
				}
			}
			if named != 1 {
				t.Fatalf("an ineligible attestation must name exactly one blocking cause, got %d: %#v", named, eligibility.Reasons)
			}
		})
	}
}
