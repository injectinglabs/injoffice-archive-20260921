package docxpatch

import (
	"bytes"
	"encoding/json"
	"os"
	"reflect"
	"strings"
	"testing"
)

func TestNativePaginationSettingsV1SharedFixture(t *testing.T) {
	data, err := os.ReadFile("../../testdata/docx-native/pagination-settings-v1.json")
	if err != nil {
		t.Fatal(err)
	}
	var settings NativePaginationSettingsV1
	if err := json.Unmarshal(data, &settings); err != nil {
		t.Fatal(err)
	}
	if err := ValidateNativePaginationSettingsV1(&settings); err != nil {
		t.Fatalf("shared pagination settings fixture: %v", err)
	}
	encoded, err := EncodeNativePaginationSettingsV1(&settings)
	if err != nil {
		t.Fatal(err)
	}
	var fixtureWire, encodedWire any
	if err := json.Unmarshal(data, &fixtureWire); err != nil {
		t.Fatal(err)
	}
	if err := json.Unmarshal(encoded, &encodedWire); err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(fixtureWire, encodedWire) {
		t.Fatalf("shared fixture fields drifted from Go binding: fixture=%v encoded=%v", fixtureWire, encodedWire)
	}
	for _, test := range []struct {
		name   string
		mutate func(*NativePaginationSettingsV1)
	}{
		{"document id slash", func(value *NativePaginationSettingsV1) { value.DocumentID += "/invalid" }},
		{"revision slash", func(value *NativePaginationSettingsV1) { value.Revision += "/invalid" }},
		{"relationship id slash", func(value *NativePaginationSettingsV1) {
			invalid := *value.RelationshipID + "/invalid"
			value.RelationshipID = &invalid
		}},
	} {
		t.Run(test.name, func(t *testing.T) {
			candidate := settings
			test.mutate(&candidate)
			if err := ValidateNativePaginationSettingsV1(&candidate); err == nil {
				t.Fatal("Go binding accepted an id outside nativeIDPattern")
			}
		})
	}
	xmlFixture, err := os.ReadFile("../../testdata/docx-native/settings-word-modern.xml")
	if err != nil {
		t.Fatal(err)
	}
	if settings.SettingsSHA256 == nil || *settings.SettingsSHA256 != nativeSHA(xmlFixture) {
		t.Fatalf("shared settings fingerprint drift: got %v, want %s", settings.SettingsSHA256, nativeSHA(xmlFixture))
	}
}

func nativePaginationSettingsParts(settings string) map[string]string {
	parts := map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/settings.xml" ContentType="` + nativeSettingsContentType + `"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="WORD/DOCUMENT.xml"/></Relationships>`,
		"Word/Document.XML":            `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>settings</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
		"Word/_RELS/Document.XML.RELS": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="settings" Type="` + relBaseTransitional + `settings" Target="SETTINGS.xml"/></Relationships>`,
		"Word/Settings.XML":            settings,
	}
	return parts
}

func TestNativePaginationPrerequisiteProjectionsJoinDeterministicallyInBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS := "Transitional", wordMLTransitional
		if strict {
			name, wordNS = "Strict", wordMLStrict
		}
		t.Run(name, func(t *testing.T) {
			settingsXML := `<w:settings xmlns:w="` + wordNS + `"><w:defaultTabStop w:val="720"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
			parts := nativePaginationSettingsParts(settingsXML)
			parts["Word/Document.XML"] = `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr><w:spacing w:before="120" w:after="240"/><w:ind w:start="720"/><w:keepNext/><w:keepLines/><w:pageBreakBefore w:val="false"/><w:widowControl/></w:pPr><w:r><w:t>joined prerequisites</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`
			if strict {
				parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			document, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			settings, err := ExtractNativePaginationSettingsV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if resolved.DocumentID != document.DocumentID || settings.DocumentID != document.DocumentID || resolved.Revision != document.Revision || settings.Revision != document.Revision || resolved.SourceParts.MainPart != document.Source.MainPart || settings.MainPart != document.Source.MainPart || settings.PackageSHA256 != document.Source.PackageSHA256 {
				t.Fatalf("native pagination prerequisites do not exact-join: document=%#v resolved=%#v settings=%#v", document.Source, resolved.SourceParts, settings)
			}
			if hasUnsupportedCode(document, "PARTIAL_PARAGRAPH_PROPERTIES") {
				t.Fatalf("exact direct pagination properties were self-refused: %#v", document.Unsupported)
			}
			properties := resolved.Paragraphs[0].Properties
			if properties.SpacingBeforeTwips == nil || *properties.SpacingBeforeTwips != 120 || properties.SpacingAfterTwips == nil || *properties.SpacingAfterTwips != 240 || properties.IndentStartTwips == nil || *properties.IndentStartTwips != 720 || properties.KeepNext == nil || !*properties.KeepNext || properties.KeepLines == nil || !*properties.KeepLines || properties.PageBreakBefore == nil || *properties.PageBreakBefore || properties.WidowControl == nil || !*properties.WidowControl {
				t.Fatalf("resolved pagination properties = %#v", properties)
			}
			encoders := []func() ([]byte, error){
				func() ([]byte, error) { return EncodeNativeDocumentV1(document) },
				func() ([]byte, error) { return EncodeNativeResolvedLayoutInputV1(resolved) },
				func() ([]byte, error) { return EncodeNativePaginationSettingsV1(settings) },
			}
			for _, encode := range encoders {
				first, encodeErr := encode()
				if encodeErr != nil {
					t.Fatal(encodeErr)
				}
				second, encodeErr := encode()
				if encodeErr != nil || !bytes.Equal(first, second) {
					t.Fatalf("native pagination prerequisite encoding is not deterministic: %v", encodeErr)
				}
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1ModernWordFixture(t *testing.T) {
	settings, err := os.ReadFile("../../testdata/docx-native/settings-word-modern.xml")
	if err != nil {
		t.Fatal(err)
	}
	data := buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(string(settings))))
	first, err := ExtractNativePaginationSettingsV1WithOptions(data, NativeExtractionOptions{DocumentID: "document:settings"})
	if err != nil {
		t.Fatal(err)
	}
	if first.Profile != "word-modern-default" || first.PackageSHA256 != nativeSHA(data) || first.SettingsPart == nil || *first.SettingsPart != "Word/Settings.XML" || first.SettingsSHA256 == nil || first.RelationshipsPart == nil || *first.RelationshipsPart != "Word/_RELS/Document.XML.RELS" || first.RelationshipsSHA256 == nil || *first.RelationshipsSHA256 != nativeSHA([]byte(nativePaginationSettingsParts(string(settings))["Word/_RELS/Document.XML.RELS"])) || first.RelationshipID == nil || *first.RelationshipID != "settings" || first.DefaultTabStopTwips != 720 || first.CompatibilityMode == nil || *first.CompatibilityMode != 15 || first.MirrorMargins || first.GutterAtTop || first.EvenAndOddHeaders || len(first.Diagnostics) != 0 {
		t.Fatalf("unexpected modern settings attestation: %#v", first)
	}
	encoded, err := EncodeNativePaginationSettingsV1(first)
	if err != nil {
		t.Fatal(err)
	}
	again, err := ExtractNativePaginationSettingsV1WithOptions(data, NativeExtractionOptions{DocumentID: "document:settings"})
	if err != nil {
		t.Fatal(err)
	}
	second, _ := EncodeNativePaginationSettingsV1(again)
	if !bytes.Equal(encoded, second) {
		t.Fatal("pagination settings attestation is not deterministic")
	}
}

func TestExtractNativePaginationSettingsV1StrictRelocatedCaseEquivalentPart(t *testing.T) {
	settings := `<w:settings xmlns:w="` + wordMLStrict + `"><w:defaultTabStop w:val="960"/><w:characterSpacingControl w:val="doNotCompress"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat><w:evenAndOddHeaders/></w:settings>`
	parts := nativePaginationSettingsParts(settings)
	parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
	parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
	parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict, 1)
	result, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if result.Profile != "word-modern-default" || result.DefaultTabStopTwips != 960 || !result.EvenAndOddHeaders || len(result.Diagnostics) != 0 {
		t.Fatalf("unexpected Strict settings projection: %#v", result)
	}
}

func TestExtractNativePaginationSettingsV1RefusesMissingCompatibilityModeInBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS := "Transitional", wordMLTransitional
		if strict {
			name, wordNS = "Strict", wordMLStrict
		}
		t.Run(name, func(t *testing.T) {
			parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordNS + `"><w:defaultTabStop w:val="720"/></w:settings>`)
			if strict {
				parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict, 1)
			}
			settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if settings.Profile != "unsupported" || settings.CompatibilityMode != nil {
				t.Fatalf("missing compatibilityMode must refuse, got %#v", settings)
			}
			found := false
			for _, diagnostic := range settings.Diagnostics {
				found = found || diagnostic.Code == "COMPATIBILITY_SETTING_UNSUPPORTED"
			}
			if !found {
				t.Fatalf("missing compatibility refusal: %#v", settings.Diagnostics)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1AbsentDefaults(t *testing.T) {
	parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordMLTransitional + `"/>`)
	delete(parts, "Word/Settings.XML")
	delete(parts, "Word/_RELS/Document.XML.RELS")
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `<Override PartName="/word/settings.xml" ContentType="`+nativeSettingsContentType+`"/>`, "", 1)
	settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if settings.Profile != "absent-default" || settings.SettingsPart != nil || settings.DefaultTabStopTwips != nativeDefaultTabStopTwips || settings.MirrorMargins || settings.GutterAtTop || settings.EvenAndOddHeaders || len(settings.Diagnostics) != 0 {
		t.Fatalf("unexpected absent settings defaults: %#v", settings)
	}
}

func TestExtractNativePaginationSettingsV1AcceptsOnlyDisabledFieldUpdates(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS := "Transitional", wordMLTransitional
		if strict {
			name, wordNS = "Strict", wordMLStrict
		}
		t.Run(name, func(t *testing.T) {
			parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordNS + `"><w:updateFields w:val="false"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`)
			if strict {
				parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict, 1)
			}
			settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if settings.Profile != "word-modern-default" || len(settings.Diagnostics) != 0 {
				t.Fatalf("disabled field updates should be attested: %#v", settings)
			}
		})
	}
}

func TestNativePaginationStatsSettingDoesNotChangeLayoutAndRejectsSmuggling(t *testing.T) {
	for _, test := range []struct {
		markup   string
		accepted bool
	}{
		{`<w:doNotIncludeSubdocsInStats/>`, true},
		{`<w:doNotIncludeSubdocsInStats w:val="true"/>`, true},
		{`<w:doNotIncludeSubdocsInStats w:val="false"/>`, true},
		{`<w:doNotIncludeSubdocsInStats w:val="maybe"/>`, false},
		{`<w:doNotIncludeSubdocsInStats w:layout="1"/>`, false},
		{`<w:doNotIncludeSubdocsInStats><w:mirrorMargins/></w:doNotIncludeSubdocsInStats>`, false},
	} {
		parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordMLTransitional + `">` + test.markup + `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`)
		settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if (settings.Profile == "word-modern-default") != test.accepted {
			t.Fatalf("%s: unexpected attestation %#v", test.markup, settings)
		}
	}
}

func TestNativePaginationInactiveHyphenationPreservesManualHyphens(t *testing.T) {
	parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordMLTransitional + `"><w:hyphenationZone w:val="360"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`)
	parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], `<w:t>settings</w:t>`, `<w:t>well-known</w:t><w:softHyphen/><w:t>word</w:t>`, 1)
	data := buildNativeDOCX(t, nativeEntries(parts))
	before := append([]byte(nil), data...)
	settings, err := ExtractNativePaginationSettingsV1(data)
	if err != nil || settings.Profile != "word-modern-default" {
		t.Fatalf("inactive automatic policy: settings=%#v err=%v", settings, err)
	}
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	runs := doc.Body.Blocks[0].Paragraph.Runs
	if len(runs) != 3 || runs[0].Text == nil || *runs[0].Text != "well-known" || runs[1].Control != "soft-hyphen" || runs[2].Text == nil || *runs[2].Text != "word" {
		t.Fatalf("manual hyphen content changed: %#v", runs)
	}
	if !bytes.Equal(before, data) {
		t.Fatal("hyphenation qualification mutated source")
	}
}

func TestNativePaginationHyphenationPolicyIsOrderIndependentAndFailClosed(t *testing.T) {
	options := `<w:hyphenationZone w:val="360"/><w:consecutiveHyphenLimit w:val="2"/><w:doNotHyphenateCaps/>`
	for _, test := range []struct {
		name, markup string
		accepted     bool
	}{
		{"default disabled", options, true},
		{"disabled before", `<w:autoHyphenation w:val="false"/>` + options, true},
		{"disabled after", options + `<w:autoHyphenation w:val="0"/>`, true},
		{"enabled before", `<w:autoHyphenation/>` + options, false},
		{"enabled after", options + `<w:autoHyphenation w:val="true"/>`, false},
		{"duplicate gate", `<w:autoHyphenation w:val="0"/>` + options + `<w:autoHyphenation w:val="1"/>`, false},
		{"invalid gate", options + `<w:autoHyphenation w:val="maybe"/>`, false},
		{"foreign gate", options + `<x:autoHyphenation xmlns:x="urn:foreign" w:val="false"/>`, false},
		{"negative distance", `<w:hyphenationZone w:val="-1"/>`, false},
		{"oversized distance", `<w:hyphenationZone w:val="1000000001"/>`, false},
		{"missing value", `<w:hyphenationZone/>`, false},
		{"nested markup", `<w:hyphenationZone w:val="360"><w:autoHyphenation/></w:hyphenationZone>`, false},
		{"unknown attribute", `<w:doNotHyphenateCaps w:layout="1"/>`, false},
	} {
		for _, ns := range []string{wordMLTransitional, wordMLStrict} {
			t.Run(test.name+ns, func(t *testing.T) {
				parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + ns + `">` + test.markup + `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`)
				if ns == wordMLStrict {
					parts["_rels/.rels"] = strings.ReplaceAll(parts["_rels/.rels"], relBaseTransitional, relBaseStrict)
					parts["Word/_RELS/Document.XML.RELS"] = strings.ReplaceAll(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict)
					parts["Word/Document.XML"] = strings.ReplaceAll(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict)
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := append([]byte(nil), data...)
				settings, err := ExtractNativePaginationSettingsV1(data)
				if err != nil {
					if test.accepted {
						t.Fatal(err)
					}
					return
				}
				if (settings.Profile == "word-modern-default") != test.accepted {
					t.Fatalf("unexpected policy: %#v", settings)
				}
				if !bytes.Equal(before, data) {
					t.Fatal("settings attestation mutated source")
				}
			})
		}
	}
}

func TestExtractNativePaginationSettingsV1UsesASCIIOnlyContentTypeEquality(t *testing.T) {
	for _, test := range []struct {
		name        string
		contentType string
		valid       bool
	}{
		{"ASCII case", strings.ToUpper(nativeSettingsContentType), true},
		{"Unicode long s", strings.Replace(nativeSettingsContentType, "settings", "settingſ", 1), false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordMLTransitional + `"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`)
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], nativeSettingsContentType, test.contentType, 1)
			settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
			if test.valid {
				if err != nil || settings.Profile != "word-modern-default" {
					t.Fatalf("ASCII case-equivalent MIME should be accepted: settings=%#v err=%v", settings, err)
				}
				return
			}
			if err == nil || !strings.Contains(err.Error(), "has content type") {
				t.Fatalf("non-ASCII MIME fold must be rejected, err=%v", err)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1RefusesLayoutAffectingSettings(t *testing.T) {
	settingsXML := `<w:settings xmlns:w="` + wordMLTransitional + `"><w:defaultTabStop w:val="360"/><w:mirrorMargins/><w:gutterAtTop/><w:evenAndOddHeaders/><w:footnotePr/><w:compat><w:usePrinterMetrics/></w:compat></w:settings>`
	settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(settingsXML))))
	if err != nil {
		t.Fatal(err)
	}
	if settings.Profile != "unsupported" || settings.DefaultTabStopTwips != 360 || !settings.MirrorMargins || !settings.GutterAtTop || !settings.EvenAndOddHeaders {
		t.Fatalf("unsupported settings were not retained: %#v", settings)
	}
	for _, code := range []string{"MIRROR_MARGINS_UNSUPPORTED", "GUTTER_AT_TOP_UNSUPPORTED", "PAGINATION_SETTING_UNSUPPORTED", "COMPATIBILITY_SETTING_UNSUPPORTED"} {
		found := false
		for _, diagnostic := range settings.Diagnostics {
			if diagnostic.Code == code {
				found = true
				break
			}
		}
		if !found {
			t.Fatalf("missing %s diagnostic: %#v", code, settings.Diagnostics)
		}
	}
}

func TestExtractNativePaginationSettingsV1AcceptsOnlyExactNoteSentinelRegistrations(t *testing.T) {
	for _, wordNS := range []string{wordMLTransitional, wordMLStrict} {
		settingsXML := `<w:settings xmlns:w="` + wordNS + `"><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
		parts := nativePaginationSettingsParts(settingsXML)
		if wordNS == wordMLStrict {
			parts["_rels/.rels"] = strings.ReplaceAll(parts["_rels/.rels"], relBaseTransitional, relBaseStrict)
			parts["Word/_RELS/Document.XML.RELS"] = strings.ReplaceAll(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict)
			parts["Word/Document.XML"] = strings.ReplaceAll(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict)
		}
		settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil || settings.Profile != "word-modern-default" || len(settings.Diagnostics) != 0 {
			t.Fatalf("%s exact note registrations: profile=%q diagnostics=%#v err=%v", wordNS, settings.Profile, settings.Diagnostics, err)
		}
	}

	for name, property := range map[string]string{
		"numbering": `<w:footnotePr><w:numFmt w:val="decimal"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`,
		"placement": `<w:footnotePr><w:pos w:val="pageBottom"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`,
		"restart":   `<w:footnotePr><w:numRestart w:val="eachPage"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`,
		"custom id": `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="7"/></w:footnotePr>`,
		"duplicate": `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`,
		"spoof":     `<w:footnotePr xmlns:x="urn:spoof"><x:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr>`,
	} {
		t.Run(name, func(t *testing.T) {
			xml := `<w:settings xmlns:w="` + wordMLTransitional + `">` + property + `<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
			settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(xml))))
			if err != nil {
				t.Fatal(err)
			}
			if settings.Profile != "unsupported" || len(settings.Diagnostics) == 0 {
				t.Fatalf("non-exact note settings were accepted: %#v", settings)
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1RefusesColumnBalanceSettingsInBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS := "Transitional", wordMLTransitional
		if strict {
			name, wordNS = "Strict", wordMLStrict
		}
		t.Run(name, func(t *testing.T) {
			settingsXML := `<w:settings xmlns:w="` + wordNS + `"><w:cachedColBalance/><w:compat><w:noColumnBalance/><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
			parts := nativePaginationSettingsParts(settingsXML)
			if strict {
				parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
				parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict, 1)
			}
			settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if settings.Profile != "unsupported" {
				t.Fatalf("column-balance settings profile = %q, want unsupported", settings.Profile)
			}
			for _, code := range []string{"PAGINATION_SETTING_UNSUPPORTED", "COMPATIBILITY_SETTING_UNSUPPORTED"} {
				found := false
				for _, diagnostic := range settings.Diagnostics {
					if diagnostic.Code == code {
						found = true
					}
				}
				if !found {
					t.Fatalf("missing %s for column-balance settings: %#v", code, settings.Diagnostics)
				}
			}
		})
	}
}

func TestExtractNativePaginationSettingsV1RejectsSpoofAndAmbiguousRelationship(t *testing.T) {
	spoof := `<w:settings xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:spoof"><x:defaultTabStop w:val="720"/></w:settings>`
	if _, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(nativePaginationSettingsParts(spoof)))); err == nil || !strings.Contains(err.Error(), "namespace spoofing") {
		t.Fatalf("namespace spoof error = %v", err)
	}
	parts := nativePaginationSettingsParts(`<w:settings xmlns:w="` + wordMLTransitional + `"/>`)
	parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], `</Relationships>`, `<Relationship Id="settings2" Type="`+relBaseTransitional+`settings" Target="SETTINGS.xml"/></Relationships>`, 1)
	if _, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "multiple settings relationships") {
		t.Fatalf("duplicate settings relationship error = %v", err)
	}
}

func TestExtractNativePaginationSettingsV1StructurallyRefusesAcceptedElementSmuggling(t *testing.T) {
	tests := []struct {
		name string
		xml  func(wordNS string) string
		code string
	}{
		{"root attribute", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `" bogus="1"><w:defaultTabStop w:val="720"/></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"nested modeled setting", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:zoom><w:mirrorMargins/></w:zoom></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"unexpected attribute", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:defaultTabStop w:val="720" bogus="1"/></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"on off child", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:evenAndOddHeaders><w:zoom w:percent="100"/></w:evenAndOddHeaders></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"non whitespace text", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:characterSpacingControl w:val="doNotCompress">smuggled</w:characterSpacingControl></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"non XML S unicode separator", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:characterSpacingControl w:val="doNotCompress">&#xA0;</w:characterSpacingControl></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"duplicate neutral singleton", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:zoom w:percent="100"/><w:zoom w:percent="90"/></w:settings>`
		}, "DUPLICATE_SETTINGS_PROPERTY"},
		{"compat child content", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"><w:mirrorMargins/></w:compatSetting></w:compat></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"compat attribute", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:compat bogus="1"><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"compat setting attribute", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15" bogus="1"/></w:compat></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"unsafe theme setting", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:themeFontLang w:val="en-US"/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"unsafe shape defaults", func(ns string) string { return `<w:settings xmlns:w="` + ns + `"><w:shapeDefaults/></w:settings>` }, "PAGINATION_SETTING_UNSUPPORTED"},
		{"unsafe template", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:attachedTemplate w:val="rId1"/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"unsafe forced compatibility upgrade", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:forceUpgrade/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"field updates enabled implicitly", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:updateFields/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"field updates enabled explicitly", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:updateFields w:val="true"/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"placeholder display semantics", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:alwaysShowPlaceholderText/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"revision display semantics", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `"><w:revisionView w:markup="false"/></w:settings>`
		}, "PAGINATION_SETTING_UNSUPPORTED"},
		{"foreign doc id child", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `" xmlns:w14="` + nativeWord14Namespace + `"><w14:docId w14:val="01234567"><w:mirrorMargins/></w14:docId></w:settings>`
		}, "INVALID_SETTINGS_STRUCTURE"},
		{"native math layout settings", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `" xmlns:m="` + nativeMathNamespace + `"><m:mathPr><m:mathFont m:val="Cambria Math"/><m:smallFrac m:val="1"/><m:lMargin m:val="0"/><m:wrapIndent m:val="1440"/></m:mathPr></w:settings>`
		}, "UNKNOWN_SETTINGS_ELEMENT"},
		{"foreign math unknown", func(ns string) string {
			return `<w:settings xmlns:w="` + ns + `" xmlns:m="` + nativeMathNamespace + `"><m:mathPr><m:unknown m:val="1"/></m:mathPr></w:settings>`
		}, "UNKNOWN_SETTINGS_ELEMENT"},
	}
	for _, strict := range []bool{false, true} {
		for _, test := range tests {
			name := "Transitional/" + test.name
			wordNS := wordMLTransitional
			if strict {
				name, wordNS = "Strict/"+test.name, wordMLStrict
			}
			t.Run(name, func(t *testing.T) {
				parts := nativePaginationSettingsParts(test.xml(wordNS))
				if strict {
					parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], relBaseTransitional, relBaseStrict, 1)
					parts["Word/_RELS/Document.XML.RELS"] = strings.Replace(parts["Word/_RELS/Document.XML.RELS"], relBaseTransitional, relBaseStrict, 1)
					parts["Word/Document.XML"] = strings.Replace(parts["Word/Document.XML"], wordMLTransitional, wordMLStrict, 1)
				}
				settings, err := ExtractNativePaginationSettingsV1(buildNativeDOCX(t, nativeEntries(parts)))
				if err != nil {
					t.Fatal(err)
				}
				if settings.Profile != "unsupported" {
					t.Fatalf("smuggled settings profile = %q, want unsupported", settings.Profile)
				}
				found := false
				for _, diagnostic := range settings.Diagnostics {
					if diagnostic.Code == test.code {
						found = true
						break
					}
				}
				if !found {
					t.Fatalf("missing %s: %#v", test.code, settings.Diagnostics)
				}
			})
		}
	}
}
