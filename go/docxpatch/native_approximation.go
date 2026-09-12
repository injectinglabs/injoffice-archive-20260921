package docxpatch

import "encoding/xml"

// NativeDocxApproximationEligibilityV1 is a separate read-only policy attestation.
// It never changes the strict pagination settings projection or mutation model.
type NativeDocxApproximationEligibilityV1 struct {
	Protocol                string                            `json:"protocol"`
	Version                 int                               `json:"version"`
	DocumentID              string                            `json:"document_id"`
	Revision                string                            `json:"revision"`
	PackageSHA256           string                            `json:"package_sha256"`
	SettingsSHA256          *string                           `json:"settings_sha256"`
	Status                  string                            `json:"status"`
	LegacyCompatibilityMode *int                              `json:"legacy_compatibility_mode"`
	Reasons                 []string                          `json:"reasons"`
	ApproximatedSettings    []NativeDocxApproximatedSettingV1 `json:"approximated_settings,omitempty"`
	AbsentFontSizes         []NativeDocxAbsentFontSizeV1      `json:"absent_font_sizes,omitempty"`
}

// ExtractNativeDocxApproximationEligibilityV1 allows exact legacy mode selection
// and bounded, typed settings facts explicitly disregarded by current layout.
// Malformed/duplicate/unknown settings remain ineligible. This does not qualify
// unsupported document content such as math or legacy VML shapes.
func ExtractNativeDocxApproximationEligibilityV1(data []byte) (*NativeDocxApproximationEligibilityV1, error) {
	settings, err := ExtractNativePaginationSettingsV1(data)
	if err != nil {
		return nil, err
	}
	result := &NativeDocxApproximationEligibilityV1{Protocol: "injoffice.docx.approximation-eligibility", Version: 1, DocumentID: settings.DocumentID, Revision: settings.Revision, PackageSHA256: settings.PackageSHA256, SettingsSHA256: settings.SettingsSHA256, Status: "ineligible", Reasons: []string{}}
	for _, diagnostic := range settings.Diagnostics {
		result.Reasons = append(result.Reasons, diagnostic.Code+": "+diagnostic.Message)
	}
	if settings.Profile == "word-modern-default" {
		result.Reasons = append(result.Reasons, "Modern settings already qualify for strict pagination; approximate legacy policy is not applicable")
		return result, nil
	}
	for _, diagnostic := range settings.Diagnostics {
		if diagnostic.Code != "COMPATIBILITY_SETTING_UNSUPPORTED" && diagnostic.Code != "PAGINATION_SETTING_UNSUPPORTED" && diagnostic.Code != "UNKNOWN_SETTINGS_ELEMENT" {
			return result, nil
		}
	}
	mode := 12
	if settings.SettingsPart != nil {
		pkg, err := openNativeDOCXPackage(data)
		if err != nil {
			return nil, err
		}
		_, strict, err := pkg.officeDocumentPart()
		if err != nil {
			return nil, err
		}
		wordNS := wordMLTransitional
		if strict {
			wordNS = wordMLStrict
		}
		root, err := parseNativeXML(*settings.SettingsPart, pkg.files[*settings.SettingsPart])
		if err != nil {
			return nil, err
		}
		covered := map[string]bool{}
		for _, child := range root.Children {
			if fact := nativeApproximateSetting(child, wordNS); fact != nil {
				result.ApproximatedSettings = append(result.ApproximatedSettings, *fact)
				covered[child.Path] = true
				result.Reasons = append(result.Reasons, nativeApproximationSettingReason(*fact))
			}
		}
		for _, diagnostic := range settings.Diagnostics {
			if diagnostic.Code != "COMPATIBILITY_SETTING_UNSUPPORTED" && !covered[diagnostic.Path] {
				return result, nil
			}
		}
		compat := directNativeChildren(root, wordNS, "compat")
		if len(compat) > 1 {
			return result, nil
		}
		if len(compat) == 1 {
			if !nativeExactContainer(compat[0]) {
				return result, nil
			}
			seen := map[string]bool{}
			for _, child := range compat[0].Children {
				if child.Name != (xml.Name{Space: wordNS, Local: "compatSetting"}) || !nativeExactLeaf(child, xml.Name{Space: wordNS, Local: "name"}, xml.Name{Space: wordNS, Local: "uri"}, xml.Name{Space: wordNS, Local: "val"}) {
					return result, nil
				}
				name, _ := nativeAttr(child, wordNS, "name")
				uri, _ := nativeAttr(child, wordNS, "uri")
				value, _ := nativeAttr(child, wordNS, "val")
				if uri != "http://schemas.microsoft.com/office/word" || seen[name] {
					return result, nil
				}
				seen[name] = true
				if name != "compatibilityMode" {
					// The first source compatSetting is reserved for the explicit mode
					// in this bounded policy. This makes its wire path unambiguous:
					// a flag cannot masquerade as omitted-mode evidence downstream.
					if !seen["compatibilityMode"] {
						return result, nil
					}
					if (name != "overrideTableStyleFontSizeAndJustification" && name != "enableOpenTypeFeatures" && name != "doNotFlipMirrorIndents") || value != "1" {
						return result, nil
					}
					fact := NativeDocxApproximatedSettingV1{Kind: name, Path: child.Path, Values: map[string]string{"val": value}}
					result.ApproximatedSettings = append(result.ApproximatedSettings, fact)
					result.Reasons = append(result.Reasons, nativeApproximationSettingReason(fact))
					continue
				}
				if (value != "12" && value != "14") || child != compat[0].Children[0] {
					return result, nil
				}
				if value == "14" {
					mode = 14
				}
			}
		}
	}
	result.Status = "eligible"
	result.LegacyCompatibilityMode = &mode
	result.Reasons = append(result.Reasons, "Read-only approximation uses InjOffice current layout policy, not legacy Microsoft Word layout semantics; page breaks and spacing may differ")
	result.AbsentFontSizes, err = nativeAbsentFontSizes(data)
	if err != nil {
		return nil, err
	}
	return result, nil
}
