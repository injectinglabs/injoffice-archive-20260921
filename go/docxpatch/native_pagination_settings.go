package docxpatch

import (
	"encoding/json"
	"encoding/xml"
	"fmt"
	"strconv"
)

const (
	NativeDOCXPaginationSettingsProtocol = "injoffice.docx.pagination-settings"
	NativeDOCXPaginationSettingsVersion  = 1
	nativeDefaultTabStopTwips            = 720
	nativeSettingsContentType            = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)

var nativePaginationSettingsDiagnosticCodes = map[string]bool{
	"UNKNOWN_SETTINGS_ELEMENT": true, "DUPLICATE_SETTINGS_PROPERTY": true,
	"INVALID_DEFAULT_TAB_STOP": true, "INVALID_SETTINGS_ON_OFF": true,
	"MIRROR_MARGINS_UNSUPPORTED": true, "GUTTER_AT_TOP_UNSUPPORTED": true,
	"CHARACTER_SPACING_CONTROL_UNSUPPORTED": true, "COMPATIBILITY_SETTING_UNSUPPORTED": true,
	"PAGINATION_SETTING_UNSUPPORTED": true, "INVALID_SETTINGS_STRUCTURE": true,
}

// NativePaginationSettingsV1 is a separate, deterministic attestation for the
// small settings.xml subset consumed by the native shaper/paginator. It does
// not change or replace the persisted NativeDocumentV1 contract.
type NativePaginationSettingsV1 struct {
	Protocol            string                                 `json:"protocol"`
	Version             int                                    `json:"version"`
	DocumentID          string                                 `json:"document_id"`
	Revision            string                                 `json:"revision"`
	PackageSHA256       string                                 `json:"package_sha256"`
	MainPart            string                                 `json:"main_part"`
	RelationshipsPart   *string                                `json:"relationships_part,omitempty"`
	RelationshipsSHA256 *string                                `json:"relationships_sha256,omitempty"`
	RelationshipID      *string                                `json:"relationship_id,omitempty"`
	SettingsPart        *string                                `json:"settings_part,omitempty"`
	SettingsSHA256      *string                                `json:"settings_sha256,omitempty"`
	Profile             string                                 `json:"profile"`
	DefaultTabStopTwips int64                                  `json:"default_tab_stop_twips"`
	MirrorMargins       bool                                   `json:"mirror_margins"`
	GutterAtTop         bool                                   `json:"gutter_at_top"`
	EvenAndOddHeaders   bool                                   `json:"even_and_odd_headers"`
	CompatibilityMode   *int                                   `json:"compatibility_mode,omitempty"`
	Diagnostics         []NativePaginationSettingsDiagnosticV1 `json:"diagnostics"`
}

type NativePaginationSettingsDiagnosticV1 struct {
	Code         string `json:"code"`
	Severity     string `json:"severity"`
	PartName     string `json:"part_name"`
	Path         string `json:"path"`
	Preservation string `json:"preservation"`
	Message      string `json:"message"`
}

// ExtractNativePaginationSettingsV1 parses a bounded, relationship-discovered
// settings projection from a DOCX. Unknown or layout-affecting settings produce
// a valid refused attestation rather than being silently ignored.
func ExtractNativePaginationSettingsV1(data []byte) (*NativePaginationSettingsV1, error) {
	return ExtractNativePaginationSettingsV1WithOptions(data, NativeExtractionOptions{})
}

func ExtractNativePaginationSettingsV1WithOptions(data []byte, options NativeExtractionOptions) (*NativePaginationSettingsV1, error) {
	doc, err := ExtractNativeDocumentV1WithOptions(data, options)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		return nil, err
	}
	mainPart, strict, err := pkg.officeDocumentPart()
	if err != nil {
		return nil, err
	}
	wordNS, relBase := wordMLTransitional, relBaseTransitional
	if strict {
		wordNS, relBase = wordMLStrict, relBaseStrict
	}
	result := &NativePaginationSettingsV1{
		Protocol: NativeDOCXPaginationSettingsProtocol, Version: NativeDOCXPaginationSettingsVersion,
		DocumentID: doc.DocumentID, Revision: doc.Revision, PackageSHA256: doc.Source.PackageSHA256, MainPart: mainPart,
		Profile: "absent-default", DefaultTabStopTwips: nativeDefaultTabStopTwips,
		Diagnostics: []NativePaginationSettingsDiagnosticV1{},
	}
	settingsPart, relationshipID, err := nativeSingletonSettingsPart(pkg, mainPart, relBase)
	if err != nil {
		return nil, err
	}
	if settingsPart != "" {
		relationshipsPart := pkg.relsPart[mainPart]
		if relationshipsPart == "" {
			return nil, fmt.Errorf("docxpatch: native pagination settings: settings relationship has no owning relationship part")
		}
		result.RelationshipsPart = nativeString(relationshipsPart)
		relationshipsHash := nativeSHA(pkg.files[relationshipsPart])
		result.RelationshipsSHA256 = &relationshipsHash
		result.RelationshipID = nativeString(relationshipID)
		if !nativeASCIIEqual(pkg.contentTypes[settingsPart], nativeSettingsContentType) {
			return nil, fmt.Errorf("docxpatch: native pagination settings: part %q has content type %q; expected %q", settingsPart, pkg.contentTypes[settingsPart], nativeSettingsContentType)
		}
		root, err := parseNativeXML(settingsPart, pkg.files[settingsPart])
		if err != nil {
			return nil, err
		}
		if root.Name != (xml.Name{Space: wordNS, Local: "settings"}) {
			return nil, fmt.Errorf("docxpatch: native pagination settings: part %q has spoofed or invalid root", settingsPart)
		}
		if err := rejectNativeSettingsNamespaceSpoofing(root, wordNS); err != nil {
			return nil, fmt.Errorf("docxpatch: native pagination settings: %w", err)
		}
		result.SettingsPart = nativeString(settingsPart)
		hash := nativeSHA(pkg.files[settingsPart])
		result.SettingsSHA256 = &hash
		result.Profile = "word-modern-default"
		parseNativePaginationSettings(result, root, wordNS)
		if result.CompatibilityMode == nil {
			result.addDiagnostic("COMPATIBILITY_SETTING_UNSUPPORTED", root, "The settings part does not explicitly attest Word compatibilityMode=15; omitted compatibilityMode defaults to an older layout mode")
		}
	}
	if len(result.Diagnostics) > 0 {
		result.Profile = "unsupported"
	}
	if _, err := EncodeNativePaginationSettingsV1(result); err != nil {
		return nil, fmt.Errorf("docxpatch: native pagination settings output is invalid: %w", err)
	}
	return result, nil
}

func nativeSingletonSettingsPart(pkg *nativePackage, mainPart, relBase string) (string, string, error) {
	want := relBase + "settings"
	wrong := relBaseStrict + "settings"
	if relBase == relBaseStrict {
		wrong = relBaseTransitional + "settings"
	}
	part := ""
	relationshipID := ""
	for _, rel := range pkg.rels[mainPart] {
		if rel.Type == wrong {
			return "", "", fmt.Errorf("docxpatch: native pagination settings: relationship %q uses the wrong Strict/Transitional namespace", rel.ID)
		}
		if rel.Type != want {
			continue
		}
		if rel.External || rel.PartName == "" {
			return "", "", fmt.Errorf("docxpatch: native pagination settings: relationship %q must be internal", rel.ID)
		}
		if part != "" {
			return "", "", fmt.Errorf("docxpatch: native pagination settings: multiple settings relationships")
		}
		part = rel.PartName
		relationshipID = rel.ID
	}
	return part, relationshipID, nil
}

const (
	nativeMCNamespace     = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	nativeMathNamespace   = "http://schemas.openxmlformats.org/officeDocument/2006/math"
	nativeWord14Namespace = "http://schemas.microsoft.com/office/word/2010/wordml"
	nativeWord15Namespace = "http://schemas.microsoft.com/office/word/2012/wordml"
)

func nativeSettingsNamespaceDeclaration(attr xml.Attr) bool {
	return attr.Name.Space == "xmlns" || (attr.Name.Space == "" && attr.Name.Local == "xmlns")
}

func nativeSettingsExactNode(result *NativePaginationSettingsV1, node *nativeXMLNode, allowedAttrs map[xml.Name]bool, allowChildren bool) bool {
	valid := true
	for _, attr := range node.Attrs {
		if nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		if !allowedAttrs[attr.Name] {
			result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, fmt.Sprintf("Unexpected attribute {%s}%s is not part of the attested settings subset", attr.Name.Space, attr.Name.Local))
			valid = false
		}
	}
	if !allowChildren && len(node.Children) != 0 {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "Nested settings content is not part of the attested settings subset")
		valid = false
	}
	if !nativeXMLWhitespaceOnly(node.Text) {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "Non-whitespace settings text is not part of the attested settings subset")
		valid = false
	}
	return valid
}

func nativeSettingsExactLeaf(result *NativePaginationSettingsV1, node *nativeXMLNode, allowedAttrs ...xml.Name) bool {
	allowed := make(map[xml.Name]bool, len(allowedAttrs))
	for _, name := range allowedAttrs {
		allowed[name] = true
	}
	return nativeSettingsExactNode(result, node, allowed, false)
}

func nativeSettingsNeutralWordElement(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	val := xml.Name{Space: wordNS, Local: "val"}
	switch node.Name.Local {
	case "stylePaneFormatFilter", "stylePaneSortMethod", "documentType":
		if !nativeSettingsExactLeaf(result, node, val) {
			return false
		}
		if _, ok := nativeAttr(node, wordNS, "val"); !ok {
			result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "Settings value leaf is missing its required w:val attribute")
			return false
		}
		return true
	case "trackRevisions", "doNotTrackMoves", "doNotTrackFormatting", "autoFormatOverride", "styleLockTheme", "styleLockQFSet", "savePreviewPicture", "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent", "doNotPromoteQF", "doNotAutoCompressPictures", "doNotIncludeSubdocsInStats":
		if !nativeSettingsExactLeaf(result, node, val) {
			return false
		}
		if raw, present := nativeAttr(node, wordNS, "val"); present {
			_, valid := nativeLexicalOnOff(raw)
			if !valid {
				result.addDiagnostic("INVALID_SETTINGS_ON_OFF", node, "Layout-neutral on/off setting has an invalid lexical value")
				return false
			}
		}
		return true
	case "updateFields":
		if !nativeSettingsExactLeaf(result, node, val) {
			return false
		}
		enabled, valid := nativeSettingsOnOff(node, wordNS)
		if !valid {
			result.addDiagnostic("INVALID_SETTINGS_ON_OFF", node, "updateFields has an invalid lexical value")
			return false
		}
		// When enabled, Word may replace visible field results on open. That can
		// change shaped text and pagination after this source snapshot was made.
		return !enabled
	case "zoom":
		return nativeSettingsExactLeaf(result, node, val, xml.Name{Space: wordNS, Local: "percent"})
	case "proofState":
		return nativeSettingsExactLeaf(result, node, xml.Name{Space: wordNS, Local: "spelling"}, xml.Name{Space: wordNS, Local: "grammar"})
	case "clrSchemeMapping":
		allowed := make([]xml.Name, 0, 12)
		for _, local := range []string{"bg1", "t1", "bg2", "t2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hyperlink", "followedHyperlink"} {
			allowed = append(allowed, xml.Name{Space: wordNS, Local: local})
		}
		return nativeSettingsExactLeaf(result, node, allowed...)
	case "rsids":
		if !nativeSettingsExactNode(result, node, map[xml.Name]bool{}, true) {
			return false
		}
		seenRoot := false
		for _, child := range node.Children {
			if child.Name.Space != wordNS || (child.Name.Local != "rsid" && child.Name.Local != "rsidRoot") || !nativeSettingsExactLeaf(result, child, val) {
				result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", child, "rsids may contain only exact w:rsid/w:rsidRoot value leaves")
				return false
			}
			if _, ok := nativeAttr(child, wordNS, "val"); !ok {
				result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", child, "rsid value is required")
				return false
			}
			if child.Name.Local == "rsidRoot" {
				if seenRoot {
					result.addDiagnostic("DUPLICATE_SETTINGS_PROPERTY", child, "Duplicate rsidRoot is ambiguous")
					return false
				}
				seenRoot = true
			}
		}
		return true
	default:
		return false
	}
}

func nativeSettingsNeutralForeignElement(result *NativePaginationSettingsV1, node *nativeXMLNode) bool {
	switch {
	case node.Name == (xml.Name{Space: nativeWord14Namespace, Local: "docId"}):
		val := xml.Name{Space: nativeWord14Namespace, Local: "val"}
		return nativeSettingsExactLeaf(result, node, val) && func() bool { _, ok := nativeAttr(node, nativeWord14Namespace, "val"); return ok }()
	case node.Name == (xml.Name{Space: nativeWord15Namespace, Local: "chartTrackingRefBased"}):
		val := xml.Name{Space: nativeWord15Namespace, Local: "val"}
		if !nativeSettingsExactLeaf(result, node, val) {
			return false
		}
		if raw, present := nativeAttr(node, nativeWord15Namespace, "val"); present {
			_, valid := nativeLexicalOnOff(raw)
			if !valid {
				result.addDiagnostic("INVALID_SETTINGS_ON_OFF", node, "chartTrackingRefBased has an invalid lexical value")
				return false
			}
		}
		return true
	default:
		return false
	}
}

func parseNativePaginationSettings(result *NativePaginationSettingsV1, root *nativeXMLNode, wordNS string) {
	seen := map[string]bool{}
	// An absent autoHyphenation element means no automatic hyphenation. Its
	// zone, cap and consecutive-line options then have no line-layout effect.
	// Resolve this gate before traversal so XML element order cannot change it.
	// https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.autohyphenation
	automaticHyphenationDisabled := true
	autoHyphenation := directNativeChildren(root, wordNS, "autoHyphenation")
	if len(autoHyphenation) > 0 {
		automaticHyphenationDisabled = false
		if len(autoHyphenation) == 1 && nativeExactLeaf(autoHyphenation[0], xml.Name{Space: wordNS, Local: "val"}) {
			enabled, valid := nativeSettingsOnOff(autoHyphenation[0], wordNS)
			automaticHyphenationDisabled = valid && !enabled
		}
	}
	rootAttrs := map[xml.Name]bool{{Space: nativeMCNamespace, Local: "Ignorable"}: true}
	nativeSettingsExactNode(result, root, rootAttrs, true)
	for _, child := range root.Children {
		if child.Name.Space != wordNS {
			key := child.Name.Space + "\x00" + child.Name.Local
			if seen[key] {
				result.addDiagnostic("DUPLICATE_SETTINGS_PROPERTY", child, "Duplicate foreign settings property is ambiguous")
				continue
			}
			seen[key] = true
			if nativeSettingsNeutralForeignElement(result, child) {
				continue
			}
			result.addDiagnostic("UNKNOWN_SETTINGS_ELEMENT", child, "Unknown foreign settings markup can alter Word layout and is not attested")
			continue
		}
		local := child.Name.Local
		if seen[local] {
			result.addDiagnostic("DUPLICATE_SETTINGS_PROPERTY", child, "Duplicate settings property is ambiguous")
			continue
		}
		seen[local] = true
		switch local {
		case "autoHyphenation", "doNotHyphenateCaps":
			if !nativeSettingsExactLeaf(result, child, xml.Name{Space: wordNS, Local: "val"}) {
				continue
			}
			_, valid := nativeSettingsOnOff(child, wordNS)
			if !valid {
				result.addDiagnostic("INVALID_SETTINGS_ON_OFF", child, "Hyphenation switch has an invalid lexical value")
				continue
			}
			if !automaticHyphenationDisabled {
				result.addDiagnostic("PAGINATION_SETTING_UNSUPPORTED", child, "Automatic hyphenation is enabled or ambiguous; dictionary-driven line breaking is not implemented")
			}
		case "hyphenationZone", "consecutiveHyphenLimit":
			if !nativeSettingsExactLeaf(result, child, xml.Name{Space: wordNS, Local: "val"}) {
				continue
			}
			raw, present := nativeAttr(child, wordNS, "val")
			value, err := strconv.ParseUint(raw, 10, 32)
			if !present || err != nil || value > 1_000_000_000 {
				result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", child, "Hyphenation distance/count requires a non-negative bounded integer")
				continue
			}
			if !automaticHyphenationDisabled {
				result.addDiagnostic("PAGINATION_SETTING_UNSUPPORTED", child, "Active hyphenation options require dictionary-driven line breaking; inactive options are preserved without changing layout")
			}
		case "defaultTabStop":
			if !nativeSettingsExactLeaf(result, child, xml.Name{Space: wordNS, Local: "val"}) {
				continue
			}
			raw, ok := nativeAttr(child, wordNS, "val")
			value, err := strconv.ParseInt(raw, 10, 64)
			if !ok || err != nil || value <= 0 || value > 1_000_000_000 {
				result.addDiagnostic("INVALID_DEFAULT_TAB_STOP", child, "defaultTabStop must be a positive bounded twip value")
				continue
			}
			result.DefaultTabStopTwips = value
		case "mirrorMargins", "gutterAtTop", "evenAndOddHeaders":
			if !nativeSettingsExactLeaf(result, child, xml.Name{Space: wordNS, Local: "val"}) {
				continue
			}
			value, valid := nativeSettingsOnOff(child, wordNS)
			if !valid {
				result.addDiagnostic("INVALID_SETTINGS_ON_OFF", child, "Settings on/off value is invalid")
				continue
			}
			switch local {
			case "mirrorMargins":
				result.MirrorMargins = value
				if value {
					result.addDiagnostic("MIRROR_MARGINS_UNSUPPORTED", child, "Mirror margins require page-parity-aware inside/outside gutter placement")
				}
			case "gutterAtTop":
				result.GutterAtTop = value
				if value {
					result.addDiagnostic("GUTTER_AT_TOP_UNSUPPORTED", child, "Top gutter changes the native body box and is not represented by pagination v1")
				}
			case "evenAndOddHeaders":
				result.EvenAndOddHeaders = value
			}
		case "compat":
			if !nativeSettingsExactNode(result, child, map[xml.Name]bool{}, true) {
				continue
			}
			parseNativeModernCompatibility(result, child, wordNS)
		case "footnotePr", "endnotePr":
			parseNativeNoteSentinelRegistrations(result, child, wordNS)
		case "characterSpacingControl":
			if !nativeSettingsExactLeaf(result, child, xml.Name{Space: wordNS, Local: "val"}) {
				continue
			}
			raw, ok := nativeAttr(child, wordNS, "val")
			if !ok || raw != "doNotCompress" {
				result.addDiagnostic("CHARACTER_SPACING_CONTROL_UNSUPPORTED", child, "Only doNotCompress is compatible with shaped native advances")
			}
		default:
			if !nativeSettingsNeutralWordElement(result, child, wordNS) {
				result.addDiagnostic("PAGINATION_SETTING_UNSUPPORTED", child, "This settings property is not proven neutral to native shaping and pagination")
			}
		}
	}
}

func parseNativeNoteSentinelRegistrations(result *NativePaginationSettingsV1, property *nativeXMLNode, wordNS string) {
	if !nativeSettingsExactNode(result, property, map[xml.Name]bool{}, true) {
		return
	}
	wantChild := "footnote"
	if property.Name.Local == "endnotePr" {
		wantChild = "endnote"
	}
	seen := map[string]bool{}
	for _, child := range property.Children {
		if child.Name != (xml.Name{Space: wordNS, Local: wantChild}) || !nativeSettingsExactLeaf(result, child, xml.Name{Space: wordNS, Local: "id"}) {
			result.addDiagnostic("PAGINATION_SETTING_UNSUPPORTED", child, "Note settings may contain only exact reserved separator sentinel registrations")
			continue
		}
		id, ok := nativeAttr(child, wordNS, "id")
		if !ok || id != "-1" && id != "0" {
			result.addDiagnostic("PAGINATION_SETTING_UNSUPPORTED", child, "Note settings numbering, placement, restart, custom, and non-sentinel registrations are unsupported")
			continue
		}
		if seen[id] {
			result.addDiagnostic("DUPLICATE_SETTINGS_PROPERTY", child, "Duplicate reserved note sentinel registration is ambiguous")
			continue
		}
		seen[id] = true
	}
	if !seen["-1"] || !seen["0"] {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", property, "Note settings must register exactly the -1 separator and 0 continuation-separator sentinels")
		result.addDiagnostic("PAGINATION_SETTING_UNSUPPORTED", property, "Empty or incomplete note properties are not the sentinel-registration-only subset")
	}
}

func parseNativeModernCompatibility(result *NativePaginationSettingsV1, compat *nativeXMLNode, wordNS string) {
	seenMode := false
	for _, child := range compat.Children {
		if child.Name != (xml.Name{Space: wordNS, Local: "compatSetting"}) {
			result.addDiagnostic("COMPATIBILITY_SETTING_UNSUPPORTED", child, "Legacy compatibility markup changes Word layout and is not resolved")
			continue
		}
		if !nativeSettingsExactLeaf(result, child,
			xml.Name{Space: wordNS, Local: "name"}, xml.Name{Space: wordNS, Local: "uri"}, xml.Name{Space: wordNS, Local: "val"}) {
			continue
		}
		name, okName := nativeAttr(child, wordNS, "name")
		uri, okURI := nativeAttr(child, wordNS, "uri")
		raw, okValue := nativeAttr(child, wordNS, "val")
		mode, err := strconv.Atoi(raw)
		if !okName || !okURI || !okValue || name != "compatibilityMode" || uri != "http://schemas.microsoft.com/office/word" || err != nil || mode != 15 || seenMode {
			result.addDiagnostic("COMPATIBILITY_SETTING_UNSUPPORTED", child, "Only one modern Word compatibilityMode=15 attestation is supported")
			continue
		}
		seenMode = true
		result.CompatibilityMode = &mode
	}
}

func nativeSettingsOnOff(node *nativeXMLNode, wordNS string) (bool, bool) {
	raw, present := nativeAttr(node, wordNS, "val")
	if !present {
		return true, true
	}
	return nativeLexicalOnOff(raw)
}

func (result *NativePaginationSettingsV1) addDiagnostic(code string, node *nativeXMLNode, message string) {
	if len(result.Diagnostics) >= NativeDOCXMaxIssues {
		return
	}
	part := ""
	if result.SettingsPart != nil {
		part = *result.SettingsPart
	}
	result.Diagnostics = append(result.Diagnostics, NativePaginationSettingsDiagnosticV1{
		Code: code, Severity: "unsupported", PartName: part, Path: node.Path,
		Preservation: "preserve-verbatim", Message: message,
	})
}

func rejectNativeSettingsNamespaceSpoofing(root *nativeXMLNode, wordNS string) error {
	known := map[string]bool{
		"settings": true, "defaultTabStop": true, "mirrorMargins": true, "gutterAtTop": true,
		"evenAndOddHeaders": true, "compat": true, "compatSetting": true,
		"characterSpacingControl": true,
	}
	return rejectNativeKnownLocalSpoofing(root, wordNS, known)
}

func ValidateNativePaginationSettingsV1(input *NativePaginationSettingsV1) error {
	if input == nil {
		return fmt.Errorf("native pagination settings is nil")
	}
	if input.Protocol != NativeDOCXPaginationSettingsProtocol || input.Version != NativeDOCXPaginationSettingsVersion {
		return fmt.Errorf("unsupported native pagination settings protocol/version")
	}
	if !nativeIDPattern.MatchString(input.DocumentID) || !nativeIDPattern.MatchString(input.Revision) {
		return fmt.Errorf("invalid native pagination settings identity")
	}
	if !nativeSHA256.MatchString(input.PackageSHA256) {
		return fmt.Errorf("invalid native pagination package fingerprint")
	}
	if err := validateNativePartName(input.MainPart); err != nil {
		return fmt.Errorf("invalid native pagination main part: %w", err)
	}
	if (input.SettingsPart == nil) != (input.SettingsSHA256 == nil) {
		return fmt.Errorf("settings part and fingerprint must appear together")
	}
	relationshipClosureCount := 0
	if input.RelationshipsPart != nil {
		relationshipClosureCount++
		if err := validateNativePartName(*input.RelationshipsPart); err != nil {
			return fmt.Errorf("invalid native pagination relationships part: %w", err)
		}
	}
	if input.RelationshipsSHA256 != nil {
		relationshipClosureCount++
		if !nativeSHA256.MatchString(*input.RelationshipsSHA256) {
			return fmt.Errorf("invalid native pagination relationships fingerprint")
		}
	}
	if input.RelationshipID != nil {
		relationshipClosureCount++
		if !nativeIDPattern.MatchString(*input.RelationshipID) {
			return fmt.Errorf("invalid native pagination settings relationship id")
		}
	}
	if relationshipClosureCount != 0 && relationshipClosureCount != 3 {
		return fmt.Errorf("relationships part, fingerprint, and relationship id must appear together")
	}
	if input.SettingsPart != nil {
		if err := validateNativePartName(*input.SettingsPart); err != nil {
			return fmt.Errorf("invalid native pagination settings part: %w", err)
		}
		if input.SettingsSHA256 == nil || !nativeSHA256.MatchString(*input.SettingsSHA256) {
			return fmt.Errorf("invalid native pagination settings fingerprint")
		}
	}
	if input.Profile != "absent-default" && input.Profile != "word-modern-default" && input.Profile != "unsupported" {
		return fmt.Errorf("invalid native pagination settings profile %q", input.Profile)
	}
	if input.DefaultTabStopTwips <= 0 || input.DefaultTabStopTwips > 1_000_000_000 {
		return fmt.Errorf("invalid default tab stop")
	}
	if input.CompatibilityMode != nil && *input.CompatibilityMode != 15 {
		return fmt.Errorf("unsupported compatibility mode")
	}
	if len(input.Diagnostics) > NativeDOCXMaxIssues {
		return fmt.Errorf("native pagination settings diagnostics exceed %d", NativeDOCXMaxIssues)
	}
	if input.Profile == "unsupported" && len(input.Diagnostics) == 0 {
		return fmt.Errorf("unsupported settings profile requires diagnostics")
	}
	if input.Profile != "unsupported" && len(input.Diagnostics) != 0 {
		return fmt.Errorf("supported settings profile cannot carry diagnostics")
	}
	if input.Profile == "absent-default" && (input.SettingsPart != nil || relationshipClosureCount != 0 || input.CompatibilityMode != nil || input.DefaultTabStopTwips != nativeDefaultTabStopTwips || input.MirrorMargins || input.GutterAtTop || input.EvenAndOddHeaders) {
		return fmt.Errorf("absent settings must use exact Word defaults")
	}
	if input.Profile == "word-modern-default" && (input.SettingsPart == nil || input.CompatibilityMode == nil || *input.CompatibilityMode != 15) {
		return fmt.Errorf("modern Word settings profile requires a settings part and explicit compatibility mode 15")
	}
	if input.Profile == "unsupported" && input.SettingsPart == nil {
		return fmt.Errorf("unsupported settings profile requires its preserved source part")
	}
	if input.SettingsPart != nil && relationshipClosureCount != 3 {
		return fmt.Errorf("settings part requires its exact owning relationship closure")
	}
	if input.Profile != "unsupported" && (input.MirrorMargins || input.GutterAtTop) {
		return fmt.Errorf("supported settings profile cannot enable unsupported margin semantics")
	}
	for _, diagnostic := range input.Diagnostics {
		if !nativePaginationSettingsDiagnosticCodes[diagnostic.Code] || diagnostic.Severity != "unsupported" || diagnostic.Preservation != "preserve-verbatim" || diagnostic.Message == "" || len(diagnostic.Message) > 4096 || diagnostic.Path == "" || len(diagnostic.Path) > 4096 {
			return fmt.Errorf("invalid native pagination settings diagnostic")
		}
		if err := validateNativePartName(diagnostic.PartName); err != nil {
			return fmt.Errorf("invalid diagnostic settings part: %w", err)
		}
	}
	return nil
}

func EncodeNativePaginationSettingsV1(input *NativePaginationSettingsV1) ([]byte, error) {
	if err := ValidateNativePaginationSettingsV1(input); err != nil {
		return nil, err
	}
	encoded, err := json.Marshal(input)
	if err != nil {
		return nil, err
	}
	if len(encoded) > NativeDOCXMaxJSONBytes {
		return nil, fmt.Errorf("native pagination settings JSON exceeds %d bytes", NativeDOCXMaxJSONBytes)
	}
	return encoded, nil
}
