package docxpatch

import (
	"encoding/xml"
	"strconv"
	"strings"
)

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
	// AbsentFontSizeShape names which source shape proved the absence, because
	// Microsoft Word resolves the two shapes to different sizes. Present
	// exactly when AbsentFontSizes is non-empty.
	AbsentFontSizeShape string                          `json:"absent_font_size_shape,omitempty"`
	AbsentFontFamilies  []NativeDocxAbsentFontFamilyV1  `json:"absent_font_families,omitempty"`
	LatinFontFallbacks  []NativeDocxLatinFontFallbackV1 `json:"latin_font_fallbacks,omitempty"`
	LegacyTableOrigins  []NativeDocxLegacyTableOriginV1 `json:"legacy_table_origins,omitempty"`
}

const nativeApproximationCompatSettingURI = "http://schemas.microsoft.com/office/word"

// nativeApproximationAdmittedDiagnostics are the strict settings diagnostic
// codes the current-layout policy can still disclose as typed "not applied"
// facts. Every other code (malformed structure, unrepresentable note
// registrations, foreign markup that strict could not even shape into a
// diagnosable leaf) keeps the attestation ineligible.
var nativeApproximationAdmittedDiagnostics = map[string]bool{
	"COMPATIBILITY_SETTING_UNSUPPORTED": true,
	"PAGINATION_SETTING_UNSUPPORTED":    true,
	"UNKNOWN_SETTINGS_ELEMENT":          true,
	"DUPLICATE_SETTINGS_PROPERTY":       true,
	// ECMA-376 17.15.1.20 w:characterSpacingControl selects East Asian
	// punctuation compression. The approximate tier never compresses advances,
	// so the value is recorded as not applied instead of refusing every
	// East-Asian-locale save; strict pagination keeps refusing it unchanged.
	"CHARACTER_SPACING_CONTROL_UNSUPPORTED": true,
}

// ExtractNativeDocxApproximationEligibilityV1 allows exact legacy mode 12 or 14,
// plus a current-layout fallback when Word attests mode 15 but extras keep the
// strict profile unsupported. Every setting the policy admits beyond strict
// parsing is disclosed as a typed approximated_settings fact joined to its strict
// diagnostic path: bounded known extras, active autoHyphenation (recorded, not
// performed), repeated or non-leading agreeing compatibilityMode attestations,
// ECMA-376 w:compat legacy options and Microsoft compatSetting flags (recorded,
// not applied), authoring-only settings.xml children, and duplicate settings
// that repeat their first occurrence or are authoring-only. Malformed settings,
// disagreeing duplicates, unknown compat markup, and more typed facts than the
// bounded disclosure vector holds remain ineligible; every one of those refusals
// names its own cause in Reasons, because the vector also carries admitted facts
// that did not block. Remaining PAGINATION_SETTING_UNSUPPORTED and UNKNOWN_SETTINGS_ELEMENT
// diagnostics are retained verbatim as reasons. This does not qualify unsupported
// document content such as math or legacy VML shapes for strict paint.
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
		if !nativeApproximationAdmittedDiagnostics[diagnostic.Code] {
			// Reasons already restates every strict settings diagnostic, admitted or
			// not, so returning here without naming the one that blocked leaves the
			// caller a list of disclosures with no way to tell which refused.
			result.Reasons = append(result.Reasons, "Approximate eligibility refused: strict settings diagnostic "+diagnostic.Code+" at "+diagnostic.Path+" is outside the current-layout admitted set")
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
		builder := newNativeApproximationBuilder(result, wordNS, settings.Diagnostics)
		if !builder.topLevel(root) || !builder.duplicates(root, settings.Diagnostics) || !builder.compat(root, &mode) {
			return result, nil
		}
	}
	result.Status = "eligible"
	result.LegacyCompatibilityMode = &mode
	result.Reasons = append(result.Reasons, "Read-only approximation uses InjOffice current layout policy, not legacy Microsoft Word layout semantics; page breaks and spacing may differ")
	result.AbsentFontSizes, result.AbsentFontSizeShape, err = nativeAbsentFontSizes(data)
	if err != nil {
		return nil, err
	}
	result.AbsentFontFamilies, err = nativeAbsentFontFamilies(data)
	if err != nil {
		return nil, err
	}
	result.LatinFontFallbacks, err = nativeLatinFontFallbacks(data)
	if err != nil {
		return nil, err
	}
	// What w:tblInd measures changed in Word's compatibilityMode 15. Measured
	// against Microsoft Word 16.112.4's own PDF exports of the rendering corpus,
	// a mode 12 or mode 14 package (and a package that attests no mode at all,
	// which this extractor reads as 12) places the leading cell's *content* at
	// the text margin plus w:tblInd, so the table's leading edge sits one left
	// cell margin further left; a mode 15 package places the table's leading
	// edge there instead. The controlled pairs are identical but for the mode:
	// tdf118812_tableStyles-comprehensive (no attestation) against
	// tdf118947_tableStyle (15), and Table_cell_auto_width_fdo69656 (14) against
	// fdo80800b_tableStyle (15) -- same TableGrid style, w:tblInd 0 and a 108
	// twip left cell margin in each, Word's leading border centred one cell
	// margin left of the text margin in the first of each pair and on the text
	// margin in the second. So the evidence is collected for every legacy mode
	// and withheld only from 15, whose current-layout origin this tier already
	// paints.
	if mode != 15 {
		result.LegacyTableOrigins, err = nativeLegacyTableOrigins(data)
		if err != nil {
			return nil, err
		}
	}
	return result, nil
}

// nativeApproximationBuilder accumulates typed facts. Every fact must join one
// strict diagnostic path, kinds and paths stay unique, and the fact count stays
// within the TS decoder bound; any violation refuses with a disclosed reason.
type nativeApproximationBuilder struct {
	result    *NativeDocxApproximationEligibilityV1
	wordNS    string
	diagnosed map[string]map[string]bool
	kinds     map[string]bool
	paths     map[string]bool
}

func newNativeApproximationBuilder(result *NativeDocxApproximationEligibilityV1, wordNS string, diagnostics []NativePaginationSettingsDiagnosticV1) *nativeApproximationBuilder {
	builder := &nativeApproximationBuilder{result: result, wordNS: wordNS, diagnosed: map[string]map[string]bool{}, kinds: map[string]bool{}, paths: map[string]bool{}}
	for _, diagnostic := range diagnostics {
		if builder.diagnosed[diagnostic.Path] == nil {
			builder.diagnosed[diagnostic.Path] = map[string]bool{}
		}
		builder.diagnosed[diagnostic.Path][diagnostic.Code] = true
	}
	return builder
}

// refuse discloses the cause and drops every accumulated fact and its reason:
// an ineligible attestation must not carry partial typed facts.
func (b *nativeApproximationBuilder) refuse(detail string) bool {
	dropped := map[string]bool{}
	for _, fact := range b.result.ApproximatedSettings {
		dropped[nativeApproximationSettingReason(fact)] = true
	}
	kept := make([]string, 0, len(b.result.Reasons))
	for _, reason := range b.result.Reasons {
		if !dropped[reason] {
			kept = append(kept, reason)
		}
	}
	b.result.Reasons = append(kept, "Approximate eligibility refused: "+detail)
	b.result.ApproximatedSettings = nil
	return false
}

func (b *nativeApproximationBuilder) add(fact NativeDocxApproximatedSettingV1) bool {
	if b.diagnosed[fact.Path] == nil {
		return b.refuse("typed fact " + fact.Kind + " at " + fact.Path + " has no strict diagnostic to join")
	}
	if b.kinds[fact.Kind] || b.paths[fact.Path] {
		return b.refuse("typed fact " + fact.Kind + " at " + fact.Path + " repeats an already recorded kind or path")
	}
	if len(b.result.ApproximatedSettings) >= nativeApproximationMaxFacts {
		return b.refuse("more than " + strconv.Itoa(nativeApproximationMaxFacts) + " typed settings facts would be required")
	}
	b.kinds[fact.Kind], b.paths[fact.Path] = true, true
	b.result.ApproximatedSettings = append(b.result.ApproximatedSettings, fact)
	b.result.Reasons = append(b.result.Reasons, nativeApproximationSettingReason(fact))
	return true
}

func (b *nativeApproximationBuilder) val() xml.Name {
	return xml.Name{Space: b.wordNS, Local: "val"}
}

// topLevel records bounded known extras, active hyphenation, and authoring-only
// settings.xml children that strict parsing diagnosed. Children strict already
// consumed as attested-neutral (no diagnostic) never become facts.
func (b *nativeApproximationBuilder) topLevel(root *nativeXMLNode) bool {
	authoring := NativeDocxApproximatedSettingV1{Kind: "authoringSettings", Values: map[string]string{}}
	for _, child := range root.Children {
		codes := b.diagnosed[child.Path]
		// No diagnostic: strict consumed the child as attested-neutral. A
		// duplicate occurrence (index >= 2) is resolved only by duplicates(), so a
		// known extra such as w:themeFontLang[2] never becomes its own fact.
		if codes == nil || codes["DUPLICATE_SETTINGS_PROPERTY"] {
			continue
		}
		if fact := nativeApproximateSetting(child, b.wordNS); fact != nil {
			if !b.add(*fact) {
				return false
			}
			continue
		}
		if child.Name.Space != b.wordNS {
			continue
		}
		switch {
		case child.Name.Local == "autoHyphenation" && codes["PAGINATION_SETTING_UNSUPPORTED"]:
			fact, ok := b.hyphenationFact(root, child)
			if !ok {
				return b.refuse("autoHyphenation or its options are not exact value leaves")
			}
			if !b.add(fact) {
				return false
			}
		case child.Name.Local == "characterSpacingControl" && codes["CHARACTER_SPACING_CONTROL_UNSUPPORTED"]:
			// Recorded, never performed: the approximate tier shapes natural
			// advances, so East Asian punctuation compression is disclosed as a
			// typed fact instead of refusing the whole attestation.
			if !nativeExactLeaf(child, b.val()) {
				return b.refuse("characterSpacingControl at " + child.Path + " is not an exact value leaf")
			}
			value, present := nativeAttr(child, b.wordNS, "val")
			if !present || !nativeApproximateCharacterSpacingControl[value] {
				return b.refuse("characterSpacingControl at " + child.Path + " is not a known ECMA-376 compression value")
			}
			if !b.add(NativeDocxApproximatedSettingV1{Kind: "characterSpacingControl", Path: child.Path, Values: map[string]string{"val": value}}) {
				return false
			}
		case nativeApproximateAuthoringSettings[child.Name.Local] && codes["PAGINATION_SETTING_UNSUPPORTED"]:
			summary, ok := nativeApproximateAttributeSummary(child)
			if !ok || !nativeXMLWhitespaceOnly(child.Text) {
				return b.refuse("authoring setting " + child.Path + " has unbounded attributes or text content")
			}
			if len(authoring.Values) >= nativeApproximationMaxGroupMembers {
				return b.refuse("more than 64 authoring-only settings")
			}
			if authoring.Path == "" {
				authoring.Path = child.Path
			}
			authoring.Values[child.Path] = summary
		}
	}
	if len(authoring.Values) > 0 && !b.add(authoring) {
		return false
	}
	return true
}

func (b *nativeApproximationBuilder) hyphenationFact(root, node *nativeXMLNode) (NativeDocxApproximatedSettingV1, bool) {
	fact := NativeDocxApproximatedSettingV1{Kind: "autoHyphenation", Path: node.Path, Values: map[string]string{}}
	if !nativeExactLeaf(node, b.val()) {
		return fact, false
	}
	// An omitted w:val on an on/off leaf means "true" (ECMA-376 17.17.4).
	value, present := nativeAttr(node, b.wordNS, "val")
	if !present {
		value = "true"
	} else if _, valid := nativeLexicalOnOff(value); !valid {
		return fact, false
	}
	fact.Values["val"] = value
	for _, name := range []string{"hyphenationZone", "consecutiveHyphenLimit", "doNotHyphenateCaps"} {
		options := directNativeChildren(root, b.wordNS, name)
		if len(options) == 0 {
			continue
		}
		if !nativeExactLeaf(options[0], b.val()) {
			return fact, false
		}
		value, present := nativeAttr(options[0], b.wordNS, "val")
		if !present && name == "doNotHyphenateCaps" {
			value = "true"
		}
		if present || name == "doNotHyphenateCaps" {
			fact.Values[name] = value
		}
	}
	return fact, true
}

// duplicates admits each DUPLICATE_SETTINGS_PROPERTY diagnostic only when the
// duplicate exactly repeats an earlier same-named sibling (strict already used
// the first occurrence) or is an authoring-only top-level setting. A repeated
// w:compat container stays ambiguous.
func (b *nativeApproximationBuilder) duplicates(root *nativeXMLNode, diagnostics []NativePaginationSettingsDiagnosticV1) bool {
	group := NativeDocxApproximatedSettingV1{Kind: "duplicateSettings", Values: map[string]string{}}
	for _, diagnostic := range diagnostics {
		if diagnostic.Code != "DUPLICATE_SETTINGS_PROPERTY" {
			continue
		}
		node := nativeApproximateNodeByPath(root, diagnostic.Path)
		if node == nil || node.parent == nil || node.Name == (xml.Name{Space: b.wordNS, Local: "compat"}) {
			return b.refuse("duplicate setting " + diagnostic.Path + " is ambiguous")
		}
		canonical, ok := nativeApproximateCanonicalNode(node, 0)
		if !ok {
			return b.refuse("duplicate setting " + diagnostic.Path + " exceeds the bounded comparison size")
		}
		agrees := false
		for _, sibling := range node.parent.Children {
			if sibling == node {
				break
			}
			if sibling.Name != node.Name {
				continue
			}
			if earlier, ok := nativeApproximateCanonicalNode(sibling, 0); ok && earlier == canonical {
				agrees = true
				break
			}
		}
		authoringOnly := node.parent == root && node.Name.Space == b.wordNS && nativeApproximateAuthoringSettings[node.Name.Local]
		if !agrees && !authoringOnly {
			return b.refuse("duplicate setting " + diagnostic.Path + " disagrees with its first occurrence")
		}
		summary, ok := nativeApproximateAttributeSummary(node)
		if !ok {
			return b.refuse("duplicate setting " + diagnostic.Path + " has unbounded attributes")
		}
		if _, seen := group.Values[diagnostic.Path]; seen || len(group.Values) >= nativeApproximationMaxGroupMembers {
			return b.refuse("duplicate settings exceed the bounded group")
		}
		if group.Path == "" {
			group.Path = diagnostic.Path
		}
		group.Values[diagnostic.Path] = summary
	}
	if len(group.Values) > 0 && !b.add(group) {
		return false
	}
	return true
}

// compat resolves the legacy mode from agreeing compatibilityMode attestations
// and records every other compat child as a typed not-applied fact.
func (b *nativeApproximationBuilder) compat(root *nativeXMLNode, mode *int) bool {
	compat := directNativeChildren(root, b.wordNS, "compat")
	if len(compat) > 1 {
		return b.refuse("repeated w:compat containers are ambiguous")
	}
	if len(compat) == 0 {
		return true
	}
	if !nativeExactContainer(compat[0]) {
		return b.refuse("w:compat carries unexpected attributes or text")
	}
	modeValue := ""
	// Word and third-party writers sometimes emit the same non-mode compatSetting
	// more than once, occasionally with disagreeing values. None of those flags is
	// an input to current layout, so every repeat is recorded verbatim in one
	// grouped fact instead of refusing the attestation over an ambiguity that
	// cannot reach this tier's output. compatibilityMode is excluded: its value is
	// consumed, so disagreeing attestations still refuse below.
	repeated := NativeDocxApproximatedSettingV1{Kind: "repeatedCompatSettings", Values: map[string]string{}}
	for _, child := range compat[0].Children {
		codes := b.diagnosed[child.Path]
		if child.Name.Space != b.wordNS {
			return b.refuse("foreign compat markup at " + child.Path)
		}
		switch {
		case child.Name.Local == "noColumnBalance":
			// Strict consumed the first exact leaf; a repeat is resolved by duplicates.
			for code := range codes {
				if code != "DUPLICATE_SETTINGS_PROPERTY" {
					return b.refuse("noColumnBalance at " + child.Path + " is not an exact value leaf")
				}
			}
		case nativeApproximateLegacyCompatFlags[child.Name.Local]:
			if !nativeExactLeaf(child, b.val()) {
				return b.refuse("compat option at " + child.Path + " is not an exact value leaf")
			}
			fact := NativeDocxApproximatedSettingV1{Kind: child.Name.Local, Path: child.Path, Values: map[string]string{}}
			if value, present := nativeAttr(child, b.wordNS, "val"); present {
				if _, valid := nativeLexicalOnOff(value); !valid {
					return b.refuse("compat option at " + child.Path + " has an invalid on/off value")
				}
				fact.Values["val"] = value
			}
			if !b.add(fact) {
				return false
			}
		case child.Name.Local == "compatSetting":
			if !nativeExactLeaf(child, xml.Name{Space: b.wordNS, Local: "name"}, xml.Name{Space: b.wordNS, Local: "uri"}, b.val()) {
				return b.refuse("compatSetting at " + child.Path + " is not an exact leaf")
			}
			name, _ := nativeAttr(child, b.wordNS, "name")
			uri, _ := nativeAttr(child, b.wordNS, "uri")
			value, _ := nativeAttr(child, b.wordNS, "val")
			if uri != nativeApproximationCompatSettingURI {
				return b.refuse("compatSetting at " + child.Path + " uses an unknown uri")
			}
			if name == "compatibilityMode" {
				if value != "12" && value != "14" && value != "15" {
					return b.refuse("compatibilityMode at " + child.Path + " is not exactly 12, 14 or 15")
				}
				if modeValue != "" && modeValue != value {
					return b.refuse("compatibilityMode attestations disagree at " + child.Path)
				}
				modeValue = value
				// The leading compatSetting is the TS-covered legacy mode slot; any
				// diagnosed repeat or non-leading attestation needs its own fact.
				if codes != nil && !strings.HasSuffix(child.Path, "/w:compatSetting[1]") {
					if !b.add(NativeDocxApproximatedSettingV1{Kind: "compatibilityMode", Path: child.Path, Values: map[string]string{"val": value}}) {
						return false
					}
				}
				continue
			}
			if !nativeApproximateCompatSettingFlag(name, value) {
				return b.refuse("compatSetting " + name + " at " + child.Path + " is not a recorded flag value")
			}
			if b.kinds[name] {
				if len(repeated.Values) >= nativeApproximationMaxGroupMembers {
					return b.refuse("more than 64 repeated compatSetting attestations")
				}
				if b.diagnosed[child.Path] == nil {
					return b.refuse("repeated compatSetting at " + child.Path + " has no strict diagnostic to join")
				}
				if repeated.Path == "" {
					repeated.Path = child.Path
				}
				repeated.Values[child.Path] = name + "=" + value
				continue
			}
			if !b.add(NativeDocxApproximatedSettingV1{Kind: name, Path: child.Path, Values: map[string]string{"val": value}}) {
				return false
			}
		default:
			return b.refuse("unknown compat markup at " + child.Path)
		}
	}
	if len(repeated.Values) > 0 && !b.add(repeated) {
		return false
	}
	switch modeValue {
	case "14":
		*mode = 14
	case "15":
		*mode = 15
	}
	return true
}

func nativeApproximateNodeByPath(node *nativeXMLNode, path string) *nativeXMLNode {
	if node.Path == path {
		return node
	}
	if !strings.HasPrefix(path, node.Path+"/") {
		return nil
	}
	for _, child := range node.Children {
		if found := nativeApproximateNodeByPath(child, path); found != nil {
			return found
		}
	}
	return nil
}
