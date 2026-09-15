package docxpatch

import (
	"encoding/xml"
	"regexp"
	"sort"
	"strconv"
	"strings"
)

// NativeDocxApproximatedSettingV1 records source facts explicitly disregarded by
// the read-only current-layout policy, never by strict pagination or mutation.
type NativeDocxApproximatedSettingV1 struct {
	Kind   string            `json:"kind"`
	Path   string            `json:"path"`
	Values map[string]string `json:"values"`
}

const (
	// nativeApproximationMaxFacts mirrors the TS decoder bound on
	// approximated_settings; more typed facts fail closed instead of joining.
	nativeApproximationMaxFacts = 8
	// nativeApproximationMaxGroupMembers bounds grouped authoring/duplicate facts.
	nativeApproximationMaxGroupMembers = 64
	// nativeApproximationMaxSummaryBytes bounds one retained attribute summary.
	nativeApproximationMaxSummaryBytes = 512
	nativeApproximationLanguageTag     = `^[A-Za-z]{2,3}(-[A-Za-z0-9]{1,8})*$`
)

var nativeApproximationLanguageTagPattern = regexp.MustCompile(nativeApproximationLanguageTag)

// nativeApproximationSettingReason must stay byte-identical to the TS mirror in
// packages/docs/src/nativeApproximationSettingsV1.ts.
func nativeApproximationSettingReason(fact NativeDocxApproximatedSettingV1) string {
	switch {
	case fact.Kind == "autoHyphenation":
		return "Current-layout approximation records autoHyphenation at " + fact.Path + " as not applied; automatic hyphenation is not performed and Word line breaks may differ"
	case fact.Kind == "compatibilityMode":
		return "Current-layout approximation records the repeated or non-leading compatibilityMode attestation at " + fact.Path + "; its agreeing value is the disclosed legacy mode and Word layout may differ"
	case fact.Kind == "themeFontLang":
		return "Current-layout approximation records themeFontLang at " + fact.Path + " as not applied; language-driven theme font selection is not performed and Word font choice may differ"
	case fact.Kind == "authoringSettings":
		return "Current-layout approximation records " + strconv.Itoa(len(fact.Values)) + " authoring-only settings anchored at " + fact.Path + " as not applied; current layout does not consume them, so Word editing, proofing, grid, template, and display behavior may differ"
	case fact.Kind == "duplicateSettings":
		return "Current-layout approximation records " + strconv.Itoa(len(fact.Values)) + " duplicate settings anchored at " + fact.Path + " as not applied; each repeats its first occurrence or is authoring-only, and the first occurrence is used"
	case nativeApproximateLegacyCompatFlags[fact.Kind]:
		return "Current-layout approximation records legacy compatibility option " + fact.Kind + " at " + fact.Path + " as not applied; Word compatibility layout is not emulated"
	default:
		return "Current-layout approximation disregards " + fact.Kind + " at " + fact.Path + "; source values are retained and Word layout may differ"
	}
}

// nativeApproximateAuthoringSettings are ECMA-376 17.15.1 settings.xml children
// that describe editor, proofing, template, grid-snapping, protection or
// save-time state. None of them is an input to native shaping or pagination, so
// the current-layout preview records them as not applied instead of refusing.
var nativeApproximateAuthoringSettings = map[string]bool{
	"activeWritingStyle": true, "attachedSchema": true, "attachedTemplate": true, "captions": true, "clickAndTypeStyle": true,
	"defaultTableStyle": true, "displayBackgroundShape": true, "displayHorizontalDrawingGridEvery": true, "displayVerticalDrawingGridEvery": true,
	"docVars": true, "documentProtection": true, "doNotDemarcateInvalidXml": true, "doNotDisplayPageBoundaries": true, "doNotEmbedSmartTags": true,
	"doNotUseMarginsForDrawingGridOrigin": true, "drawingGridHorizontalOrigin": true, "drawingGridHorizontalSpacing": true,
	"drawingGridVerticalOrigin": true, "drawingGridVerticalSpacing": true, "embedSystemFonts": true, "embedTrueTypeFonts": true,
	"forceUpgrade": true, "formsDesign": true, "hdrShapeDefaults": true, "hideGrammaticalErrors": true, "hideSpellingErrors": true,
	"linkStyles": true, "mailMerge": true, "noPunctuationKerning": true, "printFormsData": true, "printPostScriptOverText": true,
	"readModeInkLockDown": true, "removeDateAndTime": true, "removePersonalInformation": true, "rsids": true, "saveFormsData": true,
	"saveSubsetFonts": true, "saveThroughXslt": true, "saveXmlDataOnly": true, "schemaLibrary": true, "showEnvelope": true,
	"showXMLTags": true, "smartTagType": true, "stylePaneFormatFilter": true, "summaryLength": true, "useXSLTWhenSaving": true,
	"view": true, "writeProtection": true, "alwaysMergeEmptyNamespace": true,
}

// nativeApproximateLegacyCompatFlags are the ECMA-376 Part 1 17.15.3 and Part 4
// Transitional w:compat option leaves. The approximate preview already declares
// current InjOffice layout rather than Word compatibility layout, so each present
// option becomes a typed "not applied" fact instead of a refusal. compatSetting
// and noColumnBalance are parsed separately.
var nativeApproximateLegacyCompatFlags = map[string]bool{
	"adjustLineHeightInTable": true, "alignTablesRowByRow": true, "allowSpaceOfSameStyleInTable": true, "applyBreakingRules": true,
	"autofitToFirstFixedWidthCell": true, "autoSpaceLikeWord95": true, "balanceSingleByteDoubleByteWidth": true, "cachedColBalance": true,
	"convMailMergeEsc": true, "displayHangulFixedWidth": true, "doNotAutofitConstrainedTables": true, "doNotBreakConstrainedForcedTable": true,
	"doNotBreakWrappedTables": true, "doNotExpandShiftReturn": true, "doNotLeaveBackslashAlone": true, "doNotSnapToGridInCell": true,
	"doNotSuppressIndentation": true, "doNotSuppressParagraphBorders": true, "doNotUseEastAsianBreakRules": true, "doNotUseHTMLParagraphAutoSpacing": true,
	"doNotUseIndentAsNumberingTabStop": true, "doNotVertAlignCellWithSp": true, "doNotVertAlignInTxbx": true, "doNotWrapTextWithPunct": true,
	"footnoteLayoutLikeWW8": true, "forgetLastTabAlignment": true, "growAutofit": true, "layoutRawTableWidth": true, "layoutTableRowsApart": true,
	"lineWrapLikeWord6": true, "mwSmallCaps": true, "noExtraLineSpacing": true, "noLeading": true, "noSpaceRaiseLower": true, "noTabHangInd": true,
	"printBodyTextBeforeHeader": true, "printColBlack": true, "selectFldWithFirstOrLastChar": true, "shapeLayoutLikeWW8": true, "showBreaksInFrames": true,
	"spaceForUL": true, "spacingInWholePoints": true, "splitPgBreakAndParaMark": true, "subFontBySize": true, "suppressBottomSpacing": true,
	"suppressSpacingAtTopOfPage": true, "suppressSpBfAfterPgBrk": true, "suppressTopSpacing": true, "suppressTopSpacingWP": true, "swapBordersFacingPages": true,
	"truncateFontHeightsLikeWP6": true, "uiCompat97To2003": true, "ulTrailSpace": true, "underlineTabInNumList": true, "useAltKinsokuLineBreakRules": true,
	"useAnsiKerningPairs": true, "useFELayout": true, "useNormalStyleForList": true, "usePrinterMetrics": true, "useSingleBorderforContiguousCells": true,
	"useWord2002TableStyleRules": true, "useWord97LineBreakRules": true, "wpJustification": true, "wpSpaceWidth": true, "wrapTrailSpaces": true,
}

// nativeApproximateCompatSettingFlag lists the Microsoft compatSetting flags the
// current-layout policy records as typed facts. Word 2010+ emits the first four
// with val="1"; Word 2013+ adds the hyphenation and floating-table flags, whose
// 0/1 values do not change InjOffice current layout, which never hyphenates.
func nativeApproximateCompatSettingFlag(name, value string) bool {
	switch name {
	case "overrideTableStyleFontSizeAndJustification", "enableOpenTypeFeatures", "doNotFlipMirrorIndents", "differentiateMultirowTableHeaders":
		return value == "1"
	case "useWord2013TrackBottomHyphenation", "allowHyphenationAtTrackBottom", "allowTextAfterFloatingTableBreak":
		return value == "0" || value == "1"
	default:
		return false
	}
}

// nativeApproximateAttributeSummary retains an element's own attributes as one
// deterministic string: sorted qualified names with exact source values.
// Namespace declarations are omitted. Oversized summaries fail closed.
func nativeApproximateAttributeSummary(node *nativeXMLNode) (string, bool) {
	parts := make([]string, 0, len(node.Attrs))
	for _, attr := range node.Attrs {
		if nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		name := attr.Name.Local
		if attr.Name.Space != "" {
			name = nativeXMLQName(attr.Name)
		}
		parts = append(parts, name+"="+strconv.Quote(attr.Value))
	}
	sort.Strings(parts)
	summary := strings.Join(parts, " ")
	if len(summary) > nativeApproximationMaxSummaryBytes {
		return "", false
	}
	return summary, true
}

// nativeApproximateCanonicalNode serializes a settings subtree for exact
// duplicate comparison. Whitespace-only text is ignored; depth and size are bounded.
func nativeApproximateCanonicalNode(node *nativeXMLNode, depth int) (string, bool) {
	if depth > 8 {
		return "", false
	}
	summary, ok := nativeApproximateAttributeSummary(node)
	if !ok {
		return "", false
	}
	var out strings.Builder
	out.WriteString(nativeXMLQName(node.Name))
	out.WriteString("{")
	out.WriteString(summary)
	if !nativeXMLWhitespaceOnly(node.Text) {
		out.WriteString("|text=")
		out.WriteString(strconv.Quote(node.Text))
	}
	for _, child := range node.Children {
		nested, ok := nativeApproximateCanonicalNode(child, depth+1)
		if !ok {
			return "", false
		}
		out.WriteString("|")
		out.WriteString(nested)
		if out.Len() > 8*nativeApproximationMaxSummaryBytes {
			return "", false
		}
	}
	out.WriteString("}")
	return out.String(), true
}

func nativeApproximateSetting(node *nativeXMLNode, wordNS string) *NativeDocxApproximatedSettingV1 {
	result := &NativeDocxApproximatedSettingV1{Kind: node.Name.Local, Path: node.Path, Values: map[string]string{}}
	if node.Name.Space == wordNS {
		switch node.Name.Local {
		case "themeFontLang":
			// Theme font language selects theme faces per script. Current layout
			// records any well-formed BCP 47 tag (or an empty/x-none script slot)
			// and does not perform language-driven theme font selection.
			if !nativeExactLeaf(node, xml.Name{Space: wordNS, Local: "val"}, xml.Name{Space: wordNS, Local: "eastAsia"}, xml.Name{Space: wordNS, Local: "bidi"}) {
				return nil
			}
			for _, key := range []string{"val", "eastAsia", "bidi"} {
				if value, present := nativeAttr(node, wordNS, key); present {
					if !nativeApproximateLanguageSlot(value, key != "val") {
						return nil
					}
					result.Values[key] = value
				}
			}
			if _, present := result.Values["val"]; !present {
				return nil
			}
		case "decimalSymbol", "listSeparator":
			if !nativeExactLeaf(node, xml.Name{Space: wordNS, Local: "val"}) {
				return nil
			}
			value, _ := nativeAttr(node, wordNS, "val")
			if (node.Name.Local == "decimalSymbol" && value != "." && value != ",") || (node.Name.Local == "listSeparator" && value != "," && value != ";") {
				return nil
			}
			result.Values["val"] = value
		case "shapeDefaults":
			const officeNS = "urn:schemas-microsoft-com:office:office"
			const vmlNS = "urn:schemas-microsoft-com:vml"
			if !nativeExactContainer(node) || len(node.Children) != 2 {
				return nil
			}
			defaults, layout := node.Children[0], node.Children[1]
			if defaults.Name != (xml.Name{Space: officeNS, Local: "shapedefaults"}) || !nativeExactLeaf(defaults, xml.Name{Space: vmlNS, Local: "ext"}, xml.Name{Local: "spidmax"}) || layout.Name != (xml.Name{Space: officeNS, Local: "shapelayout"}) || !nativeExactContainer(layout, xml.Name{Space: vmlNS, Local: "ext"}) || len(layout.Children) != 1 {
				return nil
			}
			idmap := layout.Children[0]
			if idmap.Name != (xml.Name{Space: officeNS, Local: "idmap"}) || !nativeExactLeaf(idmap, xml.Name{Space: vmlNS, Local: "ext"}, xml.Name{Local: "data"}) {
				return nil
			}
			for _, child := range []*nativeXMLNode{defaults, layout, idmap} {
				if value, _ := nativeAttr(child, vmlNS, "ext"); value != "edit" {
					return nil
				}
			}
			spid, _ := nativeAttr(defaults, "", "spidmax")
			value, err := strconv.ParseUint(spid, 10, 31)
			if err != nil || value == 0 || strconv.FormatUint(value, 10) != spid {
				return nil
			}
			if data, _ := nativeAttr(idmap, "", "data"); data != "1" {
				return nil
			}
			result.Values["spidmax"], result.Values["idmap"] = spid, "1"
		default:
			return nil
		}
		return result
	}
	if node.Name != (xml.Name{Space: nativeMathNamespace, Local: "mathPr"}) || !nativeExactContainer(node) {
		return nil
	}
	allowed := map[string]string{"mathFont": "Cambria Math", "brkBin": "before", "brkBinSub": "--", "lMargin": "0", "rMargin": "0", "defJc": "centerGroup", "wrapIndent": "1440", "intLim": "subSup", "naryLim": "undOvr"}
	for _, child := range node.Children {
		if child.Name.Space != nativeMathNamespace || !nativeExactLeaf(child, xml.Name{Space: nativeMathNamespace, Local: "val"}) {
			return nil
		}
		key := child.Name.Local
		if _, seen := result.Values[key]; seen {
			return nil
		}
		value, present := nativeAttr(child, nativeMathNamespace, "val")
		if key == "smallFrac" || key == "dispDef" {
			if !present {
				value = "true"
			} else if _, valid := nativeLexicalOnOff(value); !valid {
				return nil
			}
		} else if expected, exists := allowed[key]; !exists || !present || value != expected {
			return nil
		}
		result.Values[key] = value
	}
	if len(result.Values) != 11 {
		return nil
	}
	return result
}

// nativeApproximateLanguageSlot accepts a well-formed BCP 47-shaped tag; the
// eastAsia/bidi script slots may also be empty or the Word x-none sentinel.
func nativeApproximateLanguageSlot(value string, scriptSlot bool) bool {
	if scriptSlot && (value == "" || value == "x-none") {
		return true
	}
	return len(value) <= 35 && nativeApproximationLanguageTagPattern.MatchString(value)
}
