package docxpatch

import (
	"encoding/xml"
	"strings"
)

// These source properties require preservation but no glyph/layout operation.
// Enabled East Asian autospace remains unsupported; this only recognizes an
// explicit false setting. Proofing preferences never change printed glyphs.
// ECMA-376 17.3.1.1: disabled adjustRightInd uses the authored right indent
// regardless of the section grid. Active grid adjustment remains unsupported.
func nativeNeutralSourceProperty(node, owner *nativeXMLNode, ns string) bool {
	if node.Name.Space != ns || !nativeExactRevisionContainer(owner, ns, "rsidRPr") || len(directNativeChildren(owner, ns, node.Name.Local)) != 1 || !nativeExactLeaf(node, xml.Name{Space: ns, Local: "val"}) {
		return false
	}
	value, valid := nativeOnOff(node, ns)
	if !valid {
		return false
	}
	switch node.Name.Local {
	case "noProof":
		return true
	case "autoSpaceDE", "autoSpaceDN", "adjustRightInd":
		return !value
	}
	return false
}

// w:webHidden hides a run in Word's Web Layout view only. Paginated layout
// draws it like any other run, so it selects no glyph, no advance and no break
// here; this tier models no Web Layout view, and records that as its own code
// rather than guessing that the run should be dropped. Only the exact CT_OnOff
// leaf qualifies: malformed, repeated or decorated markup states something this
// reading does not cover and stays refused.
func nativeWebLayoutHiddenRun(node, owner *nativeXMLNode, ns string) bool {
	if node.Name != (xml.Name{Space: ns, Local: "webHidden"}) || len(directNativeChildren(owner, ns, "webHidden")) != 1 {
		return false
	}
	if !nativeExactLeaf(node, xml.Name{Space: ns, Local: "val"}) {
		return false
	}
	_, valid := nativeOnOff(node, ns)
	return valid
}

// The Word 2010 extension namespace that carries w14:ligatures.
const nativeWordML2010 = "http://schemas.microsoft.com/office/word/2010/wordml"

// w14:ligatures w14:val="standardContextual" asks for the standard and
// contextual ligature sets and nothing else. The v1 shaper declares
// default_feature_policy 'harfbuzz-14.3.0-shape-defaults', under which
// HarfBuzz already applies liga, clig and calt to every run, so this one value
// selects exactly the shaping this tier already performs and moves no advance.
// Every other value (none, all, historical, discretional) names a feature set
// v1 has no input for, and stays refused as foreign markup.
func nativeShaperDefaultLigatureMode(node, owner *nativeXMLNode) bool {
	name := xml.Name{Space: nativeWordML2010, Local: "ligatures"}
	if node.Name != name || len(directNativeChildren(owner, nativeWordML2010, "ligatures")) != 1 {
		return false
	}
	value := xml.Name{Space: nativeWordML2010, Local: "val"}
	if !nativeExactLeaf(node, value) {
		return false
	}
	count, mode := 0, ""
	for _, attr := range node.Attrs {
		if attr.Name == value {
			count++
			mode = attr.Value
		}
	}
	return count == 1 && mode == "standardContextual"
}

// w14:cntxtAlts turns the OpenType contextual alternates feature on. The v1
// shaper declares default_feature_policy 'harfbuzz-14.3.0-shape-defaults',
// under which HarfBuzz already applies calt to every horizontal run, so an
// enabled contextual-alternates request selects exactly the shaping this tier
// already performs and moves no advance. Shaping the same text with calt=1 and
// with the shaper defaults returns byte-identical glyphs and advances on every
// manifest face measured, across Latin, Arabic and Hebrew.
//
// An explicitly disabled w14:val turns a default feature off, which v1 has no
// input for and which does move advances, so it stays refused as foreign
// markup along with malformed, decorated or repeated markup.
func nativeShaperDefaultContextualAlternates(node, owner *nativeXMLNode) bool {
	name := xml.Name{Space: nativeWordML2010, Local: "cntxtAlts"}
	if node.Name != name || len(directNativeChildren(owner, nativeWordML2010, "cntxtAlts")) != 1 {
		return false
	}
	value := xml.Name{Space: nativeWordML2010, Local: "val"}
	if !nativeExactLeaf(node, value) {
		return false
	}
	enabled, valid := nativeOnOffAttr(node, nativeWordML2010, "val", true)
	return valid && enabled
}

// w14:stylisticSets names the OpenType stylistic sets Word applies to the run,
// one w14:styleSet child per set. An element with no child names no set, so it
// asks for exactly the face's default glyph forms - the forms this tier already
// shapes with - and selects the same glyphs at the same advances as the same
// run without the element. A stylisticSets that does name a set is a different
// statement and is recorded by nativeUnappliedTypographicRunFeature instead.
func nativeAbsentStylisticSetRequest(node, owner *nativeXMLNode) bool {
	name := xml.Name{Space: nativeWordML2010, Local: "stylisticSets"}
	if node.Name != name || len(directNativeChildren(owner, nativeWordML2010, "stylisticSets")) != 1 {
		return false
	}
	return nativeExactContainer(node) && len(node.Children) == 0 && len(node.Attrs) == 0
}

// w:bdr states the border Word draws around a run. ECMA-376 17.18.2 gives
// ST_Border both a "none" and a "nil" member, and both state the same thing:
// no border. A run border that names one of them therefore selects exactly the
// absence an omitted w:bdr already states — Word paints no stroke for it and
// reserves no space around the run, so the run's glyphs, advances, line box
// and page position are the ones the same run without the element produces.
// The companion w:sz, w:space, w:color, theme and w:frame/w:shadow attributes
// describe a stroke that is never drawn, so none of them can move a line.
//
// Every other ST_Border value paints a stroke and reserves space on all four
// sides of the run, which v1 has no input for, so it stays unmodeled markup
// along with malformed, decorated or repeated w:bdr elements.
func nativeAbsentRunBorder(node, owner *nativeXMLNode, ns string) bool {
	if node.Name != (xml.Name{Space: ns, Local: "bdr"}) || len(directNativeChildren(owner, ns, "bdr")) != 1 {
		return false
	}
	value := xml.Name{Space: ns, Local: "val"}
	if !nativeExactLeaf(node, value,
		xml.Name{Space: ns, Local: "sz"}, xml.Name{Space: ns, Local: "space"}, xml.Name{Space: ns, Local: "color"},
		xml.Name{Space: ns, Local: "themeColor"}, xml.Name{Space: ns, Local: "themeTint"}, xml.Name{Space: ns, Local: "themeShade"},
		xml.Name{Space: ns, Local: "frame"}, xml.Name{Space: ns, Local: "shadow"}) {
		return false
	}
	count, declared := 0, ""
	for _, attr := range node.Attrs {
		if attr.Name == value {
			count++
			declared = attr.Value
		}
	}
	return count == 1 && (declared == "none" || declared == "nil")
}

// Word 2010 run-typography extensions this tier records and does not apply.
//
// Each one names an OpenType feature (or, for w14:props3d, a 3D paint effect)
// that v1 has no shaping or paint input for. Recording them as foreign markup
// refused the whole page; recording them under their own code lets the
// approximate tier paint the run with the feature UNAPPLIED and disclose the
// exact property it did not apply. The strict tier keeps refusing every one of
// them, because the painted advances are measurably not Word's.
//
// Measured with harfbuzzjs 1.6.0 over 224 face/sample pairs drawn from the
// manifest faces and Latin, Arabic and Hebrew samples: ss02 changes shaping for
// 20 pairs (Arial: +28,745 font units across the sample), ss04 for 6 (Calibri:
// +112), w14:numForm="oldStyle" for 68 and w14:numSpacing="proportional" for 68
// (a ten-digit Calibri run shapes 312 font units narrower at 2048 upem than the
// default advances this tier paints). So a page painted without them puts
// glyphs at advances that differ from Word's by those amounts, and that is what
// the returned message states.
//
// Only the closed set below qualifies. Any other foreign run property - and any
// markup in this set carrying structure outside its own namespace - stays
// FOREIGN_RUN_PROPERTY and keeps refusing on both tiers.
func nativeUnappliedTypographicRunFeature(node, owner *nativeXMLNode) (string, string, bool) {
	if node.Name.Space != nativeWordML2010 {
		return "", "", false
	}
	if len(directNativeChildren(owner, nativeWordML2010, node.Name.Local)) != 1 {
		return "", "", false
	}
	value := xml.Name{Space: nativeWordML2010, Local: "val"}
	single := func() (string, bool) {
		if !nativeExactLeaf(node, value) {
			return "", false
		}
		count, declared := 0, ""
		for _, attr := range node.Attrs {
			if attr.Name == value {
				count++
				declared = attr.Value
			}
		}
		return declared, count == 1
	}
	switch node.Name.Local {
	case "stylisticSets":
		if !nativeExactContainer(node) || len(node.Children) == 0 {
			return "", "", false
		}
		id := xml.Name{Space: nativeWordML2010, Local: "id"}
		sets := []string{}
		for _, child := range node.Children {
			if child.Name != (xml.Name{Space: nativeWordML2010, Local: "styleSet"}) || !nativeExactLeaf(child, id) {
				return "", "", false
			}
			count, declared := 0, ""
			for _, attr := range child.Attrs {
				if attr.Name == id {
					count++
					declared = attr.Value
				}
			}
			if count != 1 || !nativeStylisticSetID(declared) {
				return "", "", false
			}
			// Word names the sets ss01..ss20; print the authored id in that form.
			if len(declared) == 1 {
				declared = "0" + declared
			}
			sets = append(sets, "ss"+declared)
		}
		return "STYLISTIC_SET_UNAPPLIED", "Stylistic set request " + strings.Join(sets, ", ") + " is preserved and NOT applied: the run is painted with the face's default glyph forms, at advances that are measurably not Word's. Shaped with harfbuzzjs 1.6.0 over 224 face/sample pairs, ss02 changes 20 of them (Arial: +28,745 font units across the sample) and ss04 changes 6 (Calibri: +112)", true
	case "ligatures":
		mode, ok := single()
		if !ok || !nativeNamedLigatureMode(mode) || mode == "standardContextual" {
			return "", "", false
		}
		return "LIGATURE_MODE_UNAPPLIED", "Ligature mode " + mode + " is preserved and NOT applied: the run is painted with this tier's declared HarfBuzz shaping defaults (liga, clig, calt), so every ligature set this mode adds or removes is absent from the painted glyphs and their advances", true
	case "numForm":
		form, ok := single()
		if !ok || (form != "default" && form != "lining" && form != "oldStyle") {
			return "", "", false
		}
		return "NUMBER_FORM_UNAPPLIED", "Number form " + form + " is preserved and NOT applied: digits are painted in the face's default form. Shaped with harfbuzzjs 1.6.0 over 224 face/sample pairs, w14:numForm=\"oldStyle\" changes 68 of them, so the painted digit glyphs and advances are measurably not Word's", true
	case "numSpacing":
		spacing, ok := single()
		if !ok || (spacing != "default" && spacing != "proportional" && spacing != "tabular") {
			return "", "", false
		}
		return "NUMBER_SPACING_UNAPPLIED", "Number spacing " + spacing + " is preserved and NOT applied: digits are painted at the face's default advances. Shaped with harfbuzzjs 1.6.0 over 224 face/sample pairs, w14:numSpacing=\"proportional\" changes 68 of them, and a ten-digit Calibri run shapes 312 font units narrower at 2048 upem than the default advances painted here", true
	case "props3d":
		if !nativeWordML2010Subtree(node) {
			return "", "", false
		}
		return "TEXT_EFFECT_3D_UNAPPLIED", "The 3D text effect w14:props3d is preserved and NOT applied: the run's glyphs are painted flat, without the requested extrusion, bevel, contour or material ink. It selects no glyph and moves no advance, so the deviation is exactly the missing decoration - Word rasterises such a run into an image in its own PDF export", true
	case "glow", "shadow", "reflection", "textOutline", "scene3d":
		if !nativeWordML2010Subtree(node) {
			return "", "", false
		}
		return "TEXT_DECORATION_UNAPPLIED", "The Word 2010 text decoration w14:" + node.Name.Local + " is preserved and NOT applied: the run's glyphs are painted with the face's own outline and the run's own fill, without the requested decoration ink. Like w14:props3d it decorates glyphs that are already selected and already placed - it names no OpenType feature and carries no advance - so it moves no glyph, no line and no page, and the deviation is exactly the missing decoration", true
	case "cntxtAlts":
		// The enabled form states shaping this tier already performs and is
		// taken by nativeShaperDefaultContextualAlternates before this point.
		// A disabled one asks for calt OFF, which the declared shaping defaults
		// do apply and v1 has no input for, so the run IS painted with calt and
		// its advances are measurably not Word's wherever the face has a calt
		// rule for the text.
		enabled, valid := nativeOnOffAttr(node, nativeWordML2010, "val", true)
		if !nativeExactLeaf(node, value) || valid == false || enabled {
			return "", "", false
		}
		return "CONTEXTUAL_ALTERNATES_UNAPPLIED", "Disabled contextual alternates w14:cntxtAlts are preserved and NOT applied: this tier declares default_feature_policy 'harfbuzz-14.3.0-shape-defaults', under which HarfBuzz applies calt to every horizontal run, and v1 has no input for turning a default feature off. The run is painted WITH contextual alternates, so wherever the face carries a calt rule for this text the painted glyphs and advances are measurably not Word's", true
	}
	return "", "", false
}

// ECMA-376 ST_Ligatures. standardContextual is excluded by the caller because
// nativeShaperDefaultLigatureMode already states it as shaping this tier
// performs; the rest name feature sets v1 has no input for.
func nativeNamedLigatureMode(mode string) bool {
	switch mode {
	case "none", "standard", "contextual", "historical", "discretional",
		"standardContextual", "standardHistorical", "contextualHistorical",
		"standardDiscretional", "contextualDiscretional", "historicalDiscretional",
		"standardContextualHistorical", "standardContextualDiscretional",
		"standardHistoricalDiscretional", "contextualHistoricalDiscretional", "all":
		return true
	}
	return false
}

// w14:styleSet w14:id is ST_DecimalNumber; only a short unsigned decimal names
// a stylistic set, and the disclosure prints it as ssNN.
func nativeStylisticSetID(value string) bool {
	if len(value) == 0 || len(value) > 3 {
		return false
	}
	for _, digit := range value {
		if digit < '0' || digit > '9' {
			return false
		}
	}
	return true
}

// Every element of the subtree stays in the Word 2010 extension namespace, so
// no WordprocessingML content, drawing or field can hide inside a property the
// caller is about to record as entirely unapplied.
func nativeWordML2010Subtree(node *nativeXMLNode) bool {
	if node.Name.Space != nativeWordML2010 {
		return false
	}
	for _, child := range node.Children {
		if !nativeWordML2010Subtree(child) {
			return false
		}
	}
	return true
}

// What a reader of a painted page must be told when a negative w:line was
// dropped. ECMA-376 17.3.1.33 types w:line as ST_SignedTwipsMeasure, and Word
// reads a negative one as an EXACT line height of its absolute value: it
// compresses the lines and lets them overlap, ignoring the authored w:lineRule.
// This tier has no compressed line box (the approximate envelope already
// declares that in DOCX_APPROXIMATE_LINE_BOX_WARNING), so it drops the
// measurement and the paragraph keeps the line spacing it inherits from its
// style. The painted lines are therefore FURTHER APART than Word's, by the
// difference between the inherited spacing and |w:line|.
const negativeLineSpacingDisclosure = "A negative w:line measurement is preserved and NOT applied: Word reads it as an exact line height of its absolute value, compressing the lines until they overlap, and this tier has no compressed line box. The paragraph keeps the line spacing it inherits instead, so its painted lines sit FURTHER APART than Word's by the difference between that inherited spacing and the absolute authored value"

// A run property whose authored value states the absence of its own effect.
//
// Word writes these out in full when it saves a numbering level's or a style's
// run properties, one element per property, each carrying the value that means
// "off": ECMA-376 types w:caps (17.3.2.5), w:smallCaps (17.3.2.33), w:strike
// (17.3.2.37), w:dstrike (17.3.2.9) and w:specVanish (17.3.2.36) as CT_OnOff,
// so an explicit false asks for exactly the glyphs, advances and ink an omitted
// element already produces; w:spacing (17.3.2.35) adds its ST_SignedTwipsMeasure
// to every character advance, so 0 adds nothing; w:position (17.3.2.24) raises
// the baseline by its ST_SignedHpsMeasure, so 0 raises nothing; w:effect
// (17.3.2.11) names an animated text effect and w:em (17.3.2.12) an emphasis
// mark, and each has a "none" member that draws neither.
//
// So a run carrying only these values occupies the same box, selects the same
// glyphs and paints the same ink as the same run without them, and cannot move
// a line or a page. Every other value of these properties does change glyphs,
// advances or ink, and stays unmodeled markup along with malformed, decorated,
// repeated or attribute-decorated elements.
func nativeAbsentRunEffect(node, owner *nativeXMLNode, ns string) bool {
	if node.Name.Space != ns || len(directNativeChildren(owner, ns, node.Name.Local)) != 1 {
		return false
	}
	value := xml.Name{Space: ns, Local: "val"}
	if !nativeExactLeaf(node, value) {
		return false
	}
	count, declared := 0, ""
	for _, attr := range node.Attrs {
		if attr.Name == value {
			count++
			declared = attr.Value
		}
	}
	if count != 1 {
		return false
	}
	switch node.Name.Local {
	case "caps", "smallCaps", "strike", "dstrike", "specVanish":
		state, valid := nativeOnOff(node, ns)
		return valid && !state
	case "spacing", "position":
		return declared == "0"
	case "effect", "em":
		return declared == "none"
	}
	return false
}
