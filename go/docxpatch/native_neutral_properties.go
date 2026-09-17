package docxpatch

import "encoding/xml"

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
