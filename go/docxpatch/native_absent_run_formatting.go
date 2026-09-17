package docxpatch

import "encoding/xml"

// One sentence for both the extractor and the resolver: they read the same
// values from the same markup, so they state the same fact.
const nativeAbsentRunFormattingMessage = "This run property states its own absence, so it selects the glyphs, advances and line box this tier already produces and moves no line and no page"

// Microsoft Word writes a full CT_RPr of explicit "off" leaves whenever direct
// character formatting is cleared, and the numbering and style galleries ship
// the same block verbatim. Every leaf recognised here states the *absence* of a
// formatting operation: no case transform, no strike, no relief, no animation,
// no emphasis mark, zero character spacing, zero baseline offset and no
// shading. Applying it is applying nothing, which is exactly the glyph run,
// advance and line box this tier already produces, so it selects no different
// page. It is the same reading nativeAbsentRunBorder already applies to a
// w:bdr that states no border.
//
// The acceptance is per-property AND per-value, like the w14:ligatures and
// w14:cntxtAlts acceptances before it. The enabled form of each of these
// properties keeps its existing refusal in the same switch, so an explicit off
// can never override an enabled ancestor this tier applied: a reachable
// enabled caps/smallCaps/strike/shd still refuses the page. Malformed,
// decorated, repeated or unknown-valued markup states something this reading
// does not cover and stays refused.
//
// Measured with harfbuzzjs 1.6.0 against the 32 faces of the host font
// manifest, shaping the enabled effect of each property against the shaper
// defaults over Latin, Arabic and Hebrew samples:
//   - w:caps: uppercasing the run changes glyph ids and advances on 124 of 224
//     face/sample pairs (Calibri 400 regular +13,836 units at 2048 upem).
//   - w:smallCaps (smcp=1): 65 of 224 differ (Calibri 400 regular +7,722).
//   - w:spacing: a nonzero value is literal letter spacing in twips, which the
//     canonical shaper refuses outright ("nonzero letter/word spacing is not
//     qualified for canonical cluster-edge placement").
//   - w:position: a nonzero value raises or lowers the baseline.
//
// Their zero/off values apply none of that, so they are identical to the
// shaping this tier already performs by construction rather than by a
// coincidence of one face.
func nativeAbsentRunFormatting(node, owner *nativeXMLNode, ns string) bool {
	if node.Name.Space != ns || len(directNativeChildren(owner, ns, node.Name.Local)) != 1 {
		return false
	}
	value := xml.Name{Space: ns, Local: "val"}
	switch node.Name.Local {
	// CT_OnOff case, strike and relief toggles. Their enabled form either
	// reshapes the run (caps, smallCaps) or paints ink this tier does not
	// paint (strike, dstrike, outline, shadow, emboss, imprint, specVanish);
	// the disabled form does neither.
	case "caps", "smallCaps", "strike", "dstrike", "outline", "shadow", "emboss", "imprint", "specVanish":
		if !nativeExactLeaf(node, value) {
			return false
		}
		enabled, valid := nativeOnOff(node, ns)
		return valid && !enabled
	// ST_TextEffect names a Word text animation; "none" selects no animation.
	case "effect":
		if !nativeExactLeaf(node, value) {
			return false
		}
		selected, present := nativeAttr(node, ns, "val")
		return present && selected == "none"
	// ST_Em names an East Asian emphasis mark; "none" selects no mark.
	case "em":
		if !nativeExactLeaf(node, value) {
			return false
		}
		selected, present := nativeAttr(node, ns, "val")
		return present && selected == "none"
	// Character spacing in twips and baseline offset in half-points. Only the
	// exact literal zero qualifies: "0" is the identity, "00" or "+0" is
	// markup this reading does not cover.
	case "spacing", "position":
		if !nativeExactLeaf(node, value) {
			return false
		}
		selected, present := nativeAttr(node, ns, "val")
		return present && selected == "0"
	// CT_Shd with w:val="clear" selects the empty pattern, so w:color paints
	// nothing and only w:fill remains; an "auto" fill is the inherited page
	// background this tier already paints behind every run.
	case "shd":
		if !nativeExactLeaf(node, value, xml.Name{Space: ns, Local: "color"}, xml.Name{Space: ns, Local: "fill"},
			xml.Name{Space: ns, Local: "themeColor"}, xml.Name{Space: ns, Local: "themeTint"}, xml.Name{Space: ns, Local: "themeShade"},
			xml.Name{Space: ns, Local: "themeFill"}, xml.Name{Space: ns, Local: "themeFillTint"}, xml.Name{Space: ns, Local: "themeFillShade"}) {
			return false
		}
		if _, themed := nativeAttr(node, ns, "themeFill"); themed {
			return false
		}
		selected, present := nativeAttr(node, ns, "val")
		fill, filled := nativeAttr(node, ns, "fill")
		return present && selected == "clear" && filled && fill == "auto"
	}
	return false
}
