package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// NativeDOCXMaxCharacterSpacingTwips bounds the character tracking this tier
// resolves. ECMA-376 states no bound on w:spacing, but Word's own Character
// Spacing control writes at most 1584 pt of condensing or expanding, and a
// tracking delta wider than that cannot describe type on a page. Anything
// beyond it is preserved as unmodeled markup rather than guessed at.
const NativeDOCXMaxCharacterSpacingTwips = 1584 * 20

// ECMA-376 Part 1 17.3.2.35: w:spacing on w:rPr is character tracking, one
// ST_SignedTwipsMeasure added to the advance of every character of the run.
// This tier resolves it only in its exact whole-twip form: one element, one
// w:val attribute, no other attributes and no nested markup, and a nonzero
// value inside the bound above. A zero is handled by nativeAbsentRunEffect
// before this is reached, because it adds nothing and states so.
//
// The measure may be signed: a negative value condenses. The caller converts
// twips to the shared milli-point unit (1 twip = 50 milli-points), so 15 twips
// is exactly 0.75 pt.
func nativeCharacterSpacingTwips(node, owner *nativeXMLNode, ns string) (int, bool) {
	if node.Name.Space != ns || len(directNativeChildren(owner, ns, node.Name.Local)) != 1 {
		return 0, false
	}
	if !nativeExactLeaf(node, xml.Name{Space: ns, Local: "val"}) {
		return 0, false
	}
	raw, ok := nativeAttr(node, ns, "val")
	if !ok {
		return 0, false
	}
	value, err := strconv.Atoi(raw)
	if err != nil || value == 0 || value < -NativeDOCXMaxCharacterSpacingTwips || value > NativeDOCXMaxCharacterSpacingTwips {
		return 0, false
	}
	// Reject a value Word itself would not have written, so that "+15" or "015"
	// is preserved rather than silently normalised into painted geometry.
	if strconv.Itoa(value) != raw {
		return 0, false
	}
	return value, true
}
