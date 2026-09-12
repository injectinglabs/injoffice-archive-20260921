package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// ECMA-376 Part 1 17.6.5: omitted type means default (no grid). Pitch
// metadata is retained for re-enabling the grid, not applied to current layout.
func nativeInactiveSectionGrid(node *nativeXMLNode, ns string) bool {
	if !nativeExactLeaf(node, xml.Name{Space: ns, Local: "type"}, xml.Name{Space: ns, Local: "linePitch"}, xml.Name{Space: ns, Local: "charSpace"}) {
		return false
	}
	if kind, present := nativeAttr(node, ns, "type"); present && kind != "default" {
		return false
	}
	for _, name := range []string{"linePitch", "charSpace"} {
		if raw, present := nativeAttr(node, ns, name); present {
			value, err := strconv.ParseInt(raw, 10, 32)
			if err != nil || strconv.FormatInt(value, 10) != raw {
				return false
			}
		}
	}
	return true
}
