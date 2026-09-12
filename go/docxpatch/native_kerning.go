package docxpatch

import "encoding/xml"

// ECMA-376 Part 1 17.3.2.19: kern is an inherited minimum font size in
// half-points, inclusive at sz == kern. This tier accepts bounded whole units.
func nativeKerningThreshold(node *nativeXMLNode, ns string) (int, bool) {
	value, ok := nativePositiveIntAttr(node, ns, "val")
	return value, ok && value <= 3276 && nativeExactLeaf(node, xml.Name{Space: ns, Local: "val"})
}
