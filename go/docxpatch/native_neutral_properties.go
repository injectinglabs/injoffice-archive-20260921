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
