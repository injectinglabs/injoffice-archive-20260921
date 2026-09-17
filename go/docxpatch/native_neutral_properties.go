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
