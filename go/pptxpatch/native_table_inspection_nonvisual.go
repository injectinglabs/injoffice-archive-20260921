package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

// PowerPoint's observed slide scaffold stores all-zero group coordinates.
// It is not an authored frame transform; no nonzero or unknown geometry is
// accepted as a substitute for the source table's own frame coordinates.
func inspectTableRootGroup(node *nativeXMLNode, d nativeExtractDialect) bool {
	if requireOnlyNativeAttrs(node) != nil || !onlyNativeXMLSpace(node.Text) {
		return false
	}
	if len(node.Children) == 0 {
		return true
	}
	if len(node.Children) != 1 {
		return false
	}
	x := node.Children[0]
	if x.Name != (xml.Name{Space: d.drawing, Local: "xfrm"}) || requireOnlyNativeAttrs(x) != nil || !onlyNativeXMLSpace(x.Text) || len(x.Children) != 4 {
		return false
	}
	for i, name := range []string{"off", "ext", "chOff", "chExt"} {
		n := x.Children[i]
		if n.Name != (xml.Name{Space: d.drawing, Local: name}) || requireOnlyNativeChildren(n) != nil {
			return false
		}
		a, b := "x", "y"
		if i%2 == 1 {
			a, b = "cx", "cy"
		}
		if requireOnlyNativeAttrs(n, xml.Name{Local: a}, xml.Name{Local: b}) != nil {
			return false
		}
		av, aok := exactNativeAttr(n, "", a)
		bv, bok := exactNativeAttr(n, "", b)
		if !aok || !bok || av != "0" || bv != "0" {
			return false
		}
	}
	return true
}

// Unlike the strict native table projection this read-only inspector accepts
// exactly the source no-group lock and PowerPoint modification identifier.
// Unknown extensions, visibility controls and placeholders remain unavailable.
func inspectTableNonVisual(node *nativeXMLNode, d nativeExtractDialect) (string, error) {
	bad := func() (string, error) {
		return "", fmt.Errorf("table inspection: unqualified visibility or nonvisual source")
	}
	if requireOnlyNativeAttrs(node) != nil || requireOnlyNativeChildren(node, xml.Name{Space: d.presentation, Local: "cNvPr"}, xml.Name{Space: d.presentation, Local: "cNvGraphicFramePr"}, xml.Name{Space: d.presentation, Local: "nvPr"}) != nil {
		return bad()
	}
	c, err := nativeSingleton(node, d.presentation, "cNvPr", true)
	if err != nil {
		return "", err
	}
	g, err := nativeSingleton(node, d.presentation, "cNvGraphicFramePr", true)
	if err != nil {
		return "", err
	}
	n, err := nativeSingleton(node, d.presentation, "nvPr", true)
	if err != nil {
		return "", err
	}
	if len(node.Children) != 3 || node.Children[0] != c || node.Children[1] != g || node.Children[2] != n {
		return bad()
	}
	if requireOnlyNativeAttrs(c, xml.Name{Local: "id"}, xml.Name{Local: "name"}) != nil || requireOnlyNativeChildren(c) != nil {
		return bad()
	}
	id, err := canonicalNativeUnsignedID(c, "", "id", 1)
	if err != nil {
		return "", err
	}
	if requireOnlyNativeAttrs(g) != nil || requireOnlyNativeChildren(g, xml.Name{Space: d.drawing, Local: "graphicFrameLocks"}) != nil {
		return bad()
	}
	lock, err := nativeSingleton(g, d.drawing, "graphicFrameLocks", false)
	if err != nil {
		return "", err
	}
	if lock != nil {
		if requireOnlyNativeAttrs(lock, xml.Name{Local: "noGrp"}) != nil || requireOnlyNativeChildren(lock) != nil {
			return bad()
		}
		value, ok := exactNativeAttr(lock, "", "noGrp")
		if !ok || value != "1" {
			return bad()
		}
	}
	if requireOnlyNativeAttrs(n) != nil || requireOnlyNativeChildren(n, xml.Name{Space: d.presentation, Local: "extLst"}) != nil {
		return bad()
	}
	list, err := nativeSingleton(n, d.presentation, "extLst", false)
	if err != nil {
		return "", err
	}
	if list != nil {
		if requireOnlyNativeAttrs(list) != nil || len(list.Children) != 1 {
			return bad()
		}
		ext := list.Children[0]
		if ext.Name != (xml.Name{Space: d.presentation, Local: "ext"}) || requireOnlyNativeAttrs(ext, xml.Name{Local: "uri"}) != nil || len(ext.Children) != 1 || !onlyNativeXMLSpace(ext.Text) {
			return bad()
		}
		uri, ok := exactNativeAttr(ext, "", "uri")
		if !ok || uri != "{D42A27DB-BD31-4B8C-83A1-F6EECF244321}" {
			return bad()
		}
		mod := ext.Children[0]
		if mod.Name != (xml.Name{Space: "http://schemas.microsoft.com/office/powerpoint/2010/main", Local: "modId"}) || requireOnlyNativeAttrs(mod, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(mod) != nil {
			return bad()
		}
		if _, err := requiredCanonicalNativeTableInt(mod, "val", 0, 4294967295); err != nil {
			return bad()
		}
	}
	return "cNvPr-" + id, nil
}
