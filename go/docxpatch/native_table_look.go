package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// ECMA-376 17.4.54: tblLook selects conditional components of a table style.
// A source-bound, complete chain without such components makes the switches
// inactive for rendering, not safe for mutation. No style-name heuristic is used.
func (extractor *nativeExtractor) inactiveTableLook(properties, look *nativeXMLNode) bool {
	if !nativeExactContainer(properties) || !nativeExactLeaf(look, xml.Name{Space: extractor.wordNS, Local: "val"}) {
		return false
	}
	value, ok := nativeAttr(look, extractor.wordNS, "val")
	if !ok || len(value) != 4 {
		return false
	}
	for _, digit := range value {
		if !(digit >= '0' && digit <= '9' || digit >= 'A' && digit <= 'F') {
			return false
		}
	}
	mask, err := strconv.ParseUint(value, 16, 16)
	if err != nil || mask & ^uint64(0x07e0) != 0 {
		return false
	}
	// Qualify the canonical transitional bitmask only. Other representations
	// remain preserved until their independent override rules are modeled.
	styles := directNativeChildren(properties, extractor.wordNS, "tblStyle")
	if len(styles) != 1 || len(directNativeChildren(properties, extractor.wordNS, "tblLook")) != 1 || !nativeExactLeaf(styles[0], xml.Name{Space: extractor.wordNS, Local: "val"}) {
		return false
	}
	styleID, ok := nativeAttr(styles[0], extractor.wordNS, "val")
	if !ok || !nativeIDPattern.MatchString(styleID) {
		return false
	}
	extractor.loadTableLookStyles()
	seen := map[string]bool{}
	for depth := 0; depth < 64; depth++ {
		if seen[styleID] {
			return false
		}
		seen[styleID] = true
		style := extractor.tableLookStyles[styleID]
		if style == nil {
			return false
		}
		if !nativeExactContainer(style, xml.Name{Space: extractor.wordNS, Local: "type"}, xml.Name{Space: extractor.wordNS, Local: "styleId"}, xml.Name{Space: extractor.wordNS, Local: "default"}, xml.Name{Space: extractor.wordNS, Local: "customStyle"}) {
			return false
		}
		for _, attr := range []string{"default", "customStyle"} {
			if _, present := nativeAttr(style, extractor.wordNS, attr); present {
				if _, valid := nativeOnOffAttr(style, extractor.wordNS, attr, false); !valid {
					return false
				}
			}
		}
		for _, child := range style.Children {
			if child.Name.Space != extractor.wordNS || child.Name.Local == "tblStylePr" {
				return false
			}
			if !nativeTableStyleAuthoringLocal(child.Name.Local) && child.Name.Local != "pPr" && child.Name.Local != "rPr" && child.Name.Local != "tblPr" && child.Name.Local != "tcPr" && child.Name.Local != "trPr" {
				return false
			}
		}
		parents := directNativeChildren(style, extractor.wordNS, "basedOn")
		if len(parents) == 0 {
			return true
		}
		if len(parents) != 1 || !nativeExactLeaf(parents[0], xml.Name{Space: extractor.wordNS, Local: "val"}) {
			return false
		}
		styleID, ok = nativeAttr(parents[0], extractor.wordNS, "val")
		if !ok || !nativeIDPattern.MatchString(styleID) {
			return false
		}
	}
	return false
}

func (extractor *nativeExtractor) loadTableLookStyles() {
	if extractor.tableLookLoaded {
		return
	}
	extractor.tableLookLoaded = true
	resolver := nativeLayoutResolver{pkg: extractor.pkg, mainPart: extractor.mainPart, wordNS: extractor.wordNS, relBase: extractor.relNS + "/"}
	part, err := resolver.singletonRelatedPart("styles")
	if err != nil || part == "" || resolver.validateRelatedPartContentType("styles", part) != nil {
		return
	}
	root, err := parseNativeXML(part, extractor.pkg.files[part])
	if err != nil || root.Name != (xml.Name{Space: extractor.wordNS, Local: "styles"}) || !nativeExactContainer(root) || len(root.Children) > NativeDOCXMaxCollectionItems {
		return
	}
	styles := map[string]*nativeXMLNode{}
	for _, child := range root.Children {
		if child.Name.Space != extractor.wordNS {
			return
		}
		if child.Name.Local != "style" {
			continue
		}
		kind, _ := nativeAttr(child, extractor.wordNS, "type")
		if kind != "table" {
			continue
		}
		id, ok := nativeAttr(child, extractor.wordNS, "styleId")
		if !ok || !nativeIDPattern.MatchString(id) || styles[id] != nil {
			return
		}
		styles[id] = child
	}
	extractor.tableLookStyles = styles
}
