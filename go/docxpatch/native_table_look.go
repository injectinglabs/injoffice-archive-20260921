package docxpatch

import (
	"encoding/xml"
	"strconv"
)

// ECMA-376 17.4.56 w:tblLook: every switch it carries selects one conditional
// component of the applied table style. The element holds no geometry itself,
// so it is inactive for rendering exactly when the conditional components it
// selects do not exist. Two independent proofs of that are accepted, in order:
// the package defines no conditional table formatting at all, or the table's
// own resolved style chain defines none. Neither makes the source safe to
// mutate; no style-name heuristic is used.
func (extractor *nativeExtractor) inactiveTableLook(properties, look *nativeXMLNode) bool {
	if !nativeExactContainer(properties) || len(directNativeChildren(properties, extractor.wordNS, "tblLook")) != 1 {
		return false
	}
	if !extractor.qualifiedTableLookSwitches(look) {
		return false
	}
	extractor.loadTableLookStyles()
	// A package whose styles part defines no w:tblStylePr offers no conditional
	// component for any switch to select, whatever the applied style resolves
	// to -- including an absent w:tblStyle and a reference to an undefined one.
	if extractor.noTblStylePrProven {
		return true
	}
	styles := directNativeChildren(properties, extractor.wordNS, "tblStyle")
	if len(styles) != 1 || !nativeExactLeaf(styles[0], xml.Name{Space: extractor.wordNS, Local: "val"}) {
		return false
	}
	styleID, ok := nativeAttr(styles[0], extractor.wordNS, "val")
	if !ok || !nativeIDPattern.MatchString(styleID) {
		return false
	}
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

// nativeTableLookBits is the ECMA-376 17.4.56 ST_ShortHexNumber layout of
// w:val. Word writes the same six switches twice, once packed into w:val and
// once as the named attributes; the two forms must agree or the element's
// selection is ambiguous and stays unmodeled.
var nativeTableLookBits = []struct {
	attr string
	bit  uint64
}{
	{"firstRow", 0x0020},
	{"lastRow", 0x0040},
	{"firstColumn", 0x0080},
	{"lastColumn", 0x0100},
	{"noHBand", 0x0200},
	{"noVBand", 0x0400},
}

func (extractor *nativeExtractor) qualifiedTableLookSwitches(look *nativeXMLNode) bool {
	allowed := []xml.Name{{Space: extractor.wordNS, Local: "val"}}
	for _, switchBit := range nativeTableLookBits {
		allowed = append(allowed, xml.Name{Space: extractor.wordNS, Local: switchBit.attr})
	}
	if !nativeExactLeaf(look, allowed...) {
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
	// Qualify the canonical transitional bitmask only. Other representations
	// remain preserved until their independent override rules are modeled.
	mask, err := strconv.ParseUint(value, 16, 16)
	if err != nil || mask & ^uint64(0x07e0) != 0 {
		return false
	}
	for _, switchBit := range nativeTableLookBits {
		if _, present := nativeAttr(look, extractor.wordNS, switchBit.attr); !present {
			continue
		}
		on, valid := nativeOnOffAttr(look, extractor.wordNS, switchBit.attr, false)
		if !valid || on != (mask&switchBit.bit != 0) {
			return false
		}
	}
	return true
}

func (extractor *nativeExtractor) loadTableLookStyles() {
	if extractor.tableLookLoaded {
		return
	}
	extractor.tableLookLoaded = true
	resolver := nativeLayoutResolver{pkg: extractor.pkg, mainPart: extractor.mainPart, wordNS: extractor.wordNS, relBase: extractor.relNS + "/"}
	part, err := resolver.singletonRelatedPart("styles")
	if err != nil {
		return
	}
	if part == "" {
		// The style definitions this package resolves against do not exist, so
		// it carries no conditional table formatting to select.
		extractor.noTblStylePrProven = true
		return
	}
	if resolver.validateRelatedPartContentType("styles", part) != nil {
		return
	}
	root, err := parseNativeXML(part, extractor.pkg.files[part])
	// mc:Ignorable only declares ignorable namespaces, exactly as the main part
	// extractor and the document-defaults reader already accept on part roots;
	// it declares no style and no conditional component.
	if err != nil || root.Name != (xml.Name{Space: extractor.wordNS, Local: "styles"}) || !nativeExactContainer(root, xml.Name{Space: nativeMCNamespace, Local: "Ignorable"}) || len(root.Children) > NativeDOCXMaxCollectionItems {
		return
	}
	extractor.noTblStylePrProven = !nativeSubtreeDeclaresConditionalTableFormat(root)
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

// nativeSubtreeDeclaresConditionalTableFormat reports whether any descendant
// declares a conditional table format. The local name is matched in every
// namespace so that markup this extractor does not otherwise model -- a
// compatibility alternate, a foreign wrapper -- cannot hide one.
func nativeSubtreeDeclaresConditionalTableFormat(node *nativeXMLNode) bool {
	for _, child := range node.Children {
		if child.Name.Local == "tblStylePr" || nativeSubtreeDeclaresConditionalTableFormat(child) {
			return true
		}
	}
	return false
}
