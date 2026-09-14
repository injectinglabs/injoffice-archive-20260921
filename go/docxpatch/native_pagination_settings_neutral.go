package docxpatch

import (
	"encoding/xml"
	"strconv"
	"unicode/utf8"
)

const (
	nativeOfficeNamespace = "urn:schemas-microsoft-com:office:office"
	nativeVMLNamespace    = "urn:schemas-microsoft-com:vml"
)

// Word-emitted settings.xml extras from 2col-header.docx / ImageCrop.docx that
// ECMA-376 plus MS-OI29500 prove do not change native shaping or pagination.
// Extra compatSetting flags remain refused: MS-DOCX 2.3.1 changes table-style
// hierarchy evaluation, and 2.3.3 enableOpenTypeFeatures defaults to disabled
// and turns on kerning, ligatures, and related OpenType layout.
func nativeSettingsAttestedNeutralWordElement(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	switch node.Name.Local {
	case "decimalSymbol", "listSeparator":
		return nativeSettingsNeutralFieldSeparator(result, node, wordNS)
	case "themeFontLang":
		return nativeSettingsNeutralThemeFontLang(result, node, wordNS)
	case "shapeDefaults":
		return nativeSettingsNeutralShapeDefaults(result, node)
	default:
		return false
	}
}

func nativeSettingsAttestedNeutralForeignElement(result *NativePaginationSettingsV1, node *nativeXMLNode) bool {
	if node.Name == (xml.Name{Space: nativeMathNamespace, Local: "mathPr"}) {
		return nativeSettingsNeutralMathPr(result, node)
	}
	return false
}

func nativeSettingsNeutralFieldSeparator(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	// ECMA-376 17.15.1.22 / 17.15.1.56: these characters are used only while
	// evaluating field instructions. Enabled updateFields is already refused,
	// so native pagination consumes source text rather than re-evaluated fields.
	if !nativeSettingsExactLeaf(result, node, xml.Name{Space: wordNS, Local: "val"}) {
		return false
	}
	raw, ok := nativeAttr(node, wordNS, "val")
	if !ok || utf8.RuneCountInString(raw) != 1 {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "Field separator requires exactly one character")
		return false
	}
	if node.Name.Local == "decimalSymbol" {
		return raw == "." || raw == ","
	}
	return raw == "," || raw == ";"
}

func nativeSettingsNeutralThemeFontLang(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	// ECMA-376 17.15.1.88: val/eastAsia/bidi select theme fonts by language.
	// Omitted themeFontLang uses a:latin/ea/cs. Latn en-US/fr-FR with omitted or
	// x-none eastAsia/bidi do not select a non-latin script font; ja-JP and other
	// languages remain layout-affecting. MS-OI29500 2.1.423 only describes omitted
	// host-language defaults, which this explicit Latn subset does not invoke.
	if !nativeSettingsExactLeaf(result, node,
		xml.Name{Space: wordNS, Local: "val"},
		xml.Name{Space: wordNS, Local: "eastAsia"},
		xml.Name{Space: wordNS, Local: "bidi"},
	) {
		return false
	}
	val, ok := nativeAttr(node, wordNS, "val")
	if !ok {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "themeFontLang requires w:val")
		return false
	}
	if val != "en-US" && val != "fr-FR" {
		return false
	}
	for _, local := range []string{"eastAsia", "bidi"} {
		if raw, present := nativeAttr(node, wordNS, local); present && raw != "x-none" {
			return false
		}
	}
	return true
}

func nativeSettingsNeutralShapeDefaults(result *NativePaginationSettingsV1, node *nativeXMLNode) bool {
	// ECMA-376 17.15.1.79: shapeDefaults are defaults for creating new shapes.
	// The attested Word subset is v:ext="edit" authoring identity (spidmax/idmap)
	// and does not include fill, stroke, or textbox defaults that could paint.
	if !nativeSettingsExactNode(result, node, map[xml.Name]bool{}, true) || len(node.Children) != 2 {
		return false
	}
	defaults, layout := node.Children[0], node.Children[1]
	if defaults.Name != (xml.Name{Space: nativeOfficeNamespace, Local: "shapedefaults"}) ||
		!nativeSettingsExactLeaf(result, defaults, xml.Name{Space: nativeVMLNamespace, Local: "ext"}, xml.Name{Local: "spidmax"}) {
		return false
	}
	if layout.Name != (xml.Name{Space: nativeOfficeNamespace, Local: "shapelayout"}) ||
		!nativeSettingsExactNode(result, layout, map[xml.Name]bool{{Space: nativeVMLNamespace, Local: "ext"}: true}, true) ||
		len(layout.Children) != 1 {
		return false
	}
	idmap := layout.Children[0]
	if idmap.Name != (xml.Name{Space: nativeOfficeNamespace, Local: "idmap"}) ||
		!nativeSettingsExactLeaf(result, idmap, xml.Name{Space: nativeVMLNamespace, Local: "ext"}, xml.Name{Local: "data"}) {
		return false
	}
	for _, child := range []*nativeXMLNode{defaults, layout, idmap} {
		if value, _ := nativeAttr(child, nativeVMLNamespace, "ext"); value != "edit" {
			return false
		}
	}
	spid, _ := nativeAttr(defaults, "", "spidmax")
	value, err := strconv.ParseUint(spid, 10, 31)
	if err != nil || value == 0 || strconv.FormatUint(value, 10) != spid {
		return false
	}
	data, _ := nativeAttr(idmap, "", "data")
	return data == "1"
}

func nativeSettingsNeutralMathPr(result *NativePaginationSettingsV1, node *nativeXMLNode) bool {
	// ECMA-376 22.1.2.62: document mathPr apply to OMML. The Word-emitted default
	// 11-child subset matches omitted mathPr (Cambria Math, wrapIndent 1440 twips,
	// centerGroup display, subSup/undOvr limits). smallFrac must be explicitly
	// off; dispDef must be on (omitted val or 1/true/on). Inverted booleans
	// change nested fraction size and display-equation indentation.
	if !nativeSettingsExactNode(result, node, map[xml.Name]bool{}, true) {
		return false
	}
	allowed := map[string]string{
		"mathFont": "Cambria Math", "brkBin": "before", "brkBinSub": "--",
		"lMargin": "0", "rMargin": "0", "defJc": "centerGroup",
		"wrapIndent": "1440", "intLim": "subSup", "naryLim": "undOvr",
	}
	seen := map[string]string{}
	for _, child := range node.Children {
		if child.Name.Space != nativeMathNamespace || !nativeSettingsExactLeaf(result, child, xml.Name{Space: nativeMathNamespace, Local: "val"}) {
			return false
		}
		key := child.Name.Local
		if _, duplicate := seen[key]; duplicate {
			return false
		}
		value, present := nativeAttr(child, nativeMathNamespace, "val")
		switch key {
		case "smallFrac":
			enabled, valid := nativeLexicalOnOff(value)
			if !present || !valid || enabled {
				return false
			}
		case "dispDef":
			if present {
				enabled, valid := nativeLexicalOnOff(value)
				if !valid || !enabled {
					return false
				}
			}
		default:
			if expected, exists := allowed[key]; !exists || !present || value != expected {
				return false
			}
		}
		seen[key] = value
	}
	return len(seen) == 11
}
