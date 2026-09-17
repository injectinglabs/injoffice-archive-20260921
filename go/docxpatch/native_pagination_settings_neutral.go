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

// Word-emitted settings.xml extras that ECMA-376 plus MS-OI29500 prove do not
// change native shaping or pagination. Admission is per element and per
// argument, never a category waiver: each case below names the clause and says
// why that setting cannot move a glyph, a line break, a margin or a page
// boundary, and each one still fails closed outside its bounded shape.
// Extra compatSetting flags remain refused: MS-DOCX 2.3.1 changes table-style
// hierarchy evaluation, and 2.3.3 enableOpenTypeFeatures defaults to disabled
// and turns on kerning, ligatures, and related OpenType layout. So do the
// settings that do reach layout - mirrorMargins, gutterAtTop, active
// autoHyphenation, compressing characterSpacingControl, a non-Latin
// themeFontLang - and the compatibilityMode attestation itself.
func nativeSettingsAttestedNeutralWordElement(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	switch node.Name.Local {
	case "decimalSymbol", "listSeparator":
		return nativeSettingsNeutralFieldSeparator(result, node, wordNS)
	case "themeFontLang":
		return nativeSettingsNeutralThemeFontLang(result, node, wordNS)
	case "shapeDefaults":
		return nativeSettingsNeutralShapeDefaults(result, node)
	case "hdrShapeDefaults":
		// ECMA-376 17.15.1.41: defaults applied when a NEW VML shape is created in
		// a header/footer story. Nothing already in the package is described by it,
		// so no anchor, extent or text box changes. Bounded to the same v:ext="edit"
		// authoring-identity subset as w:shapeDefaults; fill, stroke, style and
		// colour-history defaults that could paint still fail closed.
		return nativeSettingsNeutralShapeIdentityDefaults(result, node, false)
	case "activeWritingStyle":
		return nativeSettingsNeutralActiveWritingStyle(result, node, wordNS)
	case "bordersDoNotSurroundHeader", "bordersDoNotSurroundFooter", "doNotUseMarginsForDrawingGridOrigin", "removeDateAndTime", "uiCompat97To2003":
		// ECMA-376 17.15.1.7/17.15.1.8: the two border switches select only which
		// stories the decorative page-border rectangle encloses. Word paints that
		// rectangle over finished layout; the body text box, its margins and the
		// page boundary come from w:sectPr, never from these switches.
		// 17.15.1.34: the drawing-grid origin switch moves the origin of the
		// interactive snap grid from the text margins to the page edge. It is an
		// authoring aid; stored drawing anchors keep their own coordinates.
		// 17.15.1.71: removeDateAndTime strips date/time from annotations when Word
		// saves. Annotation metadata is not shaped into body advances.
		// 17.15.1.90: uiCompat97To2003 greys out Word UI commands for features
		// earlier versions lack. It disables authoring, not layout of stored content.
		return nativeSettingsNeutralOnOffSwitch(result, node, wordNS)
	case "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery", "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing", "drawingGridHorizontalOrigin", "drawingGridVerticalOrigin":
		// ECMA-376 17.15.1.28-17.15.1.33: the display interval, spacing and origin
		// of the DRAWING grid, the on-canvas aid Word snaps a shape to while a
		// person drags it. It is not the typographic grid: East Asian line pitch
		// comes from w:docGrid in w:sectPr and w:snapToGrid on the paragraph, both
		// resolved elsewhere and both unaffected here. No glyph, line box or page
		// boundary is derived from these measures.
		return nativeSettingsNeutralDrawingGridMeasure(result, node, wordNS)
	default:
		return false
	}
}

// nativeSettingsNeutralOnOffSwitch admits an exact ECMA-376 17.17.4 on/off leaf.
// An omitted w:val means true; a non-lexical value still fails closed.
func nativeSettingsNeutralOnOffSwitch(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	if !nativeSettingsExactLeaf(result, node, xml.Name{Space: wordNS, Local: "val"}) {
		return false
	}
	if raw, present := nativeAttr(node, wordNS, "val"); present {
		if _, valid := nativeLexicalOnOff(raw); !valid {
			result.addDiagnostic("INVALID_SETTINGS_ON_OFF", node, "Layout-neutral on/off setting has an invalid lexical value")
			return false
		}
	}
	return true
}

// nativeSettingsNeutralDrawingGridMeasure requires the bounded non-negative
// ST_TwipsMeasure/ST_DecimalNumber value the schema declares. A missing or
// unbounded value is not the attested subset.
func nativeSettingsNeutralDrawingGridMeasure(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	if !nativeSettingsExactLeaf(result, node, xml.Name{Space: wordNS, Local: "val"}) {
		return false
	}
	raw, present := nativeAttr(node, wordNS, "val")
	if value, err := strconv.ParseUint(raw, 10, 32); !present || err != nil || value > 1_000_000_000 {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "Drawing grid measure requires a non-negative bounded integer")
		return false
	}
	return true
}

// nativeSettingsNeutralActiveWritingStyle admits the proofing registration of
// ECMA-376 17.15.1.1: which grammar/spelling engine Word last ran for one
// language. No font, metric or box is derived from it. The element is
// maxOccurs="unbounded" by schema, one per language, so repeats are its own
// shape rather than an ambiguity (see parseNativePaginationSettings).
func nativeSettingsNeutralActiveWritingStyle(result *NativePaginationSettingsV1, node *nativeXMLNode, wordNS string) bool {
	allowed := make([]xml.Name, 0, 6)
	for _, local := range []string{"appName", "lang", "vendorID", "dllVersion", "nlCheck", "checkStyle"} {
		allowed = append(allowed, xml.Name{Space: wordNS, Local: local})
	}
	if !nativeSettingsExactLeaf(result, node, allowed...) {
		return false
	}
	if _, ok := nativeAttr(node, wordNS, "lang"); !ok {
		result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "activeWritingStyle requires its w:lang registration")
		return false
	}
	return true
}

func nativeSettingsAttestedNeutralForeignElement(result *NativePaginationSettingsV1, node *nativeXMLNode) bool {
	if node.Name == (xml.Name{Space: nativeMathNamespace, Local: "mathPr"}) {
		return nativeSettingsNeutralMathPr(result, node)
	}
	// MS-DOCX 2.5.1.2 w14:defaultImageDpi is the target resolution Word applies
	// when it COMPRESSES a picture on insert or save. Every drawing already in the
	// package carries its own wp:extent/a:ext, which is what native layout reads,
	// so this value never rescales a painted picture box.
	if node.Name == (xml.Name{Space: nativeWord14Namespace, Local: "defaultImageDpi"}) {
		if !nativeSettingsExactLeaf(result, node, xml.Name{Space: nativeWord14Namespace, Local: "val"}) {
			return false
		}
		raw, present := nativeAttr(node, nativeWord14Namespace, "val")
		if value, err := strconv.ParseUint(raw, 10, 32); !present || err != nil || value > 1_000_000 {
			result.addDiagnostic("INVALID_SETTINGS_STRUCTURE", node, "defaultImageDpi requires a non-negative bounded integer")
			return false
		}
		return true
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
	return nativeSettingsNeutralShapeIdentityDefaults(result, node, true)
}

// nativeSettingsNeutralShapeIdentityDefaults admits the v:ext="edit" authoring
// identity subset shared by w:shapeDefaults and w:hdrShapeDefaults. body pins
// the original w:shapeDefaults contract exactly: both children present and the
// main-story idmap seed. The header story omits o:shapelayout in most Word
// output and seeds its own idmap, so only its bounded shape is required. Either
// way the value is an authoring id counter: nothing reads it during layout.
func nativeSettingsNeutralShapeIdentityDefaults(result *NativePaginationSettingsV1, node *nativeXMLNode, body bool) bool {
	if !nativeSettingsExactNode(result, node, map[xml.Name]bool{}, true) {
		return false
	}
	if len(node.Children) != 2 && (body || len(node.Children) != 1) {
		return false
	}
	ext := xml.Name{Space: nativeVMLNamespace, Local: "ext"}
	defaults := node.Children[0]
	if defaults.Name != (xml.Name{Space: nativeOfficeNamespace, Local: "shapedefaults"}) ||
		!nativeSettingsExactLeaf(result, defaults, ext, xml.Name{Local: "spidmax"}) {
		return false
	}
	identities := []*nativeXMLNode{defaults}
	if len(node.Children) == 2 {
		layout := node.Children[1]
		if layout.Name != (xml.Name{Space: nativeOfficeNamespace, Local: "shapelayout"}) ||
			!nativeSettingsExactNode(result, layout, map[xml.Name]bool{ext: true}, true) ||
			len(layout.Children) != 1 {
			return false
		}
		idmap := layout.Children[0]
		if idmap.Name != (xml.Name{Space: nativeOfficeNamespace, Local: "idmap"}) ||
			!nativeSettingsExactLeaf(result, idmap, ext, xml.Name{Local: "data"}) {
			return false
		}
		data, _ := nativeAttr(idmap, "", "data")
		if (body && data != "1") || !nativeSettingsBoundedShapeIdentity(data) {
			return false
		}
		identities = append(identities, layout, idmap)
	}
	for _, child := range identities {
		if value, _ := nativeAttr(child, nativeVMLNamespace, "ext"); value != "edit" {
			return false
		}
	}
	spid, _ := nativeAttr(defaults, "", "spidmax")
	return nativeSettingsBoundedShapeIdentity(spid)
}

// nativeSettingsBoundedShapeIdentity accepts one canonical positive VML shape
// identity counter, refusing padding, signs and out-of-range values.
func nativeSettingsBoundedShapeIdentity(raw string) bool {
	value, err := strconv.ParseUint(raw, 10, 31)
	return err == nil && value != 0 && strconv.FormatUint(value, 10) == raw
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
