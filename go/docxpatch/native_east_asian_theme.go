package docxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// East-Asian theme font slots.
//
// ECMA-376 17.3.2.26 lets a run name its East-Asian font by theme slot
// (w:eastAsiaTheme) instead of by family, and DrawingML's fontScheme answers
// that slot. Office's own themes ship an EMPTY <a:ea typeface=""/> and put the
// real answer in the script table that follows it. Which row of that table
// applies is not a guess: ECMA-376 17.15.1.87 says w:themeFontLang "specifies
// the languages which shall be used to determine the theme fonts", and its
// w:eastAsia attribute carries the East-Asian language. The chain is
// deterministic:
//
//	docDefaults <w:rFonts w:eastAsiaTheme="minorEastAsia"/>
//	  -> theme <a:minorFont><a:ea typeface=""/>        (empty: read the table)
//	  -> settings <w:themeFontLang w:eastAsia="zh-TW"/>
//	  -> script "Hant"
//	  -> <a:font script="Hant" typeface="新細明體"/>   (PMingLiU)
//
// Anything outside this exact chain - a duplicated slot, a script tag this map
// does not carry, an absent themeFontLang, a spoofed namespace - resolves to
// nothing and keeps the run's script font deferred exactly as before.

// nativeEastAsianThemeScript maps the bounded set of BCP-47 tags Word writes
// into w:themeFontLang/@w:eastAsia onto the fontScheme script attribute.
// Unlisted tags are not guessed.
func nativeEastAsianThemeScript(tag string) (string, bool) {
	if tag == "" || len(tag) > 32 {
		return "", false
	}
	parts := strings.Split(strings.ToLower(tag), "-")
	switch parts[0] {
	case "ja":
		return "Jpan", true
	case "ko":
		return "Hang", true
	case "zh":
		for _, sub := range parts[1:] {
			switch sub {
			case "hant", "tw", "hk", "mo":
				return "Hant", true
			case "hans", "cn", "sg":
				return "Hans", true
			}
		}
		// ECMA-376 and Word both treat a bare zh as Simplified Chinese.
		return "Hans", true
	}
	return "", false
}

// nativeThemeScriptTag bounds the fontScheme script attribute to an ISO 15924
// shaped tag before it is used as a lookup key.
func nativeThemeScriptTag(value string) bool {
	if len(value) != 4 {
		return false
	}
	for _, character := range value {
		if (character < 'A' || character > 'Z') && (character < 'a' || character > 'z') {
			return false
		}
	}
	return true
}

// nativeExactThemeSlotTypeface reads one <a:latin>/<a:ea>/<a:cs> slot. The
// second result reports that the slot is exactly authored; an authored slot may
// still carry the empty typeface Office writes, which defers to the script
// table rather than naming a face.
func nativeExactThemeSlotTypeface(fontSet *nativeXMLNode, drawingNS, local string) (string, bool) {
	if fontSet == nil {
		return "", false
	}
	slot := nativeUniqueThemeFontChild(fontSet, drawingNS, local)
	if slot == nil || !nativeExactLeaf(slot,
		xml.Name{Local: "typeface"}, xml.Name{Space: drawingNS, Local: "typeface"},
		xml.Name{Local: "panose"}, xml.Name{Space: drawingNS, Local: "panose"},
	) {
		return "", false
	}
	value, ok := nativeDrawingAttr(slot, drawingNS, "typeface")
	if !ok {
		return "", false
	}
	if value == "" {
		return "", true
	}
	if !nativeBoundedResolvedString(value, 256) {
		return "", false
	}
	return value, true
}

// nativeParseThemeScriptFonts reads the <a:font script="..." typeface="..."/>
// rows of one font set. A duplicated script, a spoofed namespace or an
// unreadable row has no single authored answer, so the whole table is dropped.
func nativeParseThemeScriptFonts(fontSet *nativeXMLNode, drawingNS string) map[string]string {
	if fontSet == nil {
		return nil
	}
	fonts := map[string]string{}
	for _, child := range fontSet.Children {
		if child.Name.Local != "font" {
			continue
		}
		if child.Name.Space != drawingNS || !nativeExactLeaf(child,
			xml.Name{Local: "script"}, xml.Name{Space: drawingNS, Local: "script"},
			xml.Name{Local: "typeface"}, xml.Name{Space: drawingNS, Local: "typeface"},
			xml.Name{Local: "panose"}, xml.Name{Space: drawingNS, Local: "panose"},
			xml.Name{Local: "pitchFamily"}, xml.Name{Space: drawingNS, Local: "pitchFamily"},
			xml.Name{Local: "charset"}, xml.Name{Space: drawingNS, Local: "charset"},
		) {
			return nil
		}
		script, hasScript := nativeDrawingAttr(child, drawingNS, "script")
		typeface, hasTypeface := nativeDrawingAttr(child, drawingNS, "typeface")
		if !hasScript || !hasTypeface || !nativeThemeScriptTag(script) || typeface == "" || !nativeBoundedResolvedString(typeface, 256) {
			return nil
		}
		if _, duplicate := fonts[script]; duplicate {
			return nil
		}
		fonts[script] = typeface
	}
	return fonts
}

// resolveEastAsiaThemeTypeface answers one w:eastAsiaTheme slot reference.
func (resolver *nativeLayoutResolver) resolveEastAsiaThemeTypeface(theme string) (string, bool) {
	fonts := resolver.themeLatinFonts
	var typeface string
	var authored bool
	var scripts map[string]string
	switch theme {
	case "majorEastAsia":
		typeface, authored, scripts = fonts.majorEastAsia, fonts.majorEastAsiaAuthored, fonts.majorScripts
	case "minorEastAsia":
		typeface, authored, scripts = fonts.minorEastAsia, fonts.minorEastAsiaAuthored, fonts.minorScripts
	default:
		return "", false
	}
	if !authored {
		return "", false
	}
	if typeface != "" {
		return typeface, true
	}
	script, ok := nativeEastAsianThemeScript(resolver.themeFontLangEastAsia)
	if !ok {
		return "", false
	}
	face, ok := scripts[script]
	return face, ok && face != ""
}

// loadSettingsThemeFontLang reads the one layout-affecting settings property
// this resolver needs. Everything else in settings.xml stays the pagination
// settings extractor's business.
func (resolver *nativeLayoutResolver) loadSettingsThemeFontLang(partName string) error {
	root, err := parseNativeXML(partName, resolver.pkg.files[partName])
	if err != nil {
		return err
	}
	if root.Name != (xml.Name{Space: resolver.wordNS, Local: "settings"}) {
		return fmt.Errorf("docxpatch: native style resolution: settings part %q has spoofed or invalid root", partName)
	}
	var found *nativeXMLNode
	for _, child := range root.Children {
		if child.Name.Local != "themeFontLang" {
			continue
		}
		if child.Name.Space != resolver.wordNS || found != nil {
			return nil
		}
		found = child
	}
	if found == nil || !nativeExactLeaf(found,
		xml.Name{Space: resolver.wordNS, Local: "val"},
		xml.Name{Space: resolver.wordNS, Local: "eastAsia"},
		xml.Name{Space: resolver.wordNS, Local: "bidi"},
	) {
		return nil
	}
	if value, present := nativeAttr(found, resolver.wordNS, "eastAsia"); present && nativeScriptLanguageTag(value) {
		resolver.themeFontLangEastAsia = value
	}
	return nil
}

// nativeEastAsianSlotRune reports whether MS-OI29500 17.3.2.26 assigns a rune
// to the East-Asian font slot with no dependence on w:hint. The ambiguous
// ranges this tier already routes through ascii/hAnsi (Latin-1 Supplement,
// Latin Extended-A/B, General Punctuation) are deliberately not here.
func nativeEastAsianSlotRune(character rune) bool {
	switch {
	case character >= 0x1100 && character <= 0x11ff: // Hangul Jamo
	case character >= 0x2e80 && character <= 0x2fdf: // CJK and Kangxi Radicals
	case character >= 0x3000 && character <= 0x30ff: // CJK punctuation, Hiragana, Katakana
	case character >= 0x3100 && character <= 0x312f: // Bopomofo
	case character >= 0x3130 && character <= 0x318f: // Hangul Compatibility Jamo
	case character >= 0x31c0 && character <= 0x31ff: // CJK strokes, Katakana extensions
	case character >= 0x3200 && character <= 0x4dbf: // Enclosed CJK, compatibility, extension A
	case character >= 0x4e00 && character <= 0x9fff: // CJK Unified Ideographs
	case character >= 0xac00 && character <= 0xd7af: // Hangul Syllables
	case character >= 0xf900 && character <= 0xfaff: // CJK Compatibility Ideographs
	case character >= 0xfe30 && character <= 0xfe4f: // CJK Compatibility Forms
	case character >= 0xff00 && character <= 0xffef: // Halfwidth and Fullwidth Forms
	case character >= 0x20000 && character <= 0x2fa1f: // Supplementary Ideographic Plane
	default:
		return false
	}
	return true
}

// nativeScriptSlotUse records which font slots a run's own text actually uses.
type nativeScriptSlotUse struct {
	// eastAsia: at least one rune MS-OI29500 assigns to the East-Asian slot.
	eastAsia bool
	// hintBound: at least one rune whose slot w:hint would change.
	hintBound bool
	// unmodelled: at least one rune this tier assigns to no slot, which keeps
	// every deferred script property blocking exactly as before.
	unmodelled bool
}

func (use nativeScriptSlotUse) merge(other nativeScriptSlotUse) nativeScriptSlotUse {
	return nativeScriptSlotUse{eastAsia: use.eastAsia || other.eastAsia, hintBound: use.hintBound || other.hintBound, unmodelled: use.unmodelled || other.unmodelled}
}

func nativeClassifyScriptSlots(text string) nativeScriptSlotUse {
	use := nativeScriptSlotUse{}
	for _, character := range text {
		switch {
		case !nativeRequiresScriptShaping(character):
			// Basic Latin never depends on w:hint; the remaining ascii/hAnsi
			// ranges do, so a hint over them is still not resolved here.
			if character > 0x7f {
				use.hintBound = true
			}
		case nativeEastAsianSlotRune(character):
			use.eastAsia = true
		default:
			use.unmodelled = true
		}
	}
	return use
}
