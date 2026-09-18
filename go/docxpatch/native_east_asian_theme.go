package docxpatch

import (
	"encoding/xml"
	"fmt"
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
// Anything outside this exact chain - a duplicated slot or script row, a
// language this bounded map does not carry, an absent themeFontLang, a spoofed
// namespace - resolves to nothing and keeps the run's script font deferred.
//
// The language is known before the theme is read, so only the one script row it
// selects is kept; the rest of the table is never modelled.

// nativeEastAsianThemeScript maps the bounded set of BCP-47 tags Word writes
// into w:themeFontLang/@w:eastAsia onto the fontScheme script attribute.
// Unlisted tags are not guessed. Matching is ASCII-case-insensitive, which is
// all a BCP-47 subtag can be.
func nativeEastAsianThemeScript(tag string) string {
	if len(tag) < 2 || len(tag) > 32 {
		return ""
	}
	primary, rest := nativeASCIIFold(tag[:2]), ""
	if len(tag) > 2 {
		if tag[2] != '-' {
			return ""
		}
		rest = nativeASCIIFold(tag[3:])
	}
	switch primary {
	case "ja":
		return "Jpan"
	case "ko":
		return "Hang"
	case "zh":
		for start := 0; start <= len(rest); {
			end := start
			for end < len(rest) && rest[end] != '-' {
				end++
			}
			switch rest[start:end] {
			case "hant", "tw", "hk", "mo":
				return "Hant"
			case "hans", "cn", "sg":
				return "Hans"
			}
			start = end + 1
		}
		// ECMA-376 and Word both treat a bare zh as Simplified Chinese.
		return "Hans"
	}
	return ""
}

// nativeExactThemeTypeface reads one <a:latin>/<a:ea>/<a:cs> slot. The second
// result reports that the slot is exactly authored; an authored slot may still
// carry the empty typeface Office writes, which defers to the script table
// rather than naming a face.
func nativeExactThemeTypeface(fontSet *nativeXMLNode, drawingNS, local string) (string, bool) {
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
	if !ok || value != "" && !nativeBoundedResolvedString(value, 256) {
		return "", false
	}
	return value, true
}

// nativeThemeScriptTypeface reads the one <a:font script="..."/> row the
// document's own East-Asian language selects. A duplicated row, a spoofed
// namespace or an unreadable row has no single authored answer, so the whole
// table is dropped rather than searched past.
func nativeThemeScriptTypeface(fontSet *nativeXMLNode, drawingNS, script string) string {
	if fontSet == nil || script == "" {
		return ""
	}
	answer := ""
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
			return ""
		}
		rowScript, hasScript := nativeDrawingAttr(child, drawingNS, "script")
		typeface, hasTypeface := nativeDrawingAttr(child, drawingNS, "typeface")
		if !hasScript || !hasTypeface {
			return ""
		}
		if rowScript != script {
			continue
		}
		if answer != "" || typeface == "" || !nativeBoundedResolvedString(typeface, 256) {
			return ""
		}
		answer = typeface
	}
	return answer
}

// resolveEastAsiaThemeTypeface answers one w:eastAsiaTheme slot reference.
func (resolver *nativeLayoutResolver) resolveEastAsiaThemeTypeface(theme string) (string, bool) {
	fonts := resolver.themeLatinFonts
	slot := fonts.minorEastAsia
	switch theme {
	case "majorEastAsia":
		slot = fonts.majorEastAsia
	case "minorEastAsia":
	default:
		return "", false
	}
	return slot.typeface, slot.authored && slot.typeface != ""
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
		resolver.themeEastAsiaScript = nativeEastAsianThemeScript(value)
	}
	return nil
}

// nativeScriptSlotUse records which font slots a run's own text actually uses.
type nativeScriptSlotUse struct {
	// eastAsia: at least one rune MS-OI29500 17.3.2.26 assigns to the
	// East-Asian slot with no dependence on w:hint.
	eastAsia bool
	// complex: at least one rune MS-OI29500 17.3.2.26 assigns to the
	// complex-script slot, narrowed to the scripts this tier's pinned shaper
	// qualifies. See nativeClassifyScriptSlots.
	complex bool
	// hintBound: at least one rune whose slot w:hint would change.
	hintBound bool
	// unmodelled: at least one rune this tier assigns to no slot, which keeps
	// every deferred script property blocking exactly as before.
	unmodelled bool
}

func (use nativeScriptSlotUse) merge(other nativeScriptSlotUse) nativeScriptSlotUse {
	return nativeScriptSlotUse{use.eastAsia || other.eastAsia, use.complex || other.complex, use.hintBound || other.hintBound, use.unmodelled || other.unmodelled}
}

// The ambiguous ranges this tier already routes through ascii/hAnsi (Latin-1
// Supplement, Latin Extended-A/B, General Punctuation) are deliberately not
// East-Asian here; w:hint is what would move them, and it is still deferred.
func nativeClassifyScriptSlots(text string) nativeScriptSlotUse {
	use := nativeScriptSlotUse{}
	for _, character := range text {
		switch {
		case !nativeRequiresScriptShaping(character):
			if character > 0x7f {
				use.hintBound = true
			}
		case character >= 0x1100 && character <= 0x11ff, // Hangul Jamo
			character >= 0x2e80 && character <= 0x2fdf, // CJK and Kangxi Radicals
			character >= 0x3000 && character <= 0x30ff, // CJK punctuation, Hiragana, Katakana
			character >= 0x3100 && character <= 0x318f, // Bopomofo, Hangul Compatibility Jamo
			character >= 0x31c0 && character <= 0x31ff, // CJK strokes, Katakana extensions
			character >= 0x3200 && character <= 0x4dbf, // Enclosed CJK, compatibility, extension A
			character >= 0x4e00 && character <= 0x9fff, // CJK Unified Ideographs
			character >= 0xac00 && character <= 0xd7af, // Hangul Syllables
			character >= 0xf900 && character <= 0xfaff, // CJK Compatibility Ideographs
			character >= 0xfe30 && character <= 0xfe4f, // CJK Compatibility Forms
			character >= 0xff00 && character <= 0xffef, // Halfwidth and Fullwidth Forms
			character >= 0x20000 && character <= 0x2fa1f: // Supplementary Ideographic Plane
			use.eastAsia = true
		case nativeComplexScriptSlotRune(character):
			use.complex = true
		default:
			use.unmodelled = true
		}
	}
	return use
}

// nativeComplexScriptSlotRune reports whether MS-OI29500 17.3.2.26 assigns a
// rune to the complex-script font slot (w:rFonts/@w:cs), narrowed to the
// Hebrew and Arabic blocks whose Unicode script this tier's pinned HarfBuzz
// provider qualifies.
//
// The narrowing is deliberate and is what bounds this slot's blast radius.
// Syriac, Thaana, NKo and the Indic, Thai and Khmer ranges are complex scripts
// too, but the provider refuses their script names, so routing them here would
// only move a refusal from one code to another while re-pointing their font
// selection with no reference export to check it against. They stay unmodelled
// and keep the refusal they have today.
//
// U+FEFD..U+FEFF are excluded: U+FEFF is the byte-order mark, script Common,
// and is not complex-script text.
func nativeComplexScriptSlotRune(character rune) bool {
	switch {
	case character >= 0x0590 && character <= 0x05ff, // Hebrew
		character >= 0x0600 && character <= 0x06ff, // Arabic
		character >= 0x0750 && character <= 0x077f, // Arabic Supplement
		character >= 0x08a0 && character <= 0x08ff, // Arabic Extended-A
		character >= 0xfb1d && character <= 0xfb4f, // Hebrew Presentation Forms
		character >= 0xfb50 && character <= 0xfdff, // Arabic Presentation Forms-A
		character >= 0xfe70 && character <= 0xfefc: // Arabic Presentation Forms-B
		return true
	default:
		return false
	}
}
