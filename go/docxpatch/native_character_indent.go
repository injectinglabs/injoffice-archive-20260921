package docxpatch

// Word writes paragraph indentation either as an absolute twip measure
// (w:left, w:firstLine, ...) or as a character-unit measure in hundredths of a
// character (w:leftChars, w:firstLineChars, ...). A character measure has no
// length until a character width is known, and that width depends on the
// paragraph's East-Asian font metrics, which this layer does not have.
//
// Two character measures are nevertheless exactly convertible without any font
// metric, and only those two are resolved here:
//
//   - A zero character measure is zero twips for every font. Word writes
//     w:leftChars="0" as the ordinary way to cancel a style's character indent,
//     so refusing it withholds layout from documents that carry no character
//     indent at all.
//   - A non-zero character measure whose absolute companion attribute is
//     present on the same w:ind element carries the producer's own conversion
//     of that measure. Word emits both attributes together and keeps the twip
//     value in step with the character value on every save, so the companion is
//     an attested length rather than a width this layer guessed.
//
// Any other non-zero character measure still refuses: the width it needs is not
// in the source and is not invented.
var nativeCharacterIndentPairs = []struct{ characters, absolute string }{
	{"leftChars", "left"},
	{"rightChars", "right"},
	{"startChars", "start"},
	{"endChars", "end"},
	{"firstLineChars", "firstLine"},
	{"hangingChars", "hanging"},
}

// nativeCharacterIndentResolution reports how one character-unit attribute of a
// w:ind element resolves. When present is false the attribute is absent. When
// present is true and founded is false the measure needs a character width the
// source does not attest, and the caller keeps refusing. When founded is true,
// zero reports whether the measure resolves to an explicit zero length (in
// which case it supersedes any absolute companion) or adopts the companion.
func nativeCharacterIndentResolution(node *nativeXMLNode, wordNS, characters, absolute string) (present, founded, zero bool) {
	if _, ok := nativeAttr(node, wordNS, characters); !ok {
		return false, false, false
	}
	value, valid := nativeSignedInt64Attr(node, wordNS, characters)
	if !valid {
		return true, false, false
	}
	if value == 0 {
		return true, true, true
	}
	if _, ok := nativeAttr(node, wordNS, absolute); !ok {
		return true, false, false
	}
	// The companion must itself be a value this layer resolves; an invalid one
	// is no conversion at all.
	if absolute == "firstLine" || absolute == "hanging" {
		if _, ok := nativeNonnegativeInt64Attr(node, wordNS, absolute); !ok {
			return true, false, false
		}
		return true, true, false
	}
	if _, ok := nativeSignedInt64Attr(node, wordNS, absolute); !ok {
		return true, false, false
	}
	return true, true, false
}
