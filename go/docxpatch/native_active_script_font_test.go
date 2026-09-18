package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeScriptPropertiesResolveAtActualRunScript(t *testing.T) {
	metadata := `<w:rFonts w:ascii="Exact Latin" w:hAnsi="Exact Latin" w:eastAsia="Unneeded CJK" w:cs="Unneeded Arabic" w:cstheme="minorBidi" w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="40"/><w:bCs/><w:iCs/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>`
	for _, strict := range []bool{false, true} {
		// The five deferred codes each speak for one slot. Complex-script text
		// now reaches a modelled slot, so only its font selection can still be
		// preserved: this metadata states both w:cs and w:cstheme, which names
		// no single face, so the slot is refused rather than guessed. A rune
		// this tier assigns to no slot at all still flushes every one of them.
		for _, test := range []struct {
			name, text, direct string
			scriptRequired     bool
			complexOnly        bool
		}{
			{"Basic Latin", "Plain text 123!", "", false, false},
			{"Arabic", "مرحبا", "", true, true},
			// The East-Asian slot this metadata states is resolved, so CJK text
			// no longer defers: it selects that face and the complex-script
			// properties beside it state formatting for text that is not here.
			{"CJK", "你好", "", false, false},
			{"mixed scripts", "Latin العربية", "", true, true},
			// MS-OI29500 17.3.2.26 resolves Latin-1 Supplement, Latin Extended and
			// General Punctuation through the ascii/hAnsi slots, not cs/eastAsia,
			// so these need no script font selection.
			{"Latin-1 Supplement", "café", "", false, false},
			{"Latin Extended-A", "Ostrov Krk — Hrvatska š", "", false, false},
			{"General Punctuation curly quotes", "he said \u201chello\u201d", "", false, false},
			{"Greek is script-bearing", "\u03b1\u03b2\u03b3", "", true, false},
			{"Cyrillic is script-bearing", "\u043f\u0440\u0438\u0432\u0435\u0442", "", true, false},
			{"Hebrew is script-bearing", "\u05e9\u05dc\u05d5\u05dd", "", true, true},
			{"explicit RTL", "ASCII", `<w:rtl/>`, true, true},
		} {
			t.Run(test.name+map[bool]string{false: " transitional", true: " strict"}[strict], func(t *testing.T) {
				styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` + metadata + `</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
				parts := resolvedStylesTestParts(styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr>` + test.direct + `</w:rPr><w:t>` + test.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
				if strict {
					for key, value := range parts {
						parts[key] = strings.ReplaceAll(strings.ReplaceAll(value, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := append([]byte(nil), data...)
				resolved, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				for _, code := range []string{"SCRIPT_FONT_PRESERVED", "FONT_HINT_PRESERVED", "COMPLEX_SCRIPT_SIZE_PRESERVED", "COMPLEX_SCRIPT_TOGGLE_PRESERVED", "SCRIPT_LANGUAGE_PRESERVED"} {
					want := test.scriptRequired && (!test.complexOnly || code == "SCRIPT_FONT_PRESERVED")
					if hasResolutionDiagnostic(resolved, code) != want {
						t.Fatalf("wrong active script decision for %s: %#v", code, resolved.Diagnostics)
					}
				}
				if !test.scriptRequired {
					run := resolved.Runs[0].Properties
					if run.FontFamily == nil || *run.FontFamily != "Exact Latin" || run.FontSizeHalfPoint == nil || *run.FontSizeHalfPoint != 24 || run.Bold != nil || run.Italic != nil || run.Language == nil || *run.Language != "en-US" {
						t.Fatalf("inactive script fields altered Latin formatting %#v", run)
					}
					// The East-Asian slot is carried only by text that reaches it.
					eastAsian := strings.ContainsRune(test.text, '你')
					if (run.EastAsiaFontFamily != nil) != eastAsian || (run.EastAsiaLanguage != nil) != eastAsian {
						t.Fatalf("east-asian slot was not carried by its own text %#v", run)
					}
					if eastAsian && (*run.EastAsiaFontFamily != "Unneeded CJK" || *run.EastAsiaLanguage != "zh-CN") {
						t.Fatalf("east-asian slot resolved to the wrong face/language %#v", run)
					}
				}
				if !bytes.Equal(before, data) {
					t.Fatal("render resolution mutated source")
				}
			})
		}
	}
}

func TestNativeScriptPropertiesKeepMalformedAndForcedScriptRefusals(t *testing.T) {
	for _, markup := range []string{
		`<w:rFonts w:ascii="Latin" w:csTheme="minorBidi"/>`,
		`<w:rFonts w:ascii="Latin" w:cstheme="unknown"/>`,
		`<w:rFonts w:ascii="Latin" w:hint="unknown"/>`,
		`<w:szCs w:val="bad"/>`, `<w:szCs w:val="24"><w:b/></w:szCs>`,
		`<w:bCs w:val="bad"/>`, `<w:iCs w:unknown="1"/>`,
		`<w:rtl w:val="false" w:unknown="1"/>`,
		`<w:lang w:val="en-US" w:unknown="ar-SA"/>`,
		`<w:lang w:val="en-US" w:bidi="en_US"/>`,
		`<w:lang w:val="en-US" w:eastAsia="   "/>`,
		`<w:cs/>`,
	} {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` + markup + `</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
		resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if len(resolved.Diagnostics) == 0 {
			t.Fatalf("unsafe script metadata accepted: %s", markup)
		}
	}
}

// ECMA-376 17.3.2.30 (w:rtl) and 17.3.2.26 (w:cs) select the complex-script
// slot per run, the paragraph mark included. Paragraph-level w:bidi (17.3.1.6)
// only orders the line, so on its own it must not defer the mark's script
// properties; an explicit mark-level w:rtl still must.
func TestNativeEmptyParagraphScriptSelectionFollowsMarkRunProperties(t *testing.T) {
	for _, test := range []struct {
		name     string
		property string
		deferred bool
	}{
		{"plain", "", false},
		{"paragraph bidi only", "<w:pPr><w:bidi/></w:pPr>", false},
		// The mark's own w:rtl still selects the complex-script slot, but the
		// slot now resolves: w:docDefaults names a w:cs face and a w:szCs, so
		// the mark is measured in that face at that size instead of refusing.
		{"mark rtl", "<w:pPr><w:rPr><w:rtl/></w:rPr></w:pPr>", false},
		{"paragraph bidi with mark rtl", "<w:pPr><w:bidi/><w:rPr><w:rtl/></w:rPr></w:pPr>", false},
		{"mark cs switch", "<w:pPr><w:rPr><w:cs/></w:rPr></w:pPr>", false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Latin" w:hAnsi="Latin" w:cs="Arabic"/><w:sz w:val="24"/><w:szCs w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p>` + test.property + `</w:p><w:sectPr/></w:body></w:document>`
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if hasResolutionDiagnostic(resolved, "SCRIPT_FONT_PRESERVED") != test.deferred || hasResolutionDiagnostic(resolved, "COMPLEX_SCRIPT_SIZE_PRESERVED") != test.deferred {
				t.Fatalf("wrong empty mark script decision %#v", resolved.Diagnostics)
			}
			mark := resolved.Paragraphs[0].ParagraphMarkProperties
			if !test.deferred && (mark.FontFamily == nil || *mark.FontFamily != "Latin" || mark.FontSizeHalfPoint == nil || *mark.FontSizeHalfPoint != 24) {
				t.Fatalf("literal-CR Latin mark changed %#v", mark)
			}
			// A mark that states the switch carries the slot it switches to; a
			// mark that does not carries nothing of it, so a paragraph that only
			// inherits a w:cs face asks for no complex-script font.
			switched := strings.Contains(test.property, "w:rtl") || strings.Contains(test.property, "w:cs")
			if (mark.ComplexScriptSlot != nil) != switched || (mark.ComplexFontFamily != nil) != switched || (mark.ComplexFontSizeHalfPoint != nil) != switched {
				t.Fatalf("complex-script slot carried without its switch %#v", mark)
			}
			if switched && (*mark.ComplexFontFamily != "Arabic" || *mark.ComplexFontSizeHalfPoint != 40) {
				t.Fatalf("complex-script slot resolved to the wrong face/size %#v", mark)
			}
		})
	}
}

// The slot decision is by Unicode range, not by an ASCII cutoff: a rune above
// U+007F that Word resolves through ascii/hAnsi must not defer script
// properties, while East-Asian and complex-script ranges still must.
func TestNativeRequiresScriptShapingByUnicodeRange(t *testing.T) {
	for _, tc := range []struct {
		name string
		char rune
		want bool
	}{
		{"ASCII upper bound", 0x007f, false},
		{"C1 control stays deferred", 0x0080, true},
		{"Latin-1 Supplement lower bound", 0x00a0, false},
		{"u with diaeresis", 0x00fc, false},
		{"Latin Extended-A s with caron", 0x0161, false},
		{"Latin Extended-B upper bound", 0x024f, false},
		{"IPA Extensions stays deferred", 0x0250, true},
		{"Greek stays deferred", 0x03b1, true},
		{"Cyrillic stays deferred", 0x0440, true},
		{"Hebrew stays deferred", 0x05d0, true},
		{"Arabic stays deferred", 0x0627, true},
		{"General Punctuation lower bound", 0x2000, false},
		{"en dash", 0x2013, false},
		{"left double quotation mark", 0x201c, false},
		{"General Punctuation upper bound", 0x206f, false},
		{"Superscripts stay deferred", 0x2070, true},
		{"CJK stays deferred", 0x4f60, true},
		{"just below the Private Use Area stays deferred", 0xdfff, true},
		{"Private Use Area lower bound", 0xe000, false},
		{"Symbol font bullet", 0xf0b7, false},
		{"Private Use Area upper bound", 0xf8ff, false},
		{"CJK Compatibility Ideographs above the Private Use Area stay deferred", 0xf900, true},
		{"Enclosed Alphanumerics resolve through High ANSI", 0x2460, false},
		{"Box Drawing above Enclosed Alphanumerics stays deferred", 0x2500, true},
		{"supplementary Private Use Area stays deferred", 0xf0000, true},
	} {
		if got := nativeRequiresScriptShaping(tc.char); got != tc.want {
			t.Fatalf("%s (U+%04X): requires script shaping %v, want %v", tc.name, tc.char, got, tc.want)
		}
	}
}
