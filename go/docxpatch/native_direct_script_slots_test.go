package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeDirectScriptSlotsRemainContextQualified(t *testing.T) {
	base := `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="CJK Face"/><w:sz w:val="24"/><w:szCs w:val="40"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="hi-IN"/>`
	for _, strict := range []bool{false, true} {
		for _, tc := range []struct {
			name, text, extra, props   string
			safe, blocked, sizeBlocked bool
		}{
			{"latin", "Hello", "", base, true, false, false},
			// The stated w:eastAsia face resolves the slot this text reaches,
			// so neither the font nor the complex-script size defers any more.
			{"cjk", "A漢", "", base, true, false, false},
			// ECMA-376 17.3.2.30: run-level w:rtl selects the complex-script
			// slot; paragraph-level w:bidi (17.3.1.6) only orders the line.
			{"paragraph-bidi", "Hello", `<w:bidi/>`, base, true, false, false},
			// w:rtl switches the whole run into the complex-script slot, and this
			// metadata names no w:cs face, so the font slot is refused. Its size
			// no longer is: w:szCs is the slot's own size and is now applied.
			{"run-rtl", "Hello", "", strings.Replace(base, `<w:lang`, `<w:rtl/><w:lang`, 1), true, true, false},
			{"paragraph-bidi-and-run-rtl", "Hello", `<w:bidi/>`, strings.Replace(base, `<w:lang`, `<w:rtl/><w:lang`, 1), true, true, false},
			// The same run with a w:cs face resolves the slot outright.
			{"run-rtl-with-cs", "Hello", "", strings.Replace(strings.Replace(base, `<w:lang`, `<w:rtl/><w:lang`, 1), `w:eastAsia="CJK Face"`, `w:eastAsia="CJK Face" w:cs="Complex Face"`, 1), true, false, false},
			// w:cs is the same switch stated on its own (ECMA-376 17.3.2.7).
			{"run-cs-switch", "Hello", "", strings.Replace(strings.Replace(base, `<w:lang`, `<w:cs/><w:lang`, 1), `w:eastAsia="CJK Face"`, `w:eastAsia="CJK Face" w:cs="Complex Face"`, 1), true, false, false},
			// Arabic reaches the slot by its own runes, with no switch at all.
			{"arabic-runes", "\u0645\u0631\u062d\u0628\u0627", "", strings.Replace(base, `w:eastAsia="CJK Face"`, `w:eastAsia="CJK Face" w:cs="Complex Face"`, 1), true, false, false},
			{"bad-size", "Hello", "", strings.Replace(base, `w:val="40"`, `w:val="0"`, 1), false, false, false},
			{"duplicate-size", "Hello", "", base + `<w:szCs w:val="40"/>`, false, false, false},
			{"size-child", "Hello", "", strings.Replace(base, `<w:szCs w:val="40"/>`, `<w:szCs w:val="40"><w:b/></w:szCs>`, 1), false, false, false},
			{"bad-language", "Hello", "", strings.Replace(base, `zh-CN`, `bad value`, 1), false, false, false},
			{"unknown-language", "Hello", "", strings.Replace(base, `w:eastAsia="zh-CN"`, `w:eastAsia="zh-CN" w:unknown="x"`, 1), false, false, false},
		} {
			t.Run(tc.name+map[bool]string{false: "-transitional", true: "-strict"}[strict], func(t *testing.T) {
				parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr>` + tc.extra + `<w:rPr>` + tc.props + `</w:rPr></w:pPr><w:r><w:rPr>` + tc.props + `</w:rPr><w:t>` + tc.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
				if strict {
					for k, v := range parts {
						parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := bytes.Clone(data)
				doc, err := ExtractNativeDocumentV1(data)
				if err != nil {
					t.Fatal(err)
				}
				encoded, err := EncodeNativeDocumentV1(doc)
				if err != nil {
					t.Fatal(err)
				}
				if strings.Contains(string(encoded), `"code":"UNMODELED_PARAGRAPH_MARK_PROPERTIES"`) == tc.safe {
					t.Fatalf("mark source guard: %s", encoded)
				}
				if tc.safe && strings.Contains(string(encoded), `"code":"PARTIAL_RUN_PROPERTIES"`) {
					t.Fatalf("inactive source slot still refused: %s", encoded)
				}
				if strings.Contains(string(encoded), `"code":"PARTIAL_PARAGRAPH_PROPERTIES"`) == tc.safe {
					t.Fatalf("paragraph mark confused preservation with invalidity: %s", encoded)
				}
				if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" || !bytes.Equal(before, data) {
					t.Fatal("source authority changed")
				}
				layout, err := ResolveNativeDocumentLayoutV1(data)
				if err != nil {
					t.Fatal(err)
				}
				if tc.safe && hasResolutionDiagnostic(layout, "SCRIPT_FONT_PRESERVED") != tc.blocked {
					t.Fatalf("script context lost: %#v", layout.Diagnostics)
				}
				if tc.safe && hasResolutionDiagnostic(layout, "COMPLEX_SCRIPT_SIZE_PRESERVED") != tc.sizeBlocked {
					t.Fatalf("size context lost: %#v", layout.Diagnostics)
				}
				if strings.HasPrefix(tc.name, "run-rtl-with-cs") || tc.name == "run-cs-switch" || tc.name == "arabic-runes" {
					run := layout.Runs[0].Properties
					if run.ComplexFontFamily == nil || *run.ComplexFontFamily != "Complex Face" || run.ComplexFontSizeHalfPoint == nil || *run.ComplexFontSizeHalfPoint != 40 || run.ComplexLanguage == nil || *run.ComplexLanguage != "hi-IN" {
						t.Fatalf("complex-script slot did not resolve beside the latin slot: %#v", run)
					}
					if run.FontFamily == nil || *run.FontFamily != "Arial" || run.FontSizeHalfPoint == nil || *run.FontSizeHalfPoint != 24 {
						t.Fatalf("complex-script slot displaced the ascii slot: %#v", run)
					}
					// The switch is exported only where a switch was stated.
					if (run.ComplexScriptSlot != nil) != (tc.name != "arabic-runes") {
						t.Fatalf("complex-script switch export wrong: %#v", run)
					}
				}
				if tc.name == "cjk" {
					run := layout.Runs[0].Properties
					if run.EastAsiaFontFamily == nil || *run.EastAsiaFontFamily != "CJK Face" || run.FontFamily == nil || *run.FontFamily != "Arial" {
						t.Fatalf("east-asian slot did not resolve beside the latin slot: %#v", run)
					}
				}
			})
		}
	}
}
