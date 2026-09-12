package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestNativeScriptPropertiesResolveAtActualRunScript(t *testing.T) {
	metadata := `<w:rFonts w:ascii="Exact Latin" w:hAnsi="Exact Latin" w:eastAsia="Unneeded CJK" w:cs="Unneeded Arabic" w:cstheme="minorBidi" w:hint="eastAsia"/><w:sz w:val="24"/><w:szCs w:val="40"/><w:bCs/><w:iCs/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>`
	for _, strict := range []bool{false, true} {
		for _, test := range []struct {
			name, text, direct string
			scriptRequired     bool
		}{
			{"Basic Latin", "Plain text 123!", "", false},
			{"Arabic", "مرحبا", "", true},
			{"CJK", "你好", "", true},
			{"mixed scripts", "Latin العربية", "", true},
			{"Latin outside bounded slice", "café", "", true},
			{"explicit RTL", "ASCII", `<w:rtl/>`, true},
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
					if hasResolutionDiagnostic(resolved, code) != test.scriptRequired {
						t.Fatalf("wrong active script decision for %s: %#v", code, resolved.Diagnostics)
					}
				}
				if !test.scriptRequired {
					run := resolved.Runs[0].Properties
					if run.FontFamily == nil || *run.FontFamily != "Exact Latin" || run.FontSizeHalfPoint == nil || *run.FontSizeHalfPoint != 24 || run.Bold != nil || run.Italic != nil || run.Language == nil || *run.Language != "en-US" {
						t.Fatalf("inactive script fields altered Latin formatting %#v", run)
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

func TestNativeEmptyParagraphScriptSelectionRetainsBidiUncertainty(t *testing.T) {
	for _, bidi := range []bool{false, true} {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Latin" w:hAnsi="Latin" w:cs="Arabic"/><w:sz w:val="24"/><w:szCs w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`)
		property := ""
		if bidi {
			property = "<w:pPr><w:bidi/></w:pPr>"
		}
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p>` + property + `</w:p><w:sectPr/></w:body></w:document>`
		resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if hasResolutionDiagnostic(resolved, "SCRIPT_FONT_PRESERVED") != bidi || hasResolutionDiagnostic(resolved, "COMPLEX_SCRIPT_SIZE_PRESERVED") != bidi {
			t.Fatalf("wrong empty mark script decision bidi=%v %#v", bidi, resolved.Diagnostics)
		}
		mark := resolved.Paragraphs[0].ParagraphMarkProperties
		if !bidi && (mark.FontFamily == nil || *mark.FontFamily != "Latin" || mark.FontSizeHalfPoint == nil || *mark.FontSizeHalfPoint != 24) {
			t.Fatalf("literal-CR Latin mark changed %#v", mark)
		}
	}
}
