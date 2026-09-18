package docxpatch

import (
	"strings"
	"testing"
)

// The rFonts range table gives the Basic Multilingual Plane Private Use Area
// the High ANSI slot, with w:hint="eastAsia" as its one escape. Word's export
// of tdf118812_tableStyles-comprehensive.docx paints its U+F0B7 bullets in
// SymbolMT - the face the level states as w:ascii/w:hAnsi - and carries no
// East-Asian face at all, so a symbol code point asks for no script font
// selection and must not flush the run's deferred script slots.
func TestNativePrivateUseAreaResolvesThroughHighANSI(t *testing.T) {
	slots := `<w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:eastAsia="CJK Face" w:cs="Arabic Face" w:hint="HINT"/><w:sz w:val="24"/><w:szCs w:val="40"/><w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/>`
	for _, tc := range []struct {
		name, text, hint string
		blocked          bool
	}{
		// The exact shape the corpus states: a private-use bullet under an
		// explicit hint that is not eastAsia.
		{"symbol bullet with default hint", "", "default", false},
		{"symbol bullet with complex-script hint", "", "cs", false},
		{"private use lower bound", "", "default", false},
		{"private use upper bound", "", "default", false},
		{"private use beside basic latin", "ab", "default", false},
		{"private use beside an en dash", "–", "default", false},
		// The escape. A package that states the hint keeps the refusal it had
		// before this range was read as High ANSI.
		{"symbol bullet with east-asian hint", "", "eastAsia", true},
		{"private use beside basic latin with east-asian hint", "ab", "eastAsia", true},
		// Neighbours of the range, and the supplementary planes the table does
		// not cover, keep deferring.
		{"just below the range", "퟿", "default", true},
		{"cjk compatibility ideograph above the range", "豈", "default", true},
		{"enclosed alphanumeric", "①", "default", true},
		{"supplementary private use", "\U000f0000", "default", true},
		{"arabic", "ا", "default", true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			props := strings.Replace(slots, "HINT", tc.hint, 1)
			parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr>` + props + `</w:rPr><w:t>` + tc.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			blocked := false
			for _, code := range []string{"SCRIPT_FONT_PRESERVED", "FONT_HINT_PRESERVED", "COMPLEX_SCRIPT_SIZE_PRESERVED", "SCRIPT_LANGUAGE_PRESERVED"} {
				if hasResolutionDiagnostic(layout, code) {
					blocked = true
				}
			}
			if blocked != tc.blocked {
				t.Fatalf("script slot decision blocked=%v, want %v: %#v", blocked, tc.blocked, layout.Diagnostics)
			}
			if tc.blocked {
				return
			}
			run := layout.Runs[0].Properties
			if run.FontFamily == nil || *run.FontFamily != "Symbol" {
				t.Fatalf("private-use text did not resolve through the ascii/hAnsi slot: %#v", run)
			}
			if run.EastAsiaFontFamily != nil {
				t.Fatalf("private-use text asked for an East-Asian face: %#v", run)
			}
		})
	}
}

// The escape is read from the run's own cascade, so a hint stated by an outer
// layer still moves the range and an inner layer still overrules it.
func TestNativePrivateUseAreaHintCascades(t *testing.T) {
	for _, tc := range []struct {
		name, defaults, direct string
		blocked                bool
	}{
		{"hint from docDefaults", `w:hint="eastAsia"`, ``, true},
		{"direct hint overrules docDefaults", `w:hint="eastAsia"`, `<w:rFonts w:hint="default"/>`, false},
		{"docDefaults default hint", `w:hint="default"`, ``, false},
		{"direct hint overrules a default", `w:hint="default"`, `<w:rFonts w:hint="eastAsia"/>`, true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" ` + tc.defaults + `/><w:sz w:val="24"/><w:szCs w:val="40"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr>` + tc.direct + `</w:rPr><w:t>` + "" + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if hasResolutionDiagnostic(layout, "COMPLEX_SCRIPT_SIZE_PRESERVED") != tc.blocked {
				t.Fatalf("cascaded hint decision blocked=%v, want %v: %#v", !tc.blocked, tc.blocked, layout.Diagnostics)
			}
		})
	}
}
