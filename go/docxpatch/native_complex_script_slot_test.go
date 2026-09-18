package docxpatch

import (
	"strings"
	"testing"
)

// The complex-script slot's rune table. It is narrowed on purpose to the
// Hebrew and Arabic blocks whose Unicode script this tier's pinned shaper
// qualifies; every other complex script keeps the refusal it has today.
func TestNativeComplexScriptSlotRuneTable(t *testing.T) {
	for _, tc := range []struct {
		name string
		char rune
		want bool
	}{
		{"below Hebrew", 0x058f, false},
		{"Hebrew lower bound", 0x0590, true},
		{"Hebrew alef", 0x05d0, true},
		{"Arabic alef", 0x0627, true},
		{"Arabic upper bound", 0x06ff, true},
		{"Syriac stays unmodelled", 0x0710, false},
		{"Arabic Supplement", 0x0750, true},
		{"Thaana stays unmodelled", 0x0780, false},
		{"Arabic Extended-A", 0x08a0, true},
		{"Devanagari stays unmodelled", 0x0905, false},
		{"Thai stays unmodelled", 0x0e01, false},
		{"Hebrew Presentation Forms", 0xfb1d, true},
		{"Arabic Presentation Forms-A", 0xfb50, true},
		{"CJK Compatibility Forms are East-Asian", 0xfe30, false},
		{"Arabic Presentation Forms-B", 0xfe70, true},
		{"Arabic Presentation Forms-B upper bound", 0xfefc, true},
		{"byte-order mark is not complex-script text", 0xfeff, false},
		{"Halfwidth and Fullwidth Forms are East-Asian", 0xff01, false},
		{"Basic Latin", 'A', false},
	} {
		if got := nativeComplexScriptSlotRune(tc.char); got != tc.want {
			t.Fatalf("%s (U+%04X): complex-script slot %v, want %v", tc.name, tc.char, got, tc.want)
		}
	}
}

// ECMA-376 17.3.2.7 states w:cs as the switch that applies the complex-script
// attributes to a run's contents, and 17.3.2.30 says the same of w:rtl. Either
// puts the whole run in the slot; without one, only the slot's own runes do.
func TestNativeComplexScriptSlotResolvesFaceSizeAndToggles(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` +
		`<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:sz w:val="22"/><w:szCs w:val="28"/>` +
		`<w:bCs/><w:iCs/><w:lang w:val="en-GB" w:bidi="ar-OM"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
	for _, tc := range []struct {
		name, rpr, text string
		slot, switched  bool
	}{
		{"plain latin stays in the ascii slot", "", "Hello", false, false},
		{"arabic runes reach the slot on their own", "", "مقدمة", true, false},
		{"w:rtl switches a Basic Latin run into the slot", `<w:rtl/>`, " ", true, true},
		{"w:cs switches a Basic Latin run into the slot", `<w:cs/>`, "Hello", true, true},
		{"w:cs w:val=false does not switch", `<w:cs w:val="false"/>`, "Hello", false, false},
		{"w:rtl w:val=false does not switch", `<w:rtl w:val="0"/>`, "Hello", false, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr>` + tc.rpr + `</w:rPr><w:t xml:space="preserve">` + tc.text + `</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			data := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			encoded, err := EncodeNativeDocumentV1(doc)
			if err != nil {
				t.Fatal(err)
			}
			// w:cs is a modelled switch now, not markup this tier cannot read.
			if strings.Contains(string(encoded), `"code":"UNMODELED_RUN_PROPERTY"`) {
				t.Fatalf("complex-script switch still unmodeled: %s", tc.rpr)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(resolved.Diagnostics) != 0 {
				t.Fatalf("resolved complex-script slot still refused: %#v", resolved.Diagnostics)
			}
			run := resolved.Runs[0].Properties
			if run.FontFamily == nil || *run.FontFamily != "Calibri" || run.FontSizeHalfPoint == nil || *run.FontSizeHalfPoint != 22 {
				t.Fatalf("ascii slot was displaced: %#v", run)
			}
			if (run.ComplexFontFamily != nil) != tc.slot || (run.ComplexFontSizeHalfPoint != nil) != tc.slot ||
				(run.ComplexBold != nil) != tc.slot || (run.ComplexItalic != nil) != tc.slot || (run.ComplexLanguage != nil) != tc.slot {
				t.Fatalf("complex-script slot was not carried by its own text: %#v", run)
			}
			if tc.slot && (*run.ComplexFontFamily != "Arial" || *run.ComplexFontSizeHalfPoint != 28 || !*run.ComplexBold || !*run.ComplexItalic || *run.ComplexLanguage != "ar-OM") {
				t.Fatalf("complex-script slot resolved wrong: %#v", run)
			}
			if (run.ComplexScriptSlot != nil) != tc.switched {
				t.Fatalf("run-level complex-script switch export wrong: %#v", run)
			}
			// The ascii slot's own bold/italic stay absent: w:bCs and w:iCs are
			// the complex slot's, and must not leak into Latin text.
			if run.Bold != nil || run.Italic != nil {
				t.Fatalf("complex-script toggles leaked into the ascii slot: %#v", run)
			}
		})
	}
}

// A run that reaches the complex-script slot and resolves no face in it has no
// font at all: it is refused rather than painted in the ascii face.
func TestNativeComplexScriptSlotWithNoFaceIsRefused(t *testing.T) {
	for _, tc := range []struct {
		name, slot, rpr string
		refused         bool
	}{
		{"no cs slot at all", "", `<w:rtl/>`, true},
		{"cs slot present", ` w:cs="Arial"`, `<w:rtl/>`, false},
		{"cstheme alone is not resolved", ` w:cstheme="minorBidi"`, `<w:rtl/>`, true},
		{"no switch asks for nothing", "", "", false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"` + tc.slot + `/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr>` + tc.rpr + `</w:rPr><w:t>Hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if hasResolutionDiagnostic(resolved, "SCRIPT_FONT_PRESERVED") != tc.refused {
				t.Fatalf("complex-script font refusal wrong: %#v", resolved.Diagnostics)
			}
			if run := resolved.Runs[0].Properties; tc.refused && run.ComplexFontFamily != nil {
				t.Fatalf("guessed a complex-script face: %#v", run)
			}
		})
	}
}

// The slot's face is a font the document genuinely needs, so it joins the
// inventory - under the weight and style w:bCs/w:iCs give it, not the run's.
func TestNativeComplexScriptSlotJoinsTheFontInventory(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Arial"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
	for _, tc := range []struct {
		name, rpr string
		want      []string
	}{
		{"inherited slot no text reaches asks for nothing", "", []string{"Calibri:400:normal"}},
		{"switched run asks for the slot's face", `<w:rtl/>`, []string{"Arial:400:normal", "Calibri:400:normal"}},
		{"w:bCs gives the slot its own weight", `<w:rtl/><w:bCs/>`, []string{"Arial:700:normal", "Calibri:400:normal"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := resolvedStylesTestParts(styles)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr>` + tc.rpr + `</w:rPr><w:t>Hello</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			inventory, err := ExtractNativeDOCXFontInventoryV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			got := map[string]bool{}
			for _, reference := range inventory.References {
				got[reference.Family+":"+itoaNativeTest(reference.Weight)+":"+reference.Style] = true
			}
			if len(got) != len(tc.want) {
				t.Fatalf("font inventory %#v, want %v", inventory.References, tc.want)
			}
			for _, want := range tc.want {
				if !got[want] {
					t.Fatalf("font inventory %#v, want %v", inventory.References, tc.want)
				}
			}
		})
	}
}

func itoaNativeTest(value int) string {
	if value == 700 {
		return "700"
	}
	return "400"
}
