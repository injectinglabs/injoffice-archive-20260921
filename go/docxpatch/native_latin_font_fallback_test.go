package docxpatch

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"strings"
	"testing"
)

func latinFallbackTestData(t *testing.T, rFonts string) []byte {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` + rFonts + `<w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"/></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>text</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	return buildNativeDOCX(t, nativeEntries(parts))
}

func TestLatinFontFallbackEvidenceLeavesStrictResolutionUntouched(t *testing.T) {
	for _, tc := range []struct {
		name, rFonts, wantFace string
		wantFacts              int
	}{
		{name: "empty complex-script slot", rFonts: `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs=""/>`, wantFace: "Calibri", wantFacts: 2},
		{name: "empty east-asian slot", rFonts: `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia=""/>`, wantFace: "Arial", wantFacts: 2},
		{name: "unbounded complex-script slot", rFonts: `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="` + strings.Repeat("a", 300) + `"/>`},
		{name: "invalid hint", rFonts: `<w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="latin"/>`},
		{name: "empty slot but empty latin faces", rFonts: `<w:rFonts w:ascii="" w:hAnsi="" w:cs=""/>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data := latinFallbackTestData(t, tc.rFonts)
			before := bytes.Clone(data)
			layout, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			// Strict resolution: no face, original diagnostic text, run and mark alike.
			if layout.Runs[0].Properties.FontFamily != nil || layout.Paragraphs[0].ParagraphMarkProperties.FontFamily != nil {
				t.Fatalf("strict face was resolved: %#v", layout.Runs[0].Properties)
			}
			found := false
			for _, diagnostic := range layout.Diagnostics {
				if diagnostic.Code == "UNMODELED_FONT_SELECTION" && diagnostic.Message == "Invalid script font slot or hint is preserved and not resolved" {
					found = true
				}
			}
			if !found {
				t.Fatalf("strict UNMODELED_FONT_SELECTION missing or reworded: %#v", layout.Diagnostics)
			}
			facts, err := nativeLatinFontFallbacks(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(facts) != tc.wantFacts {
				t.Fatalf("facts=%d want %d: %#v", len(facts), tc.wantFacts, facts)
			}
			kinds := map[string]bool{}
			for _, fact := range facts {
				kinds[fact.ScopeKind] = true
				if fact.FontFamily != tc.wantFace || fact.PartName != "word/document.xml" || !strings.HasPrefix(fact.Path, "/w:document[1]/w:body[1]/w:p[1]") || fact.ScopeID == "" || !strings.HasPrefix(fact.PackageSHA256, "sha256:") {
					t.Fatalf("unbound fact %#v", fact)
				}
			}
			if tc.wantFacts > 0 && (!kinds["run"] || !kinds["paragraph-mark"]) {
				t.Fatalf("expected run and paragraph-mark facts: %#v", facts)
			}
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(eligibility.LatinFontFallbacks) != tc.wantFacts {
				t.Fatalf("eligibility facts=%d want %d", len(eligibility.LatinFontFallbacks), tc.wantFacts)
			}
			if !bytes.Equal(data, before) {
				t.Fatal("source mutated")
			}
		})
	}
}

// The strict resolved layout for an empty w:cs slot is pinned byte-for-byte to
// the pre-evidence behavior (main a81f790e/c7b7bece): the evidence pass must
// never leak into strict output.
func TestLatinFontFallbackStrictLayoutGolden(t *testing.T) {
	data := latinFallbackTestData(t, `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs=""/>`)
	layout, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := json.Marshal(layout)
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(encoded)
	got := hex.EncodeToString(sum[:])
	const want = "a6d150b0d5f138b9f61d77f6efadba7399f993ae1309ee58461260d5547d5911"
	if got != want {
		t.Fatalf("strict resolved layout changed: %s (want %s)", got, want)
	}
}
