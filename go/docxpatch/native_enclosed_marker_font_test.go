package docxpatch

import (
	"strings"
	"testing"
)

// MS-OI29500 17.3.2.26 gives the Enclosed Alphanumerics block the High ANSI
// slot, so no rune in it needs script shaping. The neighbouring blocks are the
// negative control: they are outside the row this change reads and keep
// deferring exactly as before.
func TestNativeEnclosedAlphanumericsResolveThroughHighANSI(t *testing.T) {
	for _, tc := range []struct {
		name string
		char rune
		want bool
	}{
		{"below the block stays script-bearing", 0x245f, true},
		{"circled digit one", 0x2460, false},
		{"circled digit twenty, the decimalEnclosedCircle upper bound", 0x2473, false},
		{"block upper bound", 0x24ff, false},
		{"above the block stays script-bearing", 0x2500, true},
		{"CJK radicals stay script-bearing", 0x2e80, true},
	} {
		if got := nativeRequiresScriptShaping(tc.char); got != tc.want {
			t.Fatalf("%s (U+%04X): requires script shaping %v, want %v", tc.name, tc.char, got, tc.want)
		}
	}
}

func enclosedMarkerTestParts(numFmt, hint string) map[string]string {
	level := `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="` + numFmt + `"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/>` +
		`<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`
	parts := resolvedNumberingTestParts(`<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="0">` + level +
		`</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`,
		`<Relationship Id="styles" Type="`+relBaseTransitional+`styles" Target="styles.xml"/></Relationships>`, 1)
	// The corpus shape: a Latin ascii/hAnsi face beside East-Asian and
	// complex-script slots the document merely inherits and the marker never asks
	// for, exactly as numbering-circle.docx states them.
	parts["word/styles.xml"] = `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` +
		`<w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:eastAsia="Noto Serif CJK SC" w:cs="Lohit Devanagari"` + hint + `/>` +
		`<w:sz w:val="24"/><w:szCs w:val="24"/><w:lang w:val="de-DE" w:eastAsia="zh-CN" w:bidi="hi-IN"/>` +
		`</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
		`<w:r><w:t>first</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	return parts
}

// The whole point of the slot row: before it, an enclosed-digit marker was a
// rune in no slot at all, so every deferred w:eastAsia/w:cs property flushed and
// the paragraph was refused before shaping. After it the marker resolves the
// High ANSI face, and what is left is a repertoire fact about that face.
func TestNativeEnclosedMarkerResolvesTheHighANSIFaceAndRecordsItsRepertoire(t *testing.T) {
	for _, tc := range []struct {
		name, numFmt, hint                string
		scriptRefused, repertoireRecorded bool
	}{
		{name: "enclosed circle marker", numFmt: "decimalEnclosedCircle", repertoireRecorded: true},
		{name: "plain decimal marker is untouched", numFmt: "decimal"},
		{name: "an eastAsia hint keeps the previous refusal", numFmt: "decimalEnclosedCircle", hint: ` w:hint="eastAsia"`, scriptRefused: true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			data := buildNativeDOCX(t, nativeEntries(enclosedMarkerTestParts(tc.numFmt, tc.hint)))
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if hasResolutionDiagnostic(resolved, "SCRIPT_FONT_PRESERVED") != tc.scriptRefused {
				t.Fatalf("script font refusal wrong: %#v", resolved.Diagnostics)
			}
			if hasResolutionDiagnostic(resolved, "ENCLOSED_NUMBER_MARKER_FONT_PRESERVED") != tc.repertoireRecorded {
				t.Fatalf("repertoire diagnostic wrong: %#v", resolved.Diagnostics)
			}
			marker := resolved.Paragraphs[0].Numbering
			if marker == nil || marker.Marker.FontFamily == nil || *marker.Marker.FontFamily != "Liberation Serif" {
				t.Fatalf("marker did not resolve the High ANSI face: %#v", marker)
			}
			// The slot decision never re-points the marker: strict resolution
			// keeps the authored family whatever the diagnostic says.
			if marker.Marker.EastAsiaFontFamily != nil || marker.Marker.ComplexFontFamily != nil {
				t.Fatalf("marker took a script slot it never reaches: %#v", marker.Marker)
			}
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(eligibility.EnclosedMarkerFonts) != map[bool]int{true: 1, false: 0}[tc.repertoireRecorded] {
				t.Fatalf("enclosed-marker evidence wrong: %#v", eligibility.EnclosedMarkerFonts)
			}
			for _, fact := range eligibility.EnclosedMarkerFonts {
				if fact.ScopeKind != "numbering-marker" || fact.ScopeID != resolved.Paragraphs[0].ParagraphID ||
					fact.SourceFamily != "Liberation Serif" || fact.PartName != "word/document.xml" || !strings.HasPrefix(fact.Path, "/w:document") {
					t.Fatalf("enclosed-marker fact does not name its own scope and resolved face: %#v", fact)
				}
			}
		})
	}
}
