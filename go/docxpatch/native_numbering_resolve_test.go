package docxpatch

import (
	"bytes"
	"crypto/sha256"
	"encoding/base64"
	"fmt"
	"os"
	"strings"
	"testing"
)

// wordMacListFixture returns the untouched bytes of python-docx's Apache-2.0
// features/steps/test_files/num-having-numbering-part.docx at commit
// e45454602b53e8e572b179ccf1c91093ec9f4ed7, saved by Microsoft Macintosh
// Word 14.0. The textual envelope keeps the repository patch reviewable.
func wordMacListFixture(t *testing.T) []byte {
	t.Helper()
	encoded, err := os.ReadFile("testdata/word-mac-list-numbering.docx.base64")
	if err != nil {
		t.Fatal(err)
	}
	data, err := base64.StdEncoding.DecodeString(strings.Join(strings.Fields(string(encoded)), ""))
	if err != nil {
		t.Fatal(err)
	}
	if digest := fmt.Sprintf("%x", sha256.Sum256(data)); digest != "66b45ace7afa65acab6036bdc63d9416554576df633c13a6774bf41d4a60a47e" {
		t.Fatalf("Word fixture digest = %s", digest)
	}
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(pkg.files["docProps/app.xml"], []byte("Microsoft Macintosh Word")) {
		t.Fatal("Word producer metadata is missing from genuine fixture")
	}
	return data
}

func strictNumberingFixture(t *testing.T) []byte {
	t.Helper()
	encoded, err := os.ReadFile("testdata/native-numbering-strict-v1.docx.base64")
	if err != nil {
		t.Fatal(err)
	}
	data, err := base64.StdEncoding.DecodeString(strings.Join(strings.Fields(string(encoded)), ""))
	if err != nil {
		t.Fatal(err)
	}
	if digest := fmt.Sprintf("%x", sha256.Sum256(data)); digest != "9865a76e1111e0ab529c84a9bbf0b36b5f0b71fca612ef36f012825f8753b5f5" {
		t.Fatalf("Strict numbering fixture digest = %s", digest)
	}
	return data
}

func TestNativeNumberingGenuineMicrosoftWordFixture(t *testing.T) {
	resolved, err := ResolveNativeDocumentLayoutV1(wordMacListFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Paragraphs) != 1 {
		t.Fatalf("Word fixture paragraph count = %d", len(resolved.Paragraphs))
	}
	marker := resolved.Paragraphs[0].Numbering
	if marker == nil || marker.NumID != "6" || marker.AbstractNumID != "8" || marker.LevelStyleID == nil || *marker.LevelStyleID != "ListNumber" || marker.ResolvedText != "1." || marker.Alignment != "left" || marker.NumberingTabTwips == nil || *marker.NumberingTabTwips != 360 || marker.LabelStartTwips != 0 || marker.TextStartTwips != 360 {
		t.Fatalf("Word fixture marker projection is not exact: %#v", marker)
	}
	if hasResolutionDiagnostic(resolved, "UNSUPPORTED_MULTI_LEVEL_BEHAVIOR") {
		t.Fatalf("ordinary Word numbering emitted a stale blocking diagnostic: %#v", resolved.Diagnostics)
	}
	if resolved.NumberingSource == nil || resolved.NumberingSource.ModelSHA256 != nativeResolvedNumberingModelSHA256(resolved) {
		t.Fatalf("Word fixture numbering provenance is incomplete: %#v", resolved.NumberingSource)
	}
}

func TestNativeNumberingGenuineStrictWordprocessingMLFixture(t *testing.T) {
	data := strictNumberingFixture(t)
	pkg, err := openNativeDOCXPackage(data)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(pkg.files["Strict/document.xml"], []byte(wordMLStrict)) || !bytes.Contains(pkg.files["Strict/Numbering.XML"], []byte(wordMLStrict)) || !bytes.Contains(pkg.files["Strict/_rels/document.xml.rels"], []byte(relBaseStrict+"numbering")) {
		t.Fatal("fixture is not a genuine Strict WordprocessingML package and numbering relationship")
	}
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if len(resolved.Paragraphs) != 3 {
		t.Fatalf("Strict fixture paragraph count = %d", len(resolved.Paragraphs))
	}
	styleRTL, first, second := resolved.Paragraphs[0].Numbering, resolved.Paragraphs[1].Numbering, resolved.Paragraphs[2].Numbering
	if styleRTL == nil || styleRTL.Level != 1 || styleRTL.LevelStyleID == nil || *styleRTL.LevelStyleID != "ListStrict" || styleRTL.ResolvedText != "I)" || styleRTL.Alignment != "right" || styleRTL.NumberingTabTwips == nil || *styleRTL.NumberingTabTwips != 900 || styleRTL.LabelStartTwips != 540 || styleRTL.LabelEndTwips != 900 || styleRTL.TextStartTwips != 900 {
		t.Fatalf("Strict style/pStyle/tab/RTL marker projection is not exact: %#v", styleRTL)
	}
	if first == nil || second == nil || first.ResolvedText != "1." || first.CounterValue != 1 || second.ResolvedText != "2." || second.CounterValue != 2 {
		t.Fatalf("Strict concrete numId counters changed: first=%#v second=%#v", first, second)
	}
	if resolved.NumberingSource == nil || resolved.NumberingSource.RelationshipsPart != "Strict/_rels/document.xml.rels" || resolved.NumberingSource.RelationshipID != "rIdNumberingStrict" || resolved.NumberingSource.RelationshipType != relBaseStrict+"numbering" || resolved.NumberingSource.RelationshipTarget != "Numbering.XML" || resolved.NumberingSource.PartName != "Strict/Numbering.XML" || resolved.NumberingSource.ModelSHA256 != nativeResolvedNumberingModelSHA256(resolved) {
		t.Fatalf("Strict numbering relationship/model attestation is not exact: %#v", resolved.NumberingSource)
	}
}

func TestNativeNumberingParagraphStyleSelectsExactLinkedAbstractLevel(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl>
	<w:lvl w:ilvl="2"><w:pStyle w:val="BaseList"/><w:start w:val="4"/><w:numFmt w:val="upperRoman"/><w:lvlText w:val="%3)"/><w:pPr><w:ind w:left="1080" w:hanging="360"/></w:pPr></w:lvl>
</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="BaseList"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="ListDeep"><w:basedOn w:val="BaseList"/></w:style></w:styles>`
	parts := resolvedNumberingTestParts(numbering)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="styles" Type="`+relBaseTransitional+`styles" Target="styles.xml"/></Relationships>`, 1)
	parts["word/styles.xml"] = styles
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="ListDeep"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p><w:p><w:pPr><w:pStyle w:val="ListDeep"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>direct</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	styled, direct := resolved.Paragraphs[0].Numbering, resolved.Paragraphs[1].Numbering
	if styled == nil || styled.Level != 2 || styled.LevelStyleID == nil || *styled.LevelStyleID != "BaseList" || styled.ResolvedText != "IV)" {
		t.Fatalf("style ilvl was trusted instead of exact abstract pStyle mapping: %#v", styled)
	}
	if direct == nil || direct.Level != 0 || direct.LevelStyleID != nil || direct.ResolvedText != "1." {
		t.Fatalf("direct paragraph numPr semantics changed: %#v", direct)
	}
	originalDefinition := styled.DefinitionSHA256
	*styled.LevelStyleID = "ForgedStyle"
	resolved.Paragraphs[0].StyleID = nativeString("ForgedStyle")
	resolved.NumberingSource.ModelSHA256 = nativeResolvedNumberingModelSHA256(resolved)
	if styled.DefinitionSHA256 != originalDefinition || ValidateNativeResolvedLayoutInputV1(resolved) == nil {
		t.Fatal("style-level selection provenance tamper with recomputed model digest was accepted")
	}
}

func TestNativeNumberingParagraphStyleLevelMappingRefusalsAreBounded(t *testing.T) {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:styleId="ListDeep"><w:pPr><w:numPr><w:ilvl w:val="8"/><w:numId w:val="2"/></w:numPr></w:pPr></w:style></w:styles>`
	for _, test := range []struct {
		name, levels, code string
	}{
		{"missing", `<w:lvl w:ilvl="0"><w:pStyle w:val="Other"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl>`, "MISSING_NUMBERING_STYLE_LEVEL"},
		{"ambiguous", `<w:lvl w:ilvl="0"><w:pStyle w:val="ListDeep"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:pStyle w:val="ListDeep"/><w:pPr><w:ind w:left="720" w:hanging="180"/></w:pPr></w:lvl>`, "AMBIGUOUS_NUMBERING_STYLE_LEVEL"},
		{"duplicate", `<w:lvl w:ilvl="0"><w:pStyle w:val="ListDeep"/><w:pStyle w:val="Other"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl>`, "DUPLICATE_NUMBERING_STYLE_LEVEL"},
	} {
		t.Run(test.name, func(t *testing.T) {
			numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1">` + test.levels + `</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
			parts := resolvedNumberingTestParts(numbering)
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
			parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="styles" Type="`+relBaseTransitional+`styles" Target="styles.xml"/></Relationships>`, 1)
			parts["word/styles.xml"] = styles
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="ListDeep"/></w:pPr><w:r><w:t>styled</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if resolved.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(resolved, test.code) {
				t.Fatalf("%s style mapping was guessed or silent: %#v", test.name, resolved)
			}
		})
	}
}

func TestNativeNumberingUnusedDefinitionsDoNotEmitSemanticDiagnostics(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:numPicBullet w:numPicBulletId="9"/><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num><w:abstractNum w:abstractNumId="3"><w:unknown/><w:lvl w:ilvl="0"><w:numFmt w:val="decimal" w:format="custom"/><w:lvlText w:val="%1"/><w:lvlPicBulletId w:val="9"/><w:isLgl/><w:legacy/><w:mystery/></w:lvl></w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="3"/><w:mystery/></w:num></w:numbering>`
	unused := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "ordinary"))
	for _, code := range []string{"PICTURE_BULLET_PRESERVED", "CUSTOM_NUMBER_FORMAT", "UNMODELED_NUMBERING_LEVEL", "UNMODELED_ABSTRACT_NUMBERING", "UNMODELED_NUMBERING_INSTANCE"} {
		if hasResolutionDiagnostic(unused, code) {
			t.Fatalf("unused definition emitted %s: %#v", code, unused.Diagnostics)
		}
	}
	referenced := resolveNumberingFixture(t, numbering, numberedParagraph("4", 0, "unsupported"))
	if referenced.Paragraphs[0].Numbering != nil {
		t.Fatalf("referenced picture/custom level was modeled: %#v", referenced.Paragraphs[0].Numbering)
	}
	for _, code := range []string{"PICTURE_BULLET_PRESERVED", "CUSTOM_NUMBER_FORMAT", "UNMODELED_NUMBERING_LEVEL", "UNMODELED_ABSTRACT_NUMBERING", "UNMODELED_NUMBERING_INSTANCE"} {
		if !hasResolutionDiagnostic(referenced, code) {
			t.Fatalf("referenced definition omitted %s: %#v", code, referenced.Diagnostics)
		}
	}
	for _, diagnostic := range referenced.Diagnostics {
		if diagnostic.ScopeID == referenced.DocumentID {
			t.Fatalf("referenced numbering semantic remained document-scoped: %#v", diagnostic)
		}
	}
}

func TestNativeNumberingRootSemanticsDeferUntilConcreteNumIDReference(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `" xmlns:x="urn:foreign-numbering"><w:globalSemantic w:val="opaque"/><x:foreignGlobal x:mode="opaque"/><x:abstractNum x:abstractNumId="spoofed-but-preserved"/><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	unused := resolveNumberingFixture(t, numbering, `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`)
	for _, code := range []string{"UNMODELED_NUMBERING_ROOT", "FOREIGN_NUMBERING_ROOT"} {
		if hasResolutionDiagnostic(unused, code) {
			t.Fatalf("unused numbering-root semantic emitted %s: %#v", code, unused.Diagnostics)
		}
	}

	referenced := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "numbered"))
	if referenced.Paragraphs[0].Numbering == nil {
		t.Fatalf("ordinary marker semantics disappeared instead of remaining visibly refused downstream: %#v", referenced)
	}
	for _, code := range []string{"UNMODELED_NUMBERING_ROOT", "FOREIGN_NUMBERING_ROOT"} {
		if !hasResolutionDiagnostic(referenced, code) {
			t.Fatalf("referenced numbering-root semantic omitted %s: %#v", code, referenced.Diagnostics)
		}
	}
	for _, diagnostic := range referenced.Diagnostics {
		if (diagnostic.Code == "UNMODELED_NUMBERING_ROOT" || diagnostic.Code == "FOREIGN_NUMBERING_ROOT") && diagnostic.ScopeID != referenced.Paragraphs[0].ParagraphID {
			t.Fatalf("numbering-root semantic was not scoped to its concrete reference: %#v", diagnostic)
		}
	}
}

func TestNativeResolvedTwipsAreBoundedBeforeMilliPointConversion(t *testing.T) {
	maximum := nativeMaxTwipsForMilliPoints
	properties := NativeResolvedParagraphPropertiesV1{
		SpacingBeforeTwips: nativeInt64(maximum),
		IndentStartTwips:   nativeInt64(-maximum),
	}
	if err := validateNativeResolvedParagraphProperties(properties); err != nil {
		t.Fatalf("exact maximum convertible twips were rejected: %v", err)
	}
	properties.SpacingBeforeTwips = nativeInt64(maximum + 1)
	if err := validateNativeResolvedParagraphProperties(properties); err == nil {
		t.Fatal("nonnegative twips beyond the exact x50 boundary were accepted")
	}
	properties.SpacingBeforeTwips = nativeInt64(maximum)
	properties.IndentStartTwips = nativeInt64(-maximum - 1)
	if err := validateNativeResolvedParagraphProperties(properties); err == nil {
		t.Fatal("signed twips beyond the exact x50 boundary were accepted")
	}
}

func TestNativeNumberingSyntheticOOXMLFixture(t *testing.T) {
	data, err := os.ReadFile("testdata/native-numbering-v1.docx")
	if err != nil {
		t.Fatal(err)
	}
	first, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	want := []string{"1.", "1.a)", "1.", "2."}
	if len(first.Paragraphs) != len(want) {
		t.Fatalf("fixture paragraph count = %d; want %d", len(first.Paragraphs), len(want))
	}
	for index, paragraph := range first.Paragraphs {
		if paragraph.Numbering == nil || paragraph.Numbering.ResolvedText != want[index] {
			t.Fatalf("fixture marker %d = %#v; want %q", index, paragraph.Numbering, want[index])
		}
	}
	if first.NumberingSource == nil || first.NumberingSource.PartSHA256 == "" || first.NumberingSource.RelationshipsSHA256 == "" || first.NumberingSource.ModelSHA256 != nativeResolvedNumberingModelSHA256(first) {
		t.Fatalf("fixture numbering provenance is incomplete: %#v", first.NumberingSource)
	}
	encoded, err := EncodeNativeResolvedLayoutInputV1(first)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Contains(encoded, []byte(`"fonts":null`)) || !bytes.Contains(encoded, []byte(`"fonts":[]`)) {
		t.Fatalf("empty resolved collections must remain strict JSON arrays: %s", encoded)
	}
	second, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	reencoded, err := EncodeNativeResolvedLayoutInputV1(second)
	if err != nil || !bytes.Equal(encoded, reencoded) {
		t.Fatalf("fixture resolution is nondeterministic: err=%v", err)
	}
}

func TestValidateNativeResolvedNumberingRejectsProvenanceAndModelSubstitution(t *testing.T) {
	data, err := os.ReadFile("testdata/native-numbering-v1.docx")
	if err != nil {
		t.Fatal(err)
	}
	mutations := []struct {
		name   string
		mutate func(*NativeResolvedLayoutInputV1)
	}{
		{"relationship type", func(input *NativeResolvedLayoutInputV1) {
			input.NumberingSource.RelationshipType = "https://example.invalid/numbering"
		}},
		{"raw part hash", func(input *NativeResolvedLayoutInputV1) {
			input.NumberingSource.PartSHA256 = "sha256:" + strings.Repeat("0", 64)
		}},
		{"relationship target", func(input *NativeResolvedLayoutInputV1) {
			input.NumberingSource.RelationshipTarget = "substituted.xml"
			input.NumberingSource.ModelSHA256 = nativeResolvedNumberingModelSHA256(input)
		}},
		{"model hash", func(input *NativeResolvedLayoutInputV1) {
			input.NumberingSource.ModelSHA256 = "sha256:" + strings.Repeat("0", 64)
		}},
		{"resolved marker text", func(input *NativeResolvedLayoutInputV1) { input.Paragraphs[0].Numbering.ResolvedText = "forged" }},
		{"definition hash", func(input *NativeResolvedLayoutInputV1) { input.Paragraphs[0].Numbering.DefinitionSHA256 = "forged" }},
		{"alignment with recomputed model", func(input *NativeResolvedLayoutInputV1) {
			input.Paragraphs[0].Numbering.Alignment = "left"
			input.NumberingSource.ModelSHA256 = nativeResolvedNumberingModelSHA256(input)
		}},
		{"duplicate counter level", func(input *NativeResolvedLayoutInputV1) {
			marker := input.Paragraphs[1].Numbering
			marker.CounterValues = append(marker.CounterValues, marker.CounterValues[0])
		}},
		{"unsafe label geometry", func(input *NativeResolvedLayoutInputV1) {
			input.Paragraphs[0].Numbering.LabelEndTwips = 9007199254740992
			input.Paragraphs[0].Numbering.TextStartTwips = 9007199254740992
			input.NumberingSource.ModelSHA256 = nativeResolvedNumberingModelSHA256(input)
		}},
	}
	for _, mutation := range mutations {
		t.Run(mutation.name, func(t *testing.T) {
			input, resolveErr := ResolveNativeDocumentLayoutV1(data)
			if resolveErr != nil {
				t.Fatal(resolveErr)
			}
			mutation.mutate(input)
			if err := ValidateNativeResolvedLayoutInputV1(input); err == nil {
				t.Fatal("substituted numbering projection was accepted")
			}
		})
	}
}

func numberedParagraph(numID string, level int, text string) string {
	return fmt.Sprintf(`<w:p><w:pPr><w:numPr><w:ilvl w:val="%d"/><w:numId w:val="%s"/></w:numPr></w:pPr><w:r><w:t>%s</w:t></w:r></w:p>`, level, numID, text)
}

func resolveNumberingFixture(t *testing.T, numbering, paragraphs string) *NativeResolvedLayoutInputV1 {
	t.Helper()
	parts := resolvedNumberingTestParts(numbering)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` + paragraphs + `<w:sectPr/></w:body></w:document>`
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	return resolved
}

func TestNativeNumberingSourceOrderCountersRestartsAndOverrides(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `">
<w:abstractNum w:abstractNumId="3">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="tab"/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Test" w:hAnsi="Test"/><w:sz w:val="20"/></w:rPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:suff w:val="space"/><w:lvlJc w:val="start"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Test" w:hAnsi="Test"/><w:lang w:val="en-US"/></w:rPr></w:lvl>
</w:abstractNum>
<w:num w:numId="5"><w:abstractNumId w:val="3"/></w:num>
<w:num w:numId="6"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>
</w:numbering>`
	paragraphs := numberedParagraph("5", 0, "one") + numberedParagraph("5", 1, "one-a") + numberedParagraph("5", 1, "one-b") + numberedParagraph("5", 0, "two") + numberedParagraph("5", 1, "two-a") + numberedParagraph("6", 0, "override") + numberedParagraph("5", 0, "shared-family")
	resolved := resolveNumberingFixture(t, numbering, paragraphs)
	wantText := []string{"1.", "1.a)", "1.b)", "2.", "2.a)", "1.", "3."}
	wantValues := []int{1, 1, 2, 2, 1, 1, 3}
	if len(resolved.Paragraphs) != len(wantText) {
		t.Fatalf("paragraph count = %d", len(resolved.Paragraphs))
	}
	for index, paragraph := range resolved.Paragraphs {
		if paragraph.Numbering == nil || paragraph.Numbering.ResolvedText != wantText[index] || paragraph.Numbering.CounterValue != wantValues[index] {
			t.Fatalf("marker %d = %#v; want text %q value %d", index, paragraph.Numbering, wantText[index], wantValues[index])
		}
		if paragraph.Numbering.MarkerID == "" || paragraph.Numbering.DefinitionSHA256 == "" || paragraph.Numbering.LabelEndTwips != paragraph.Numbering.TextStartTwips {
			t.Fatalf("marker %d lacks durable provenance or exact geometry: %#v", index, paragraph.Numbering)
		}
	}
	if resolved.NumberingSource == nil || resolved.NumberingSource.PartName != "word/numbering.xml" || resolved.NumberingSource.RelationshipsPart != "word/_rels/document.xml.rels" || resolved.NumberingSource.RelationshipID != "numbering" || resolved.NumberingSource.RelationshipTarget != "numbering.xml" {
		t.Fatalf("numbering relationship closure was not attested: %#v", resolved.NumberingSource)
	}
	if resolved.NumberingSource.PartSHA256 != nativeSHA([]byte(numbering)) || resolved.NumberingSource.ModelSHA256 != nativeResolvedNumberingModelSHA256(resolved) {
		t.Fatalf("numbering hashes do not attest raw bytes and canonical marker model: %#v", resolved.NumberingSource)
	}
}

func TestNativeNumberingCountersAreScopedToConcreteInstances(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `">
<w:abstractNum w:abstractNumId="1">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:pPr><w:ind w:left="720" w:hanging="180"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="10"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="11"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="7"/></w:lvlOverride></w:num>
</w:numbering>`
	paragraphs := numberedParagraph("10", 0, "a1") + numberedParagraph("11", 0, "b7") + numberedParagraph("10", 0, "a2") + numberedParagraph("11", 0, "b8") +
		numberedParagraph("10", 1, "a2a") + numberedParagraph("11", 0, "b9") + numberedParagraph("10", 1, "a2b")
	resolved := resolveNumberingFixture(t, numbering, paragraphs)
	want := []string{"1.", "7.", "2.", "8.", "2.a)", "9.", "2.b)"}
	for index, paragraph := range resolved.Paragraphs {
		if paragraph.Numbering == nil || paragraph.Numbering.ResolvedText != want[index] {
			t.Fatalf("interleaved marker %d = %#v; want %q", index, paragraph.Numbering, want[index])
		}
	}
}

func TestNativeNumberingReplacementLevelsHonorStartsAndWordRestartPolicy(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `">
	<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum>
	<w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
<w:num w:numId="2"><w:abstractNumId w:val="1"/><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:lvlOverride></w:num>
<w:abstractNum w:abstractNumId="3"/>
<w:num w:numId="4"><w:abstractNumId w:val="3"/><w:lvlOverride w:ilvl="1"><w:startOverride w:val="7"/><w:lvl w:ilvl="1"><w:start w:val="4"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:lvlOverride></w:num>
<w:abstractNum w:abstractNumId="5">
  <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl>
  <w:lvl w:ilvl="1"><w:start w:val="1"/><w:lvlRestart w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:pPr><w:ind w:left="720" w:hanging="180"/></w:pPr></w:lvl>
</w:abstractNum>
<w:num w:numId="6"><w:abstractNumId w:val="5"/><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="1"/><w:lvlRestart w:val="0"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:pPr><w:ind w:left="720" w:hanging="180"/></w:pPr></w:lvl></w:lvlOverride></w:num>
</w:numbering>`
	paragraphs := numberedParagraph("1", 1, "base-start") + numberedParagraph("2", 1, "replacement-start") + numberedParagraph("4", 1, "start-override") +
		numberedParagraph("6", 0, "one") + numberedParagraph("6", 1, "one-one") + numberedParagraph("6", 0, "two") + numberedParagraph("6", 1, "two-one")
	resolved := resolveNumberingFixture(t, numbering, paragraphs)
	want := []string{"1)", "1)", "7)", "1.", "1.1", "2.", "2.1"}
	for index, paragraph := range resolved.Paragraphs {
		if paragraph.Numbering == nil || paragraph.Numbering.ResolvedText != want[index] {
			t.Fatalf("replacement marker %d = %#v; want %q", index, paragraph.Numbering, want[index])
		}
	}
}

func TestNativeNumberingRestartRangeAndSpecifiedOrEarlierSemantics(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:pPr><w:ind w:left="180" w:hanging="90"/></w:pPr></w:lvl>
<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2"/><w:pPr><w:ind w:left="360" w:hanging="90"/></w:pPr></w:lvl>
<w:lvl w:ilvl="2"><w:start w:val="1"/><w:lvlRestart w:val="2"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3"/><w:pPr><w:ind w:left="540" w:hanging="90"/></w:pPr></w:lvl>
</w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	paragraphs := numberedParagraph("2", 2, "a") + numberedParagraph("2", 2, "b") + numberedParagraph("2", 1, "l1") + numberedParagraph("2", 2, "c") + numberedParagraph("2", 2, "d") + numberedParagraph("2", 0, "l0") + numberedParagraph("2", 2, "e")
	resolved := resolveNumberingFixture(t, numbering, paragraphs)
	want := []string{"1", "2", "1", "1", "2", "1", "1"}
	for index, paragraph := range resolved.Paragraphs {
		if paragraph.Numbering == nil || paragraph.Numbering.ResolvedText != want[index] {
			t.Fatalf("restart marker %d = %#v; want %q", index, paragraph.Numbering, want[index])
		}
	}

	invalid := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="8"><w:lvlRestart w:val="8"/></w:lvl></w:abstractNum></w:numbering>`
	if _, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(resolvedNumberingTestParts(invalid)))); err == nil || !strings.Contains(err.Error(), "invalid lvlRestart") {
		t.Fatalf("lvlRestart 8 was accepted: %v", err)
	}
}

func TestNativeNumberingParsesAndAttestsNumTabStops(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:suff w:val="tab"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="840"/></w:tabs><w:ind w:left="840" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	resolved := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "tabbed"))
	marker := resolved.Paragraphs[0].Numbering
	if marker == nil || marker.NumberingTabTwips == nil || *marker.NumberingTabTwips != 840 || hasResolutionDiagnostic(resolved, "UNMODELED_PARAGRAPH_PROPERTY") {
		t.Fatalf("ordinary Word num tab was not parsed exactly: marker=%#v diagnostics=%#v", marker, resolved.Diagnostics)
	}
	originalDefinition := marker.DefinitionSHA256
	*marker.NumberingTabTwips = 841
	resolved.NumberingSource.ModelSHA256 = nativeResolvedNumberingModelSHA256(resolved)
	if marker.DefinitionSHA256 != originalDefinition || ValidateNativeResolvedLayoutInputV1(resolved) == nil {
		t.Fatal("numbering tab substitution with a self-consistent model digest was accepted")
	}

	for _, tabs := range []string{
		`<w:tabs><w:tab w:val="num"/></w:tabs>`,
		`<w:tabs><w:tab w:val="left" w:pos="840"/></w:tabs>`,
		`<w:tabs><w:tab w:val="num" w:pos="840"/><w:tab w:val="num" w:pos="900"/></w:tabs>`,
	} {
		malformed := strings.Replace(numbering, `<w:tabs><w:tab w:val="num" w:pos="840"/></w:tabs>`, tabs, 1)
		if _, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(func() map[string]string {
			parts := resolvedNumberingTestParts(malformed)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` + numberedParagraph("2", 0, "bad") + `<w:sectPr/></w:body></w:document>`
			return parts
		}()))); err == nil {
			t.Fatalf("malformed numbering tabs were accepted: %s", tabs)
		}
	}
}

func TestNativeNumberingUnsupportedGeometryReturnsDiagnostic(t *testing.T) {
	numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	resolved := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "no hanging geometry"))
	if resolved.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(resolved, "UNSUPPORTED_NUMBERING_GEOMETRY") {
		t.Fatalf("unsupported marker geometry was emitted or silent: %#v", resolved)
	}
}

func TestNativeNumberingOmittedAndMalformedLevelAlignment(t *testing.T) {
	base := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/>ALIGN<w:pPr><w:ind w:start="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	omitted := resolveNumberingFixture(t, strings.Replace(base, "ALIGN", "", 1), numberedParagraph("2", 0, "default"))
	if marker := omitted.Paragraphs[0].Numbering; marker == nil || marker.Alignment != "left" {
		t.Fatalf("omitted lvlJc did not retain the directional Word default: %#v", marker)
	}
	rtlParagraph := `<w:p><w:pPr><w:bidi/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>rtl default</w:t></w:r></w:p>`
	rtl := resolveNumberingFixture(t, strings.Replace(base, "ALIGN", "", 1), rtlParagraph)
	if marker := rtl.Paragraphs[0].Numbering; marker == nil || marker.Alignment != "right" {
		t.Fatalf("omitted RTL lvlJc did not materialize the right default: %#v", marker)
	}
	for _, alignment := range []string{`<w:lvlJc/>`, `<w:lvlJc w:val="bogus"/>`, `<w:lvlJc w:val="both"/>`} {
		result := resolveNumberingFixture(t, strings.Replace(base, "ALIGN", alignment, 1), numberedParagraph("2", 0, "bad"))
		if result.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(result, "UNSUPPORTED_NUMBER_ALIGNMENT") {
			t.Fatalf("malformed explicit lvlJc was guessed: %s %#v", alignment, result)
		}
	}
}

func TestNativeNumberingFormatsUseWordLetterAndRomanSequences(t *testing.T) {
	letterCases := map[int]string{1: "a", 26: "z", 27: "aa", 52: "zz", 53: "aaa", 806: strings.Repeat("z", 31)}
	for value, want := range letterCases {
		got, ok := nativeFormatAlphabetic(value, false)
		if !ok || got != want {
			t.Fatalf("letter %d = %q, %v; want %q", value, got, ok, want)
		}
	}
	if _, ok := nativeFormatAlphabetic(807, false); ok {
		t.Fatal("letter output beyond the 31-character bound was accepted")
	}
	for _, test := range []struct {
		value int
		upper bool
		want  string
	}{{4, true, "IV"}, {9, false, "ix"}, {4000, true, "MMMM"}} {
		got, ok := nativeFormatRoman(test.value, test.upper)
		if !ok || got != test.want {
			t.Fatalf("Roman %d = %q, %v; want %q", test.value, got, ok, test.want)
		}
	}
}

func TestNativeNumberingLiteralBulletPercentAndStrictPlaceholderRefusals(t *testing.T) {
	bullet := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="•😀é%"/><w:suff w:val="nothing"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	resolved := resolveNumberingFixture(t, bullet, numberedParagraph("2", 0, "bullet"))
	if marker := resolved.Paragraphs[0].Numbering; marker == nil || marker.ResolvedText != "•😀é%" || marker.Format != "bullet" || len(marker.CounterValues) != 0 {
		t.Fatalf("bullet literal was not preserved exactly: %#v", marker)
	}

	literal := `%%|%0|%x|😀é|%1%`
	literalNumbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="` + literal + `"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	literalResult := resolveNumberingFixture(t, literalNumbering, numberedParagraph("2", 0, "literal"))
	if marker := literalResult.Paragraphs[0].Numbering; marker == nil || marker.ResolvedText != "%%|%0|%x|😀é|1%" || len(marker.CounterValues) != 1 {
		t.Fatalf("literal percent text was not preserved exactly: %#v", marker)
	}

	for _, template := range []string{"%10", "%2"} {
		t.Run(template, func(t *testing.T) {
			numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="` + template + `"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
			result := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "bad"))
			if result.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(result, "MALFORMED_NUMBERING_TEXT") {
				t.Fatalf("malformed placeholder %q was guessed or silent: %#v", template, result)
			}
		})
	}

	oversized := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="` + strings.Repeat("•", 32) + `"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
	oversizedResult := resolveNumberingFixture(t, oversized, numberedParagraph("2", 0, "oversized"))
	if oversizedResult.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(oversizedResult, "MALFORMED_NUMBERING_TEXT") {
		t.Fatalf("oversized Unicode bullet was guessed or silent: %#v", oversizedResult)
	}
}

func TestNativeNumberingDefinitionCanonicalDigestVector(t *testing.T) {
	styleID := "List<&>😀"
	marker := NativeResolvedNumberingV1{
		NumID: "7", AbstractNumID: "3", Level: 2, LevelStyleID: &styleID,
		Start: 4, Format: "decimal", Text: "<>&\u2028\u2029😀é%3", Suffix: "space", Alignment: "center",
		RestartAfterLevel: nativeInt(1), NumberingTabTwips: nativeInt64(840),
		LabelStartTwips: 360, LabelEndTwips: 1080, TextStartTwips: 1080,
	}
	got := nativeResolvedNumberingDefinitionSHA256(&marker, "sha256:"+strings.Repeat("a", 64))
	const want = "sha256:654cfe4ae66ee21d8b0e330ec13d58eef947f7fd76378515500cdd5a6489a31d"
	if got != want {
		t.Fatalf("canonical numbering definition digest = %q; want %q", got, want)
	}
}

func TestNativeNumberingRefusesUnsupportedFormatsAndAmbiguousRelationships(t *testing.T) {
	for _, format := range []string{"legal", "ordinal", "cardinalText", "ordinalText"} {
		t.Run(format, func(t *testing.T) {
			numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="` + format + `"/><w:lvlText w:val="%1"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
			result := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "unsupported"))
			if result.Paragraphs[0].Numbering != nil || !hasResolutionDiagnostic(result, "UNSUPPORTED_NUMBER_FORMAT") {
				t.Fatalf("format %q was guessed or silent: %#v", format, result)
			}
		})
	}
	for _, kind := range []string{"singleLevel", "multilevel", "hybridMultilevel"} {
		t.Run(kind, func(t *testing.T) {
			numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:multiLevelType w:val="` + kind + `"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/><w:pPr><w:ind w:left="360" w:hanging="180"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/><w:pPr><w:ind w:left="720" w:hanging="180"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`
			result := resolveNumberingFixture(t, numbering, numberedParagraph("2", 0, "one")+numberedParagraph("2", 1, "one-one"))
			if result.Paragraphs[0].Numbering == nil || result.Paragraphs[1].Numbering == nil || result.Paragraphs[1].Numbering.ResolvedText != "1.1" || hasResolutionDiagnostic(result, "UNSUPPORTED_MULTI_LEVEL_BEHAVIOR") {
				t.Fatalf("valid %s metadata changed actual level semantics: %#v", kind, result)
			}
		})
	}

	parts := resolvedNumberingTestParts(`<w:numbering xmlns:w="` + wordMLTransitional + `"/>`)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="numbering2" Type="`+relBaseTransitional+`numbering" Target="numbering.xml"/></Relationships>`, 1)
	if _, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "multiple numbering relationships") {
		t.Fatalf("duplicate numbering relationship error = %v", err)
	}

	external := resolvedNumberingTestParts(`<w:numbering xmlns:w="` + wordMLTransitional + `"/>`)
	external["word/_rels/document.xml.rels"] = strings.Replace(external["word/_rels/document.xml.rels"], `Target="numbering.xml"`, `Target="https://example.invalid/numbering.xml" TargetMode="External"`, 1)
	if _, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(external))); err == nil || !strings.Contains(err.Error(), "must be internal") {
		t.Fatalf("external numbering relationship error = %v", err)
	}
}
