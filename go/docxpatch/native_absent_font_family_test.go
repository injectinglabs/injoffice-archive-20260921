package docxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// nativeAbsentFamilyParts is a package with no styles part at all: the shape the
// implicit-default-family evidence is scoped to.
func nativeAbsentFamilyParts(body string) map[string]string {
	return map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":            `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` + body + `<w:sectPr/></w:body></w:document>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"></Relationships>`,
	}
}

// A package that states no font anywhere records the omission per scope, and
// strict resolution still invents nothing.
func TestAbsentFontFamilyRecordsOmissionAndInventsNothing(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		parts := nativeAbsentFamilyParts(`<w:p><w:r><w:t>Source</w:t></w:r></w:p>`)
		if ns == wordMLStrict {
			for k, v := range parts {
				parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
			}
		}
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := bytes.Clone(data)
		eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
		if err != nil {
			t.Fatal(err)
		}
		if len(eligibility.AbsentFontFamilies) != 2 {
			t.Fatalf("facts=%d want 2: %#v", len(eligibility.AbsentFontFamilies), eligibility.AbsentFontFamilies)
		}
		kinds := map[string]bool{}
		for _, fact := range eligibility.AbsentFontFamilies {
			kinds[fact.ScopeKind] = true
			if fact.PackageSHA256 != eligibility.PackageSHA256 || fact.PartName != "word/document.xml" || fact.ScopeID == "" || !strings.HasPrefix(fact.Path, "/w:document[1]/w:body[1]/w:p[1]") {
				t.Fatalf("unbound fact %#v", fact)
			}
		}
		if !kinds["run"] || !kinds["paragraph-mark"] {
			t.Fatalf("missing scope kind: %#v", eligibility.AbsentFontFamilies)
		}
		layout, err := ResolveNativeDocumentLayoutV1(data)
		if err != nil {
			t.Fatal(err)
		}
		if layout.Runs[0].Properties.FontFamily != nil || layout.Paragraphs[0].ParagraphMarkProperties.FontFamily != nil {
			t.Fatal("strict family was invented")
		}
		if len(nativeDOCXFontReferences(layout)) != 0 {
			t.Fatal("strict inventory gained a reference")
		}
		if !bytes.Equal(data, before) {
			t.Fatal("source mutated")
		}
	}
}

// A package that states a font anywhere keeps using what it states: no evidence
// is produced, even when the statement is one strict resolution refuses to read.
func TestAbsentFontFamilyRefusesAnyStatedFontSelection(t *testing.T) {
	for _, tc := range []struct {
		name, body, styles string
	}{
		{name: "run states ascii/hAnsi", body: `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr><w:t>Source</w:t></w:r></w:p>`},
		{name: "paragraph mark states a face", body: `<w:p><w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:pPr><w:r><w:t>Source</w:t></w:r></w:p>`},
		{name: "unresolvable script-only selection", body: `<w:p><w:r><w:rPr><w:rFonts w:eastAsia="SimSun"/></w:rPr><w:t>Source</w:t></w:r></w:p>`},
		{name: "unmodelled selection", body: `<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" unknown="x"/></w:rPr><w:t>Source</w:t></w:r></w:p>`},
		{name: "theme slot with no theme part", body: `<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/></w:rPr><w:t>Source</w:t></w:r></w:p>`},
		{name: "empty selection element", body: `<w:p><w:r><w:rPr><w:rFonts/></w:rPr><w:t>Source</w:t></w:r></w:p>`},
		{name: "docDefaults state a face", styles: `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`},
		{name: "style states a face", styles: `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/></w:rPr></w:style></w:styles>`},
	} {
		t.Run(tc.name, func(t *testing.T) {
			var parts map[string]string
			if tc.styles != "" {
				parts = resolvedStylesTestParts(tc.styles)
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:t>Source</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			} else {
				parts = nativeAbsentFamilyParts(tc.body)
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(eligibility.AbsentFontFamilies) != 0 {
				t.Fatalf("stated font selection produced evidence: %#v", eligibility.AbsentFontFamilies)
			}
		})
	}
}

// A numbered paragraph also needs a marker face this evidence does not cover, so
// none of its scopes qualify.
// A numbered paragraph is covered whole, marker included. The whole-package
// gate already proves the marker carries no family either -- it refuses any
// package whose resolved list marker has one -- so the marker is the same
// omission as the paragraph mark beside it and gets its own scope. Before it
// had one the paragraph was skipped whole, which left listWithLgl.docx, where
// every paragraph is numbered, with no evidence at all.
func TestAbsentFontFamilyCoversTheListMarkerScope(t *testing.T) {
	level := `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>`
	for _, tc := range []struct {
		name, numID string
		want        []string
	}{
		// The numbered paragraph contributes a mark, a marker and its run; the
		// plain paragraph beside it contributes a mark and its run.
		{name: "numbered", numID: "1", want: []string{"paragraph-mark", "numbering-marker", "run", "paragraph-mark", "run"}},
		// A reference that resolves to no marker paints no marker, so there is
		// no marker scope to approximate -- and the paragraph is still covered.
		{name: "unresolved instance", numID: "7", want: []string{"paragraph-mark", "run", "paragraph-mark", "run"}},
	} {
		t.Run(tc.name, func(t *testing.T) {
			numbering := `<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="0">` + level + `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`
			parts := resolvedNumberingTestParts(numbering)
			parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="` + tc.numID + `"/></w:numPr></w:pPr><w:r><w:t>Numbered</w:t></w:r></w:p><w:p><w:r><w:t>Plain</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
			data := buildNativeDOCX(t, nativeEntries(parts))
			eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
			if err != nil {
				t.Fatal(err)
			}
			kinds := []string{}
			for _, fact := range eligibility.AbsentFontFamilies {
				kinds = append(kinds, fact.ScopeKind)
			}
			if strings.Join(kinds, ",") != strings.Join(tc.want, ",") {
				t.Fatalf("scopes=%v want %v: %#v", kinds, tc.want, eligibility.AbsentFontFamilies)
			}
			// The marker shares its paragraph's anchor and id, so the pair
			// (scope_kind, scope_id) is what keeps the two facts distinct.
			for _, fact := range eligibility.AbsentFontFamilies {
				if fact.ScopeKind != "numbering-marker" {
					continue
				}
				mark := eligibility.AbsentFontFamilies[0]
				if fact.ScopeID != mark.ScopeID || fact.Path != mark.Path || fact.PartName != mark.PartName {
					t.Fatalf("marker scope does not name its own paragraph: %#v vs %#v", fact, mark)
				}
			}
		})
	}
}

// The glossary document (ECMA-376 17.12.6) is a separate document with its own
// stories and styles; its font selection states nothing about this one.
func TestAbsentFontFamilyIgnoresGlossaryDocumentFontSelection(t *testing.T) {
	parts := nativeAbsentFamilyParts(`<w:p><w:r><w:t>Source</w:t></w:r></w:p>`)
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/glossary/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml"/></Types>`, 1)
	parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="glossary" Type="`+relBaseTransitional+`glossaryDocument" Target="glossary/document.xml"/></Relationships>`, 1)
	parts["word/glossary/document.xml"] = `<w:glossaryDocument xmlns:w="` + wordMLTransitional + `"><w:docParts><w:docPart><w:docPartBody><w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorHAnsi"/></w:rPr><w:t>Placeholder</w:t></w:r></w:p></w:docPartBody></w:docPart></w:docParts></w:glossaryDocument>`
	data := buildNativeDOCX(t, nativeEntries(parts))
	eligibility, err := ExtractNativeDocxApproximationEligibilityV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if len(eligibility.AbsentFontFamilies) != 2 {
		t.Fatalf("glossary font selection suppressed main-document evidence: %#v", eligibility.AbsentFontFamilies)
	}
}
