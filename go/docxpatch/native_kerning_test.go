package docxpatch

import (
	"strings"
	"testing"
)

func TestNativeKerningCascade(t *testing.T) {
	for _, ns := range []string{wordMLTransitional, wordMLStrict} {
		styles := `<w:styles xmlns:w="` + ns + `"><w:docDefaults><w:rPrDefault><w:rPr><w:kern w:val="28"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Base"><w:rPr><w:kern w:val="24"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Child"><w:basedOn w:val="Base"/></w:style><w:style w:type="character" w:styleId="Char"><w:rPr><w:kern w:val="20"/></w:rPr></w:style></w:styles>`
		parts := resolvedStylesTestParts(styles)
		parts["word/document.xml"] = `<w:document xmlns:w="` + ns + `"><w:body><w:p><w:pPr><w:pStyle w:val="Child"/><w:rPr><w:kern w:val="18"/></w:rPr></w:pPr><w:r><w:rPr><w:rStyle w:val="Char"/><w:kern w:val="16"/></w:rPr><w:t>AV</w:t></w:r><w:r><w:rPr><w:rStyle w:val="Char"/></w:rPr><w:t>AV</w:t></w:r><w:r><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
		if ns == wordMLStrict {
			for k, v := range parts {
				parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
			}
		}
		data := buildNativeDOCX(t, nativeEntries(parts))
		before := string(data)
		layout, err := ResolveNativeDocumentLayoutV1(data)
		if err != nil {
			t.Fatal(err)
		}
		for i, want := range []int{16, 20, 24} {
			got := layout.Runs[i].Properties.KerningMinSizeHalfPoints
			if got == nil || *got != want {
				t.Fatalf("run%d: %v want%d", i, got, want)
			}
		}
		mark := layout.Paragraphs[0].ParagraphMarkProperties.KerningMinSizeHalfPoints
		if mark == nil || *mark != 18 {
			t.Fatalf("mark %v", mark)
		}
		if string(data) != before {
			t.Fatal("source changed")
		}
		doc, err := ExtractNativeDocumentV1(data)
		if err != nil {
			t.Fatal(err)
		}
		paragraph := doc.Body.Blocks[0].Paragraph
		if paragraph.EditPolicy.Mode != "read-only" {
			t.Fatal("kerning grants editing authority")
		}
		if _, err := ApplyNativeTextMutationsV1(data, doc.Source.PackageSHA256, []NativeDOCXTextMutationV1{{TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256, Text: "replacement"}}); err == nil {
			t.Fatal("replacement discarded kerning metadata")
		}
	}
}

func TestNativeKerningInvalidSource(t *testing.T) {
	for _, value := range []string{``, `w:val="0"`, `w:val="-1"`, `w:val="3277"`, `w:val="10pt"`, `w:val="20" extra="1"`} {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr><w:kern ` + value + `/></w:rPr><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
		layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if layout.Runs[0].Properties.KerningMinSizeHalfPoints != nil {
			t.Fatal("invalid threshold resolved")
		}
		found := false
		for _, d := range layout.Diagnostics {
			if d.Code == "INVALID_KERNING_THRESHOLD" {
				found = true
			}
		}
		if !found {
			t.Fatal("missing diagnostic")
		}
	}
}

func TestNativeKerningDuplicateRefuses(t *testing.T) {
	parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:r><w:rPr><w:kern w:val="12"/><w:kern w:val="24"/></w:rPr><w:t>AV</w:t></w:r></w:p></w:body></w:document>`
	layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if layout.Runs[0].Properties.KerningMinSizeHalfPoints != nil {
		t.Fatal("duplicate threshold resolved")
	}
	found := false
	for _, d := range layout.Diagnostics {
		if d.Code == "DUPLICATE_RUN_PROPERTY" {
			found = true
		}
	}
	if !found {
		t.Fatal("missing duplicate diagnostic")
	}
}

func TestNativeKerningNumbering(t *testing.T) {
	parts := resolvedNumberingTestParts(`<w:numbering xmlns:w="` + wordMLTransitional + `"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:rPr><w:kern w:val="12"/></w:rPr></w:lvl></w:abstractNum><w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num></w:numbering>`)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>text</w:t></w:r></w:p></w:body></w:document>`
	parts["word/numbering.xml"] = strings.Replace(parts["word/numbering.xml"], `<w:rPr>`, `<w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr><w:rPr>`, 1)
	layout, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	numbering := layout.Paragraphs[0].Numbering
	if numbering == nil || numbering.Marker.KerningMinSizeHalfPoints == nil || *numbering.Marker.KerningMinSizeHalfPoints != 12 {
		t.Fatalf("numbering %#v diagnostics %#v", numbering, layout.Diagnostics)
	}
	if layout.Runs[0].Properties.KerningMinSizeHalfPoints != nil {
		t.Fatal("marker kerning leaked into body text")
	}
}
