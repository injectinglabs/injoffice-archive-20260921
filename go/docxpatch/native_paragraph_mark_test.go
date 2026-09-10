package docxpatch

import (
	"strings"
	"testing"
)

func paragraphMarkPackage(t *testing.T, properties string) []byte {
	t.Helper()
	return buildNativeDOCX(t, nativeEntries(map[string]string{
		"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="main" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":   `<w:document xmlns:w="` + testW + `"><w:body><w:p><w:pPr>` + properties + `</w:pPr><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>Body</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	}))
}

func TestNativeParagraphMarkFormattingIsRenderOnly(t *testing.T) {
	data := paragraphMarkPackage(t, `<w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="22"/><w:b w:val="false"/><w:color w:val="123456"/><w:lang w:val="en-US"/></w:rPr>`)
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	for _, unsupported := range doc.Unsupported {
		if strings.Contains(unsupported.Code, "PARAGRAPH") || strings.Contains(unsupported.Code, "RUN_PROPERTIES") {
			t.Fatalf("qualified mark refused: %+v", unsupported)
		}
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph.EditPolicy.Mode == "read-write" {
		t.Fatal("mark qualification broadened editing authority")
	}
	resolved, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	mark := resolved.Paragraphs[0].ParagraphMarkProperties
	if mark.FontSizeHalfPoint == nil || *mark.FontSizeHalfPoint != 22 || mark.FontFamily == nil || *mark.FontFamily != "DejaVu Sans" {
		t.Fatalf("lost mark metrics: %+v", mark)
	}
	if resolved.Runs[0].Properties.FontSizeHalfPoint == nil || *resolved.Runs[0].Properties.FontSizeHalfPoint != 44 {
		t.Fatal("mark properties leaked into body run")
	}
}

func TestNativeParagraphMarkUnknownSourcesRemainRefused(t *testing.T) {
	for _, properties := range []string{
		`<w:rPr/><w:rPr/>`,
		`<w:rPr><w:sz w:val="22"/><w:sz w:val="24"/></w:rPr>`,
		`<w:rPr><w:rStyle w:val="Unknown"/></w:rPr>`,
		`<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>`,
		`<w:rPr><w:sz w:val="0"/></w:rPr>`,
		`<w:rPr><w:b w:val="perhaps"/></w:rPr>`,
		`<w:rPr><w:sz w:val="22" custom="1"/></w:rPr>`,
		`<w:rPr><w:sz w:val="22"><w:b/></w:sz></w:rPr>`,
		`<w:rPr custom="1"/>`,
		`<w:rPr><x:b xmlns:x="urn:foreign"/></w:rPr>`,
	} {
		t.Run(properties, func(t *testing.T) {
			doc, err := ExtractNativeDocumentV1(paragraphMarkPackage(t, properties))
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, unsupported := range doc.Unsupported {
				if strings.Contains(unsupported.Code, "PARAGRAPH") || strings.Contains(unsupported.Code, "RUN_PROPERTIES") {
					found = true
				}
			}
			if !found {
				t.Fatal("unsafe paragraph mark accepted")
			}
			if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode == "read-write" {
				t.Fatal("unsafe paragraph became writable")
			}
		})
	}
}
