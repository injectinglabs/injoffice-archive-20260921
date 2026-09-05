package docxpatch

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func qualifiedSyntheticNoteTestDocx(t *testing.T) []byte {
	t.Helper()
	source := buildNoteTestDocx(t)
	documentXML, ok := zipPart(t, source, docPart)
	if !ok {
		t.Fatal("authored structural fixture document.xml missing")
	}
	styles, ok := zipPart(t, source, "word/styles.xml")
	if !ok {
		t.Fatal("authored structural fixture styles.xml missing")
	}
	if !strings.Contains(styles, `w:ascii="Fixture Sans"`) || !strings.Contains(styles, `w:styleId="Heading3"`) {
		t.Fatal("authored structural styles fixture changed")
	}
	rels, _ := zipPart(t, source, docRelsPart)
	relsWithFont, err := appendRelationship(rels, "rIdFixtureFontTable", relBaseTransitional+"fontTable", "fontTable.xml")
	if err != nil {
		t.Fatal(err)
	}
	contentTypeXML, _ := zipPart(t, source, contentTypes)
	contentTypesWithTable, err := overridePartWith(contentTypeXML, "/word/fontTable.xml", "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml")
	if err != nil {
		t.Fatal(err)
	}
	contentTypesWithFont, err := overridePartWith(contentTypesWithTable, "/word/fonts/FixtureSans.odttf", nativeDOCXObfuscatedFontType)
	if err != nil {
		t.Fatal(err)
	}
	contentTypesWithFonts, err := overridePartWith(contentTypesWithFont, "/word/fonts/FixtureSansBold.odttf", nativeDOCXObfuscatedFontType)
	if err != nil {
		t.Fatal(err)
	}
	fontBytes, err := os.ReadFile(filepath.Join("..", "..", "node_modules", "dejavu-fonts-ttf", "ttf", "DejaVuSans.ttf"))
	if err != nil {
		t.Fatal(err)
	}
	boldFontBytes, err := os.ReadFile(filepath.Join("..", "..", "node_modules", "dejavu-fonts-ttf", "ttf", "DejaVuSans-Bold.ttf"))
	if err != nil {
		t.Fatal(err)
	}
	fontTable := `<w:fonts xmlns:w="` + wordMLTransitional + `" xmlns:r="` + relNSTransitional + `"><w:font w:name="Fixture Sans"><w:embedRegular r:id="rIdFixtureFont" w:fontKey="` + nativeFontTestKey + `" w:subsetted="false"/><w:embedBold r:id="rIdFixtureFontBold" w:fontKey="` + nativeFontTestKey + `" w:subsetted="false"/></w:font></w:fonts>`
	fontRels := `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="rIdFixtureFont" Type="` + relBaseTransitional + `font" Target="fonts/FixtureSans.odttf"/><Relationship Id="rIdFixtureFontBold" Type="` + relBaseTransitional + `font" Target="fonts/FixtureSansBold.odttf"/></Relationships>`
	qualified, err := ApplyPatch(source, Patch{
		Replace: map[string][]byte{
			docPart:           []byte(documentXML),
			"word/styles.xml": []byte(styles),
			docRelsPart:       []byte(relsWithFont),
			contentTypes:      []byte(contentTypesWithFonts),
		},
		Add: map[string][]byte{
			"word/fontTable.xml":               []byte(fontTable),
			"word/_rels/fontTable.xml.rels":    []byte(fontRels),
			"word/fonts/FixtureSans.odttf":     nativeFontObfuscate(fontBytes),
			"word/fonts/FixtureSansBold.odttf": nativeFontObfuscate(boldFontBytes),
		},
		Delete: map[string]bool{},
	})
	if err != nil {
		t.Fatal(err)
	}
	return qualified
}

func TestNoteWriterStructuralFixture(t *testing.T) {
	for _, kind := range []NoteKind{Footnote, Endnote} {
		t.Run(string(kind), func(t *testing.T) {
			withNote, _, err := InsertNote(buildNoteTestDocx(t), kind, 0, "Synthetic fixture note.")
			if err != nil {
				t.Fatal(err)
			}
			document, err := ExtractNativeDocumentV1(withNote)
			if err != nil {
				t.Fatal(err)
			}
			if len(document.Notes) != 3 {
				t.Fatalf("structural note stories = %d, want two sentinels and one content note", len(document.Notes))
			}
			resolved, err := ResolveNativeDocumentLayoutV1(withNote)
			if err != nil {
				t.Fatal(err)
			}
			foundVerticalRefusal := false
			for _, diagnostic := range resolved.Diagnostics {
				foundVerticalRefusal = foundVerticalRefusal || diagnostic.Code == "VERTICAL_ALIGNMENT_UNSUPPORTED"
			}
			if !foundVerticalRefusal {
				t.Fatalf("Word superscript authoring must remain explicit unsupported native evidence: %#v", resolved.Diagnostics)
			}
			notesPart := "word/" + string(kind) + "s.xml"
			notes, ok := zipPart(t, withNote, notesPart)
			if !ok || strings.Contains(notes, ">---</") || !strings.Contains(notes, "<w:separator/>") || !strings.Contains(notes, "<w:continuationSeparator/>") {
				t.Fatalf("%s is not structural-only instruction sentinel output: %s", notesPart, notes)
			}
			styles, ok := zipPart(t, withNote, "word/styles.xml")
			wantText, wantReference := "FootnoteText", "FootnoteReference"
			if kind == Endnote {
				wantText, wantReference = "EndnoteText", "EndnoteReference"
			}
			if !ok || !strings.Contains(styles, `w:styleId="`+wantText+`"`) || !strings.Contains(styles, `w:styleId="`+wantReference+`"`) {
				t.Fatalf("note styles are dangling: %s", styles)
			}
		})
	}
}
