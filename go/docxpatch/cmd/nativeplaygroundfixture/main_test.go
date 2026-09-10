package main

import (
	"bytes"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func TestPlaygroundFixtureIsMeaningfulDeterministicAndEditable(t *testing.T) {
	first, err := buildPlaygroundFixture()
	if err != nil {
		t.Fatal(err)
	}
	second, err := buildPlaygroundFixture()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, second) {
		t.Fatal("playground DOCX generator is not deterministic")
	}

	document, err := docxpatch.ExtractNativeDocumentV1(first)
	if err != nil {
		t.Fatal(err)
	}
	if got := len(document.Body.Blocks); got != 10 {
		t.Fatalf("body blocks = %d, want 10", got)
	}
	if document.Body.Blocks[4].Table == nil || len(document.Body.Blocks[4].Table.Rows) != 4 {
		t.Fatalf("decision table was not extracted: %#v", document.Body.Blocks[4])
	}
	for _, cell := range document.Body.Blocks[4].Table.Rows[0].Cells {
		if cell.ShadingRGB == nil || *cell.ShadingRGB != "234F78" {
			t.Fatal("white header text requires an explicitly modeled blue background")
		}
	}
	title := document.Body.Blocks[0].Paragraph
	if title == nil || title.EditPolicy.Mode != "read-write" || len(title.Runs) != 1 || title.Runs[0].Text == nil || *title.Runs[0].Text != "Northstar Launch Brief" {
		t.Fatalf("meaningful editable title missing: %#v", title)
	}

	replacement := "Northstar Launch Decision Brief"
	result, err := docxpatch.ApplyNativeTextMutationsV1(first, document.Source.PackageSHA256, []docxpatch.NativeDOCXTextMutationV1{{
		TargetKind:        "run",
		TargetID:          title.Runs[0].ID,
		ExpectedXMLSHA256: title.Runs[0].Anchor.XMLSHA256,
		Text:              replacement,
	}})
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Equal(first, result.Package) || result.Document.Body.Blocks[0].Paragraph.Runs[0].Text == nil || *result.Document.Body.Blocks[0].Paragraph.Runs[0].Text != replacement {
		t.Fatalf("guarded title mutation did not survive exact-byte readback")
	}
	if !strings.Contains(documentXML(), "24 workflows validated") || !strings.Contains(documentXML(), "18 design partners confirmed") {
		t.Fatal("fixture lost its launch decision evidence")
	}
}
