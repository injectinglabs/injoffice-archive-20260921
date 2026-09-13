package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

// Synthetic OOXML source fixture, not a Word visual reference. Keeps real
// package/font/relationship provenance through extraction and optional paint QA.
func TestNativeFootnoteParagraphChainSource(t *testing.T) {
	font, err := os.ReadFile("../../../../node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf")
	if os.IsNotExist(err) {
		t.Skip("optional installed DejaVu font unavailable")
	}
	if err != nil {
		t.Fatal(err)
	}
	base, err := build(font)
	if err != nil {
		t.Fatal(err)
	}
	reader, err := zip.NewReader(bytes.NewReader(base), int64(len(base)))
	if err != nil {
		t.Fatal(err)
	}
	parts := map[string][]byte{}
	for _, entry := range reader.File {
		opened, err := entry.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(opened)
		opened.Close()
		if err != nil {
			t.Fatal(err)
		}
		parts[entry.Name] = data
	}
	parts["[Content_Types].xml"] = []byte(strings.Replace(string(parts["[Content_Types].xml"]), "</Types>", `<Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/></Types>`, 1))
	parts["word/_rels/document.xml.rels"] = []byte(strings.Replace(string(parts["word/_rels/document.xml.rels"]), "</Relationships>", `<Relationship Id="footnotes" Type="`+rns+`/footnotes" Target="footnotes.xml"/></Relationships>`, 1))
	parts["word/settings.xml"] = []byte(strings.Replace(string(parts["word/settings.xml"]), "</w:settings>", `<w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr></w:settings>`, 1))
	parts["word/styles.xml"] = []byte(strings.ReplaceAll(string(parts["word/styles.xml"]), `w:after="120"`, `w:after="0"`))
	parts["word/document.xml"] = []byte(`<w:document xmlns:w="` + wns + `"><w:body><w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Prefix paragraph</w:t></w:r></w:p><w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Reference </w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r></w:p><w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Following paragraph</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="12800" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
	parts["word/footnotes.xml"] = []byte(`<w:footnotes xmlns:w="` + wns + `"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:pPr><w:keepNext/><w:keepLines/></w:pPr><w:r><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> The complete footnote stays with its reference and uses a kept paragraph with source text long enough to wrap.</w:t></w:r></w:p><w:p><w:pPr><w:keepLines/><w:keepNext w:val="false"/></w:pPr><w:r><w:t>Second authored paragraph stays in the same footnote group.</w:t></w:r></w:p></w:footnote></w:footnotes>`)
	var archive bytes.Buffer
	writer := zip.NewWriter(&archive)
	// ApplyPatch provides deterministic entry ordering and preserves all source parts.
	for name, data := range parts {
		out, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := out.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	source, err := docxpatch.ApplyPatch(archive.Bytes(), docxpatch.Patch{})
	if err != nil {
		t.Fatal(err)
	}
	before := append([]byte(nil), source...)
	document, err := docxpatch.ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	resolved, err := docxpatch.ResolveNativeDocumentLayoutV1(source)
	if err != nil {
		t.Fatal(err)
	}
	settings, err := docxpatch.ExtractNativePaginationSettingsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	inventory, err := docxpatch.ExtractNativeDOCXFontInventoryV1(source)
	if err != nil {
		t.Fatal(err)
	}
	assets, err := docxpatch.ResolveNativeDOCXPagePaintFontAssetsV1(source, inventory)
	if err != nil {
		t.Fatal(err)
	}
	if len(document.Unsupported) != 0 || len(resolved.Diagnostics) != 0 || settings.Profile != "word-modern-default" {
		t.Fatalf("fixture was not qualified: source=%#v resolved=%#v settings=%#v", document.Unsupported, resolved.Diagnostics, settings)
	}
	if len(document.Notes) != 3 {
		t.Fatalf("want three source stories, got %d", len(document.Notes))
	}
	labels, kept, sentinels := 0, 0, 0
	for _, story := range document.Notes {
		if story.NoteRole == "content" {
			if len(story.Blocks) != 2 {
				t.Fatal("source paragraphs were dropped")
			}
			for index, block := range story.Blocks {
				if block.Paragraph.Properties.KeepNext == nil || *block.Paragraph.Properties.KeepNext != (index == 0) {
					t.Fatal("authored chain projection lost")
				}
				if block.Paragraph.Properties.KeepLines != nil && *block.Paragraph.Properties.KeepLines {
					kept++
				}
				for _, run := range block.Paragraph.Runs {
					if run.Reference != nil && run.Reference.Role == "label" {
						labels++
					}
				}
			}
		} else {
			sentinels++
			if len(story.Blocks) != 1 || len(story.Blocks[0].Paragraph.Runs) != 0 {
				t.Fatal("instruction sentinel was not preserved exactly")
			}
		}
	}
	if labels != 1 || kept != 2 || sentinels != 2 || !bytes.Equal(before, source) {
		t.Fatalf("source changed: labels=%d kept=%d sentinels=%d", labels, kept, sentinels)
	}
	if out := os.Getenv("INJOFFICE_FOOTNOTE_PARAGRAPHS_EVIDENCE_DIR"); out != "" {
		if err := os.MkdirAll(out, 0755); err != nil {
			t.Fatal(err)
		}
		docJSON, err := docxpatch.EncodeNativeDocumentV1(document)
		if err != nil {
			t.Fatal(err)
		}
		layoutJSON, err := docxpatch.EncodeNativeResolvedLayoutInputV1(resolved)
		if err != nil {
			t.Fatal(err)
		}
		settingsJSON, err := docxpatch.EncodeNativePaginationSettingsV1(settings)
		if err != nil {
			t.Fatal(err)
		}
		inventoryJSON, err := docxpatch.EncodeNativeDOCXFontInventoryV1(inventory)
		if err != nil {
			t.Fatal(err)
		}
		envelope, err := json.Marshal(map[string]any{"document": json.RawMessage(docJSON), "resolved_layout": json.RawMessage(layoutJSON), "pagination_settings": json.RawMessage(settingsJSON), "font_inventory_json": string(inventoryJSON), "font_assets": assets})
		if err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(out, "footnote-paragraphs.docx"), source, 0600); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(out, "source.json"), envelope, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
