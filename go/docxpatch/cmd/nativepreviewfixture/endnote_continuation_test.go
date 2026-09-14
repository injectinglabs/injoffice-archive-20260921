package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

// Synthetic OOXML source fixture, not a Word visual reference. Keeps real
// package/font/relationship provenance through extraction and optional paint QA.
func TestNativeEndnoteContinuationSource(t *testing.T) {
	testNativeNoteContinuationSource(t, "endnote", false, false)
}
func TestNativeFootnoteContinuationSource(t *testing.T) {
	testNativeNoteContinuationSource(t, "footnote", false, false)
}
func TestNativeFootnoteLineContinuationSource(t *testing.T) {
	testNativeNoteContinuationSource(t, "footnote", true, false)
}
func TestNativeFootnoteSharedFlowSource(t *testing.T) {
	testNativeNoteContinuationSource(t, "footnote", true, true)
}
func testNativeNoteContinuationSource(t *testing.T, kind string, splitLines, sharedFlow bool) {
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
	parts["[Content_Types].xml"] = []byte(strings.Replace(string(parts["[Content_Types].xml"]), "</Types>", `<Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/></Types>`, 1))
	parts["word/_rels/document.xml.rels"] = []byte(strings.Replace(string(parts["word/_rels/document.xml.rels"]), "</Relationships>", `<Relationship Id="endnotes" Type="`+rns+`/endnotes" Target="endnotes.xml"/></Relationships>`, 1))
	parts["word/settings.xml"] = []byte(strings.Replace(string(parts["word/settings.xml"]), "</w:settings>", `<w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>`, 1))
	parts["word/styles.xml"] = []byte(strings.ReplaceAll(string(parts["word/styles.xml"]), `w:after="120"`, `w:after="0"`))
	parts["word/document.xml"] = []byte(`<w:document xmlns:w="` + wns + `"><w:body><w:p><w:r><w:t>Reference </w:t></w:r><w:r><w:endnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="13200" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
	if sharedFlow {
		var later strings.Builder
		for index := 2; index <= 9; index++ {
			fmt.Fprintf(&later, `<w:p><w:r><w:t>Later body paragraph %d</w:t></w:r>`, index)
			if index == 3 || index == 5 {
				fmt.Fprintf(&later, `<w:r><w:endnoteReference w:id="%d"/></w:r>`, (index+1)/2)
			}
			later.WriteString(`</w:p>`)
		}
		parts["word/document.xml"] = []byte(strings.Replace(string(parts["word/document.xml"]), `<w:sectPr>`, later.String()+`<w:sectPr>`, 1))
	}
	var content strings.Builder
	count := 7
	if splitLines {
		count = 1
	}
	for index := 1; index <= count; index++ {
		if splitLines {
			content.WriteString(`<w:p><w:pPr><w:keepLines w:val="0"/><w:widowControl/></w:pPr>`)
		} else {
			content.WriteString(`<w:p><w:pPr><w:keepLines/></w:pPr>`)
		}
		if index == 1 {
			content.WriteString(`<w:r><w:endnoteRef/></w:r>`)
		}
		text := fmt.Sprintf(" Endnote paragraph %d", index)
		if splitLines {
			text = strings.Repeat(" Footnote line continuation preserves source text.", 30)
		}
		if index == 2 {
			text = strings.Repeat(" Endnote paragraph 2 with retained lines.", 4)
		}
		content.WriteString(`<w:r><w:t>` + text + `</w:t></w:r></w:p>`)
	}
	parts["word/endnotes.xml"] = []byte(`<w:endnotes xmlns:w="` + wns + `"><w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1">` + content.String() + `</w:endnote></w:endnotes>`)
	if sharedFlow {
		var additional strings.Builder
		for index := 2; index <= 3; index++ {
			fmt.Fprintf(&additional, `<w:endnote w:id="%d"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t> Additional note %d</w:t></w:r></w:p></w:endnote>`, index, index)
		}
		parts["word/endnotes.xml"] = []byte(strings.Replace(string(parts["word/endnotes.xml"]), `</w:endnotes>`, additional.String()+`</w:endnotes>`, 1))
	}
	if kind == "footnote" {
		converted := map[string][]byte{}
		for name, data := range parts {
			if strings.HasSuffix(name, ".xml") || strings.HasSuffix(name, ".rels") {
				data = []byte(strings.NewReplacer("endnote", "footnote", "Endnote", "Footnote").Replace(string(data)))
			}
			converted[strings.ReplaceAll(name, "endnote", "footnote")] = data
		}
		parts = converted
	}
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
	expectedNotes, expectedLabels := 3, 1
	if sharedFlow {
		expectedNotes, expectedLabels = 5, 3
	}
	if len(document.Notes) != expectedNotes {
		t.Fatalf("unexpected source story count: %d", len(document.Notes))
	}
	labels, kept, sentinels := 0, 0, 0
	for _, story := range document.Notes {
		if story.NoteRole == "content" {
			if len(story.Blocks) != count {
				t.Fatal("source paragraphs were dropped")
			}
			for _, block := range story.Blocks {
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
	expectedKept := 7
	if splitLines {
		expectedKept = 0
	}
	if labels != expectedLabels || kept != expectedKept || sentinels != 2 || !bytes.Equal(before, source) {
		t.Fatalf("source changed: labels=%d kept=%d sentinels=%d", labels, kept, sentinels)
	}
	envKind := strings.ToUpper(kind)
	if sharedFlow {
		envKind += "_FLOW"
	} else if splitLines {
		envKind += "_LINE"
	}
	if out := os.Getenv("INJOFFICE_" + envKind + "_EVIDENCE_DIR"); out != "" {
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
		if err := os.WriteFile(filepath.Join(out, kind+"-continuation.docx"), source, 0600); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(out, "source.json"), envelope, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
