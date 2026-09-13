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
func TestNativeUnequalColumnSource(t *testing.T) {
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
	parts["word/settings.xml"] = []byte(strings.Replace(string(parts["word/settings.xml"]), "<w:compat>", "<w:compat><w:noColumnBalance/>", 1))
	parts["word/styles.xml"] = []byte(strings.ReplaceAll(string(parts["word/styles.xml"]), `w:after="120"`, `w:after="0"`))
	var content strings.Builder
	for index := 1; index <= 9; index++ {
		content.WriteString(fmt.Sprintf(`<w:p><w:pPr><w:keepLines/></w:pPr><w:r><w:t>Column paragraph %d has enough text to wrap at different authored widths.</w:t></w:r></w:p>`, index))
	}
	parts["word/document.xml"] = []byte(`<w:document xmlns:w="` + wns + `"><w:body>` + content.String() + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="12000" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:equalWidth="0" w:num="2"><w:col w:w="3000" w:space="720"/><w:col w:w="5640"/></w:cols></w:sectPr></w:body></w:document>`)
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
	if len(document.Unsupported) != 1 || document.Unsupported[0].Code != "UNEQUAL_SECTION_COLUMNS" || len(resolved.Diagnostics) != 0 || settings.Profile != "word-modern-default" || settings.NoColumnBalance == nil || !*settings.NoColumnBalance {
		t.Fatalf("fixture source profile mismatch: source=%#v resolved=%#v settings=%#v", document.Unsupported, resolved.Diagnostics, settings)
	}
	if len(document.Body.Blocks) != 9 || !bytes.Equal(before, source) {
		t.Fatal("source paragraphs changed")
	}
	if out := os.Getenv("INJOFFICE_UNEQUAL_COLUMN_EVIDENCE_DIR"); out != "" {
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
		if err := os.WriteFile(filepath.Join(out, "unequal-columns.docx"), source, 0600); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filepath.Join(out, "source.json"), envelope, 0600); err != nil {
			t.Fatal(err)
		}
	}
}
