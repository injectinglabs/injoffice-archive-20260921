package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func TestNativeTextboxPageSource(t *testing.T) {
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
	drawing, err := os.ReadFile("../../../../testdata/docx-native/page-textbox-drawing.xml")
	if err != nil {
		t.Fatal(err)
	}
	parts["word/styles.xml"] = []byte(`<w:styles xmlns:w="` + wns + `"/>`)
	parts["word/document.xml"] = []byte(`<w:document xmlns:w="` + wns + `"><w:body><w:p><w:r>` + string(drawing) + `</w:r><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/></w:rPr><w:t>Body text with a page-placed rectangle.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
	var archive bytes.Buffer
	writer := zip.NewWriter(&archive)
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
	inspection, err := docxpatch.InspectNativePartialSourceV1(source)
	if err != nil {
		t.Fatal(err)
	}
	var inspected struct {
		Document json.RawMessage                           `json:"document"`
		Resolved json.RawMessage                           `json:"resolved_layout"`
		Geometry docxpatch.NativeTextboxGeometryEvidenceV1 `json:"textbox_geometry"`
	}
	if err := json.Unmarshal(inspection, &inspected); err != nil {
		t.Fatal(err)
	}
	if len(inspected.Geometry.Items) != 1 || inspected.Geometry.Items[0].PageAnchor == nil || inspected.Geometry.Items[0].Geometry == nil {
		t.Fatalf("missing page geometry: %s", inspection)
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
	if out := os.Getenv("INJOFFICE_TEXTBOX_PAGE_EVIDENCE_DIR"); out != "" {
		if err := os.MkdirAll(out, 0755); err != nil {
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
		envelope, err := json.Marshal(map[string]any{"document": inspected.Document, "resolved_layout": inspected.Resolved, "textbox_geometry": inspected.Geometry, "pagination_settings": json.RawMessage(settingsJSON), "font_inventory_json": string(inventoryJSON), "font_assets": assets})
		if err != nil {
			t.Fatal(err)
		}
		for name, data := range map[string][]byte{"source.json": envelope, "page-textbox.docx": source} {
			if err := os.WriteFile(filepath.Join(out, name), data, 0600); err != nil {
				t.Fatal(err)
			}
		}
	}
}
