package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"strconv"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func TestNativeTextboxPageSource(t *testing.T)          { testNativeTextboxPageSource(t, 1) }
func TestNativeMultipleTextboxPagesSource(t *testing.T) { testNativeTextboxPageSource(t, 2) }
func TestNativeRelativeTextboxPagesSource(t *testing.T) { testNativeTextboxPageSource(t, 2, true) }
func TestNativeParityTextboxPagesSource(t *testing.T)   { testNativeTextboxPageSource(t, 2, false) }
func TestNativeStackedTextboxPagesSource(t *testing.T)  { testNativeTextboxPageSource(t, 4) }
func testNativeTextboxPageSource(t *testing.T, count int, relative ...bool) {
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
	if count == 2 {
		second := strings.NewReplacer(`id="1"`, `id="2"`, `name="Rectangle"`, `name="Second rectangle"`, `>914400<`, `>3657600<`, `>1828800<`, `>2743200<`, `FFF2CC`, `DDEEFF`, `Rectangle source`, `Second rectangle`).Replace(string(drawing))
		paragraph := `<w:p><w:r>` + second + `</w:r><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/></w:rPr><w:t>Body text on the second textbox page.</w:t></w:r></w:p>`
		main := strings.Replace(string(parts["word/document.xml"]), `<w:sectPr>`, paragraph+`<w:sectPr>`, 1)
		parts["word/document.xml"] = []byte(strings.Replace(main, `w:bottom="1440"`, `w:bottom="14000"`, 1))
	}
	if len(relative) > 0 && relative[0] {
		main := string(parts["word/document.xml"])
		main = strings.Replace(main, `<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset>`, `<wp:positionH relativeFrom="margin"><wp:align>center</wp:align>`, 1)
		main = strings.Replace(main, `<wp:positionV relativeFrom="page"><wp:posOffset>1828800</wp:posOffset>`, `<wp:positionV relativeFrom="page"><wp:align>center</wp:align>`, 1)
		main = strings.Replace(main, `<wp:positionH relativeFrom="page"><wp:posOffset>3657600</wp:posOffset>`, `<wp:positionH relativeFrom="column"><wp:posOffset>-457200</wp:posOffset>`, 1)
		main = strings.Replace(main, `<wp:positionV relativeFrom="page"><wp:posOffset>2743200</wp:posOffset>`, `<wp:positionV relativeFrom="paragraph"><wp:posOffset>914400</wp:posOffset>`, 1)
		parts["word/document.xml"] = []byte(main)
	}
	if len(relative) > 0 && !relative[0] {
		main := strings.Replace(string(parts["word/document.xml"]), `w:bottom="14000"`, `w:bottom="1440"`, 1)
		main = strings.Replace(main, `</w:p><w:p><w:r>`, `</w:p>`+strings.Repeat(`<w:p><w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/></w:rPr><w:t>Flow.</w:t></w:r></w:p>`, 50)+`<w:p><w:r>`, 1)
		for _, axis := range []string{"H", "V"} {
			for _, offset := range []string{"914400", "1828800", "3657600", "2743200"} {
				main = strings.ReplaceAll(main, `<wp:position`+axis+` relativeFrom="page"><wp:posOffset>`+offset+`</wp:posOffset>`, `<wp:position`+axis+` relativeFrom="insideMargin"><wp:align>center</wp:align>`)
			}
		}
		main = strings.NewReplacer(`cx="2743200" cy="914400"`, `cx="609600" cy="457200"`, `Rectangle source`, `Box`, `>Second rectangle<`, `>Box<`).Replace(main)
		parts["word/document.xml"] = []byte(main)
	}
	if count == 4 {
		var boxes strings.Builder
		for i, rank := range []string{"4294967295", "30", "10", "30"} {
			behind, y, color := "0", "1828800", "DDEEFF"
			if i == 0 {
				behind, y, color = "1", "914400", "FFDDEE"
			}
			if i == 2 {
				color = "DDFFDD"
			}
			if i == 3 {
				color = "FFF2CC"
				y = "2286000"
			}
			box := strings.NewReplacer(`id="1"`, `id="`+strconv.Itoa(i+1)+`"`, `relativeHeight="0"`, `relativeHeight="`+rank+`"`, `behindDoc="0"`, `behindDoc="`+behind+`"`, `>1828800<`, `>`+y+`<`, `FFF2CC`, color).Replace(string(drawing))
			boxes.WriteString(`<w:r>` + box + `</w:r>`)
		}
		main := string(parts["word/document.xml"])
		main = strings.Replace(main, `<w:r>`+string(drawing)+`</w:r>`, boxes.String(), 1)
		parts["word/document.xml"] = []byte(main)
	}
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
	if len(inspected.Geometry.Items) != count || inspected.Geometry.OmittedCount != 0 {
		t.Fatalf("missing page geometry: %s", inspection)
	}
	for _, item := range inspected.Geometry.Items {
		if item.PageAnchor == nil || item.Geometry == nil || item.Owner.Status != "supported" {
			t.Fatalf("incomplete textbox source: %+v", item)
		}
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
	env := "INJOFFICE_TEXTBOX_PAGE_EVIDENCE_DIR"
	if count == 2 {
		env = "INJOFFICE_TEXTBOX_PAGES_EVIDENCE_DIR"
	}
	if len(relative) > 0 && relative[0] {
		env = "INJOFFICE_TEXTBOX_POSITION_EVIDENCE_DIR"
	}
	if count == 4 {
		env = "INJOFFICE_TEXTBOX_STACK_EVIDENCE_DIR"
	}
	if len(relative) > 0 && !relative[0] {
		env = "INJOFFICE_TEXTBOX_PARITY_EVIDENCE_DIR"
	}
	if out := os.Getenv(env); out != "" {
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
