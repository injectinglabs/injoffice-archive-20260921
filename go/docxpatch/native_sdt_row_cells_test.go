package docxpatch

import (
	"strings"
	"testing"
)

// A content control around a table row's cells is common in form templates.
// Its w:sdtContent holds the same w:tc elements, so reading through it must
// produce the row's cells, disclose the wrapper, and keep the table read-only.
func TestNativeExtractStructuredTagRowCells(t *testing.T) {
	document := func(row string) []byte {
		parts := resolvedStylesTestParts(`<w:styles xmlns:w="` + wordMLTransitional + `"/>`)
		parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="1710"/><w:gridCol w:w="1710"/></w:tblGrid>` + row + `</w:tbl><w:p/><w:sectPr/></w:body></w:document>`
		return buildNativeDOCX(t, nativeEntries(parts))
	}
	cell := func(text string) string {
		return `<w:tc><w:tcPr><w:tcW w:w="1710" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>` + text + `</w:t></w:r></w:p></w:tc>`
	}
	codes := func(doc *NativeDocumentV1) []string {
		out := []string{}
		for _, entry := range doc.Unsupported {
			if entry.Capability == "table-structure" {
				out = append(out, entry.Code)
			}
		}
		return out
	}
	texts := func(doc *NativeDocumentV1) []string {
		out := []string{}
		for _, block := range doc.Body.Blocks {
			if block.Table == nil {
				continue
			}
			for _, row := range block.Table.Rows {
				for _, item := range row.Cells {
					for _, paragraph := range item.Paragraphs {
						for _, run := range paragraph.Runs {
							if run.Text != nil {
								out = append(out, *run.Text)
							}
						}
					}
				}
			}
		}
		return out
	}

	t.Run("wrapped cells are laid out and disclosed", func(t *testing.T) {
		doc, err := ExtractNativeDocumentV1(document(`<w:tr><w:sdt><w:sdtPr><w:text/></w:sdtPr><w:sdtContent>` + cell("A1") + cell("B1") + `</w:sdtContent></w:sdt></w:tr>`))
		if err != nil {
			t.Fatal(err)
		}
		if got := strings.Join(texts(doc), ","); got != "A1,B1" {
			t.Fatalf("cell text = %q, want %q", got, "A1,B1")
		}
		if got := strings.Join(codes(doc), ","); got != "WRAPPED_ROW_CELLS" {
			t.Fatalf("table-structure diagnostics = %q, want WRAPPED_ROW_CELLS", got)
		}
		table := doc.Body.Blocks[0].Table
		if table.EditPolicy.Mode != "read-only" {
			t.Fatalf("a table read through a content control must stay read-only, got %#v", table.EditPolicy)
		}
	})

	t.Run("wrapped and direct cells keep source order", func(t *testing.T) {
		doc, err := ExtractNativeDocumentV1(document(`<w:tr>` + cell("A1") + `<w:sdt><w:sdtContent>` + cell("B1") + `</w:sdtContent></w:sdt></w:tr>`))
		if err != nil {
			t.Fatal(err)
		}
		if got := strings.Join(texts(doc), ","); got != "A1,B1" {
			t.Fatalf("cell text = %q, want %q", got, "A1,B1")
		}
	})

	// Reading through a control is lossless only while it holds nothing but
	// cells. Anything else keeps the refusal, because those children would be
	// row content this extractor would silently drop.
	for _, test := range []struct{ name, row string }{
		{"content beside the cells", `<w:tr><w:sdt><w:sdtContent>` + cell("A1") + `<w:customXml/></w:sdtContent></w:sdt></w:tr>`},
		{"no cells at all", `<w:tr><w:sdt><w:sdtContent><w:p/></w:sdtContent></w:sdt></w:tr>`},
		{"two content elements", `<w:tr><w:sdt><w:sdtContent>` + cell("A1") + `</w:sdtContent><w:sdtContent>` + cell("B1") + `</w:sdtContent></w:sdt></w:tr>`},
		{"unmodeled child beside the content", `<w:tr><w:sdt><w:customXml/><w:sdtContent>` + cell("A1") + `</w:sdtContent></w:sdt></w:tr>`},
		{"not a content control", `<w:tr><w:customXml>` + cell("A1") + `</w:customXml></w:tr>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			doc, err := ExtractNativeDocumentV1(document(test.row))
			if err != nil {
				t.Fatal(err)
			}
			if got := strings.Join(codes(doc), ","); !strings.Contains(got, "UNMODELED_ROW_CONTENT") {
				t.Fatalf("table-structure diagnostics = %q, want UNMODELED_ROW_CONTENT", got)
			}
			if got := strings.Join(texts(doc), ","); strings.Contains(got, "A1") {
				t.Fatalf("refused row content must not be exposed as cells, got %q", got)
			}
		})
	}
}
