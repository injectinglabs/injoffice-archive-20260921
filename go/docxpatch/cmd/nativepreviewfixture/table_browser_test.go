package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"io"
	"os"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

// Original fixed-grid document, not an Office-authored fidelity reference.
func TestNativePreviewRepeatingTableBrowserFixture(t *testing.T) {
	output := os.Getenv("INJOFFICE_TABLE_FIXTURE_OUTPUT")
	if output == "" {
		t.Skip("optional native table browser fixture export")
	}
	font, err := os.ReadFile(os.Getenv("INJOFFICE_TABLE_FIXTURE_FONT"))
	if err != nil {
		t.Fatal(err)
	}
	source, err := build(font)
	if err != nil {
		t.Fatal(err)
	}
	run := `<w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/>`
	var body strings.Builder
	body.WriteString(`<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="40" w:type="dxa"/><w:left w:w="40" w:type="dxa"/><w:bottom w:w="40" w:type="dxa"/><w:right w:w="40" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="8" w:color="000000"/><w:left w:val="single" w:sz="8" w:color="000000"/><w:bottom w:val="single" w:sz="8" w:color="000000"/><w:right w:val="single" w:sz="8" w:color="000000"/><w:insideH w:val="single" w:sz="8" w:color="000000"/><w:insideV w:val="single" w:sz="8" w:color="000000"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>`)
	for row := 0; row < 5; row++ {
		body.WriteString(`<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="720" w:hRule="exact"/>`)
		if row == 0 {
			body.WriteString(`<w:tblHeader/>`)
		}
		body.WriteString(`</w:trPr>`)
		for column := 0; column < 2; column++ {
			body.WriteString(`<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/>`)
			if row == 0 {
				body.WriteString(`<w:shd w:val="clear" w:fill="DDEEFF"/>`)
			}
			body.WriteString(`</w:tcPr><w:p><w:pPr><w:spacing w:before="0" w:after="0"/><w:rPr>` + run + `</w:rPr></w:pPr><w:r><w:rPr>` + run + `</w:rPr><w:t>`)
			if row == 0 {
				fmt.Fprintf(&body, "Repeated heading %d", column+1)
			} else {
				fmt.Fprintf(&body, "Body row %d column %d", row, column+1)
			}
			body.WriteString(`</w:t></w:r></w:p></w:tc>`)
		}
		body.WriteString(`</w:tr>`)
	}
	body.WriteString(`</w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="4480" w:orient="landscape"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>`)
	document := []byte(`<w:document xmlns:w="` + wns + `"><w:body>` + body.String() + `</w:body></w:document>`)
	if os.Getenv("INJOFFICE_TABLE_PERCENT") == "2500" {
		document = []byte(strings.Replace(string(document), `<w:tblW w:w="9360" w:type="dxa"/>`, `<w:tblW w:w="2500" w:type="pct"/>`, 1))
	}
	if os.Getenv("INJOFFICE_TABLE_SPLIT") == "true" {
		text := string(document)
		text = strings.ReplaceAll(text, `<w:cantSplit/><w:trHeight w:val="720" w:hRule="exact"/>`, ``)
		text = strings.Replace(text, `<w:trPr><w:tblHeader/>`, `<w:trPr><w:cantSplit/><w:trHeight w:val="720" w:hRule="exact"/><w:tblHeader/>`, 1)
		text = strings.Replace(text, `Body row 1 column 1`, strings.Repeat(`Long natural row `, 100), 1)
		text = strings.ReplaceAll(text, `</w:tcPr><w:p>`, `<w:shd w:val="clear" w:fill="EEF5EE"/></w:tcPr><w:p>`)
		// Header retains its original shading; no duplicate properties.
		text = strings.ReplaceAll(text, `<w:shd w:val="clear" w:fill="DDEEFF"/><w:shd w:val="clear" w:fill="EEF5EE"/>`, `<w:shd w:val="clear" w:fill="DDEEFF"/>`)
		document = []byte(text)
	}
	reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var result bytes.Buffer
	writer := zip.NewWriter(&result)
	for _, entry := range reader.File {
		input, err := entry.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(input)
		input.Close()
		if err != nil {
			t.Fatal(err)
		}
		if entry.Name == "word/document.xml" {
			data = document
		}
		part, err := writer.Create(entry.Name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := part.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	doc, err := docxpatch.ExtractNativeDocumentV1(result.Bytes())
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Body.Blocks) != 1 || doc.Body.Blocks[0].Table == nil || len(doc.Body.Blocks[0].Table.Rows) != 5 || !*doc.Body.Blocks[0].Table.Rows[0].RepeatHeader {
		t.Fatal("source table extraction changed")
	}
	if err := os.WriteFile(output, result.Bytes(), 0600); err != nil {
		t.Fatal(err)
	}
}
