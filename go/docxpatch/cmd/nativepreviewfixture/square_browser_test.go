package main

import (
	"archive/zip"
	"bytes"
	"io"
	"os"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func TestNativePreviewSquareWrapBrowserFixture(t *testing.T) {
	output := os.Getenv("INJOFFICE_SQUARE_FIXTURE_OUTPUT")
	if output == "" {
		t.Skip("optional square-wrap source export")
	}
	font, err := os.ReadFile(os.Getenv("INJOFFICE_SQUARE_FIXTURE_FONT"))
	if err != nil {
		t.Fatal(err)
	}
	source, err := buildWithAllRendering(font, true, false, false, "", true)
	if err != nil {
		t.Fatal(err)
	}
	reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var result bytes.Buffer
	writer := zip.NewWriter(&result)
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		data, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		if file.Name == "word/document.xml" {
			xml := string(data)
			at := strings.Index(xml, "<wp:anchor ")
			if at < 0 {
				t.Fatal("floating source missing")
			}
			start := strings.LastIndex(xml[:at], "<w:p>")
			endOffset := strings.Index(xml[at:], "</w:p>")
			if start < 0 || endOffset < 0 {
				t.Fatal("source paragraph missing")
			}
			end := at + endOffset + len("</w:p>")
			paragraph := xml[start:end]
			paragraph = strings.Replace(paragraph, "<wp:wrapNone/>", `<wp:wrapSquare wrapText="bothSides"/>`, 1)
			paragraph = strings.Replace(paragraph, ">3657600<", ">914400<", 1)
			paragraph = strings.Replace(paragraph, ">2743200<", ">914400<", 1)
			paragraph = strings.Replace(paragraph, `cx="1828800" cy="457200"`, `cx="1828800" cy="1270000"`, 1)
			text := `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:sz w:val="24"/></w:rPr><w:t xml:space="preserve">` + strings.Repeat("Source text wraps alongside the square image and returns to full width below it. ", 20) + `</w:t></w:r>`
			paragraph = strings.Replace(paragraph, "</w:p>", text+"</w:p>", 1)
			bodyStart := strings.Index(xml, "<w:body>") + len("<w:body>")
			section := strings.Index(xml, "<w:sectPr>")
			if start < 0 || section < bodyStart {
				t.Fatal("source body geometry missing")
			}
			data = []byte(xml[:bodyStart] + paragraph + xml[section:])
		}
		entry, err := writer.Create(file.Name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = entry.Write(data); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := docxpatch.ExtractNativeDocumentV1(result.Bytes()); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(output, result.Bytes(), 0600); err != nil {
		t.Fatal(err)
	}
}
