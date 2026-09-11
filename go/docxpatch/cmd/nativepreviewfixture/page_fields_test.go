package main

import (
	"archive/zip"
	"bytes"
	"github.com/injectinglabs/injoffice/go/docxpatch"
	"io"
	"os"
	"strings"
	"testing"
)

func TestNativePageFieldSource(t *testing.T) {
	font, err := os.ReadFile("../../../../node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf")
	if os.IsNotExist(err) {
		t.Skip("optional installed DejaVu font unavailable")
	}
	if err != nil {
		t.Fatal(err)
	}
	source, err := buildWithFields(font, false, true)
	if err != nil {
		t.Fatal(err)
	}
	doc, err := docxpatch.ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	count := 0
	for _, story := range append(doc.Headers, doc.Footers...) {
		if story.Anchor.XMLSHA256 == "" {
			t.Fatal("source XML root is not digest-bound")
		}
		for _, block := range story.Blocks {
			if block.Paragraph == nil {
				continue
			}
			if block.Paragraph.EditPolicy.Mode != "read-only" {
				t.Fatal("field paragraph gained editing authority")
			}
			for _, run := range block.Paragraph.Runs {
				if run.PageField != "" {
					count++
					if run.Text == nil || *run.Text != "" {
						t.Fatal("stale cached result escaped")
					}
				}
			}
		}
	}
	if count != 4 {
		t.Fatalf("want four PAGE/NUMPAGES fields, got %d", count)
	}
	for _, replacement := range []string{`w:instr=" DATE "`, `w:instr=" PAGE \\* ROMAN "`, `w:instr=" PAGE " w:fldLock="true"`, `w:instr=" PAGE " w:instr=" DATE "`} {
		t.Run(replacement, func(t *testing.T) {
			reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
			if err != nil {
				t.Fatal(err)
			}
			var changed bytes.Buffer
			writer := zip.NewWriter(&changed)
			for _, file := range reader.File {
				opened, err := file.Open()
				if err != nil {
					t.Fatal(err)
				}
				data, err := io.ReadAll(opened)
				opened.Close()
				if err != nil {
					t.Fatal(err)
				}
				if file.Name == "word/header1.xml" {
					data = []byte(strings.Replace(string(data), `w:instr=" PAGE "`, replacement, 1))
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
			invalid, err := docxpatch.ExtractNativeDocumentV1(changed.Bytes())
			if err != nil {
				return
			} // malformed XML may be rejected before projection
			refused := false
			for _, entry := range invalid.Unsupported {
				if entry.Code == "FIELD_SEMANTICS" {
					refused = true
				}
			}
			if !refused {
				t.Fatal("unsupported field instruction was admitted")
			}
		})
	}
}
