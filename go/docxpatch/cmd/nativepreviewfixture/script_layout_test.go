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

func TestNativeScriptSourceAndResolvedProperties(t *testing.T) {
	font, err := os.ReadFile("../../../../node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf")
	if os.IsNotExist(err) {
		t.Skip("optional DejaVu font unavailable")
	}
	if err != nil {
		t.Fatal(err)
	}
	data, err := buildWithScripts(font, false, false, true)
	if err != nil {
		t.Fatal(err)
	}
	document, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	resolved, err := docxpatch.ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		t.Fatal(err)
	}
	found := map[string]string{}
	for _, block := range document.Body.Blocks {
		if block.Paragraph == nil {
			continue
		}
		for _, run := range block.Paragraph.Runs {
			if run.Properties != nil && run.Properties.VerticalAlignment != nil {
				found[run.ID] = *run.Properties.VerticalAlignment
				if block.Paragraph.EditPolicy.Mode != "read-write" {
					t.Fatal("script decoration blocked text editing")
				}
				result, err := docxpatch.ApplyNativeTextMutationsV1(data, document.Source.PackageSHA256, []docxpatch.NativeDOCXTextMutationV1{{TargetKind: "run", TargetID: run.ID, ExpectedXMLSHA256: run.Anchor.XMLSHA256, Text: "Edited"}})
				if err != nil {
					t.Fatal(err)
				}
				before, after := scriptTestMainXML(t, data), scriptTestMainXML(t, result.Package)
				start, end := *run.Anchor.StartByte, *run.Anchor.EndByte
				raw := before[start:end]
				want := before[:start] + raw[:strings.Index(raw, ">")+1] + "Edited" + raw[strings.LastIndex(raw, "<"):] + before[end:]
				if after != want {
					t.Fatal("script text edit changed preserved properties")
				}
			}
		}
	}
	if len(found) != 2 {
		t.Fatalf("expected two script runs, got %#v", found)
	}
	for _, run := range resolved.Runs {
		if expected, ok := found[run.RunID]; ok {
			if run.Properties.VerticalAlignment == nil || *run.Properties.VerticalAlignment != expected {
				t.Fatalf("lost resolved script property: %+v", run)
			}
		}
	}
	if len(document.Unsupported) != 0 || len(resolved.Diagnostics) != 0 {
		t.Fatalf("qualified text scripts were refused: %+v %+v", document.Unsupported, resolved.Diagnostics)
	}
}

func scriptTestMainXML(t *testing.T, data []byte) string {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range reader.File {
		if file.Name != "word/document.xml" {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		return string(content)
	}
	t.Fatal("missing main part")
	return ""
}
