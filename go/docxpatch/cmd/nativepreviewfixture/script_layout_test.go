package main

import (
	"github.com/injectinglabs/injoffice/go/docxpatch"
	"os"
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
				if block.Paragraph.EditPolicy.Mode != "read-only" {
					t.Fatal("script paragraph gained editing authority")
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
