package main

import (
	"bytes"
	"os"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func TestBounds(t *testing.T) {
	for _, value := range [][]byte{nil, make([]byte, 31), make([]byte, 8*1024*1024+1)} {
		if _, err := build(value); err == nil {
			t.Fatal("accepted invalid font")
		}
	}
}

func TestRealEmbeddedFontFixture(t *testing.T) {
	// CI runs npm ci before the TS/browser job. Pure Go jobs need no npm fonts.
	fontPath := "../../../../node_modules/dejavu-fonts-ttf/ttf/DejaVuSans.ttf"
	font, err := os.ReadFile(fontPath)
	if os.IsNotExist(err) {
		t.Skip("optional installed DejaVu fixture font unavailable")
	}
	if err != nil {
		t.Fatal(err)
	}
	first, err := build(font)
	if err != nil {
		t.Fatal(err)
	}
	second, err := build(font)
	if err != nil || !bytes.Equal(first, second) {
		t.Fatalf("non-deterministic fixture: %v", err)
	}
	doc, err := docxpatch.ExtractNativeDocumentV1(first)
	if err != nil || len(doc.Body.Blocks) != 4 {
		t.Fatalf("expected four paragraphs: %v", err)
	}
	inventory, err := docxpatch.ExtractNativeDOCXFontInventoryV1(first)
	if err != nil {
		t.Fatal(err)
	}
	assets, err := docxpatch.ResolveNativeDOCXPagePaintFontAssetsV1(first, inventory)
	if err != nil || len(assets) != 1 {
		t.Fatalf("expected one embeddable font: %v", err)
	}
}
