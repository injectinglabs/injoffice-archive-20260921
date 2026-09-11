package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/base64"
	"fmt"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func TestPPTXPreviewImageBytesAreSelectedAndSourceBound(t *testing.T) {
	content := []byte("owned source bytes; raster validity is checked by the worker")
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	part, err := writer.Create("ppt/media/owned.png")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := part.Write(content); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	data := buffer.Bytes()
	id, length, digest := "asset:owned", int64(len(content)), fmt.Sprintf("%x", sha256.Sum256(content))
	makeDeck := func() pptxpatch.NativePPTXDeck {
		return pptxpatch.NativePPTXDeck{
			Slides: []pptxpatch.NativeSlide{{Elements: []pptxpatch.NativeElement{{Kind: pptxpatch.NativeElementKindGroup, Children: []pptxpatch.NativeElement{{Kind: pptxpatch.NativeElementKindPicture, AssetID: &id}}}}}},
			Assets: []pptxpatch.NativeAsset{{ID: id, ContentType: "image/png", SHA256: digest, ByteLength: &length, Source: &pptxpatch.NativeSourceAnchor{PartName: "ppt/media/owned.png", FingerprintSHA256: digest}}, {ID: "unselected", ContentType: "image/png"}},
		}
	}
	deck := makeDeck()
	if err := attachPPTXPreviewImages(context.Background(), data, &deck, 0); err != nil {
		t.Fatal(err)
	}
	if deck.Assets[0].DataBase64 == nil || *deck.Assets[0].DataBase64 != base64.StdEncoding.EncodeToString(content) || deck.Assets[1].DataBase64 != nil {
		t.Fatal("selected source bytes drifted or unrelated asset loaded")
	}
	chartDeck := makeDeck()
	chartDeck.Slides[0].Elements[0].Children = []pptxpatch.NativeElement{{Kind: pptxpatch.NativeElementKindChart, Chart: &pptxpatch.NativeOpaqueChart{PreviewAssetID: &id}}}
	if err := attachPPTXPreviewImages(context.Background(), data, &chartDeck, 0); err != nil {
		t.Fatal(err)
	}
	if chartDeck.Assets[0].DataBase64 == nil || *chartDeck.Assets[0].DataBase64 != base64.StdEncoding.EncodeToString(content) || chartDeck.Assets[1].DataBase64 != nil {
		t.Fatal("grouped chart-only preview bytes missing or unrelated asset loaded")
	}
	for _, test := range []struct {
		name   string
		mutate func(*pptxpatch.NativePPTXDeck)
	}{
		{"wrong-digest", func(d *pptxpatch.NativePPTXDeck) {
			d.Assets[0].SHA256 = strings.Repeat("0", 64)
			d.Assets[0].Source.FingerprintSHA256 = d.Assets[0].SHA256
		}},
		{"wrong-anchor", func(d *pptxpatch.NativePPTXDeck) { d.Assets[0].Source.FingerprintSHA256 = strings.Repeat("0", 64) }},
		{"missing-part", func(d *pptxpatch.NativePPTXDeck) { d.Assets[0].Source.PartName = "/etc/passwd" }},
		{"unsupported-type", func(d *pptxpatch.NativePPTXDeck) { d.Assets[0].ContentType = "image/svg+xml" }},
		{"missing-asset", func(d *pptxpatch.NativePPTXDeck) { d.Assets = nil }},
	} {
		t.Run(test.name, func(t *testing.T) {
			bad := makeDeck()
			test.mutate(&bad)
			if err := attachPPTXPreviewImages(context.Background(), data, &bad, 0); err == nil {
				t.Fatal("accepted unbound asset")
			}
		})
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if err := attachPPTXPreviewImages(ctx, data, &deck, 0); err == nil {
		t.Fatal("ignored cancellation")
	}
}
