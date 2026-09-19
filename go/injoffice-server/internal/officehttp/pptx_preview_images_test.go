package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/binary"
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

// A derived asset names a transform the host runs on bytes it just verified.
// Both ends stay pinned: the part that goes in and the bytes that come out.
func TestPPTXPreviewImageDerivedRasterIsReDerivedAndPinnedAtBothEnds(t *testing.T) {
	source := controlSnapshotMetafile(t)
	rastered, err := pptxpatch.RasterizeNativeMetafilePNG(source)
	if err != nil {
		t.Fatalf("fixture metafile does not rasterise: %v", err)
	}
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	part, err := writer.Create("ppt/media/control.wmf")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := part.Write(source); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	data := buffer.Bytes()

	id := "asset:control"
	sourceLength, rasterLength := int64(len(source)), int64(len(rastered))
	sourceDigest := fmt.Sprintf("%x", sha256.Sum256(source))
	rasterDigest := fmt.Sprintf("%x", sha256.Sum256(rastered))
	transform := pptxpatch.NativeAssetSourceTransformWmfRasterV1
	makeDeck := func() pptxpatch.NativePPTXDeck {
		wmfTransform := transform
		return pptxpatch.NativePPTXDeck{
			Slides: []pptxpatch.NativeSlide{{Elements: []pptxpatch.NativeElement{{Kind: pptxpatch.NativeElementKindPicture, AssetID: &id}}}},
			Assets: []pptxpatch.NativeAsset{{
				ID: id, ContentType: "image/png", SHA256: rasterDigest, ByteLength: &rasterLength,
				SourceTransform: &wmfTransform, SourceByteLength: &sourceLength,
				Source: &pptxpatch.NativeSourceAnchor{PartName: "ppt/media/control.wmf", FingerprintSHA256: sourceDigest},
			}},
		}
	}
	deck := makeDeck()
	if err := attachPPTXPreviewImages(context.Background(), data, &deck, 0); err != nil {
		t.Fatalf("derived asset was not served: %v", err)
	}
	if deck.Assets[0].DataBase64 == nil || *deck.Assets[0].DataBase64 != base64.StdEncoding.EncodeToString(rastered) {
		t.Fatal("served bytes are not the raster the deck stated")
	}

	for _, test := range []struct {
		name   string
		mutate func(*pptxpatch.NativePPTXDeck)
	}{
		// The part must be the one the deck pinned.
		{"wrong-source-digest", func(d *pptxpatch.NativePPTXDeck) {
			d.Assets[0].Source.FingerprintSHA256 = strings.Repeat("0", 64)
		}},
		{"wrong-source-length", func(d *pptxpatch.NativePPTXDeck) {
			length := sourceLength + 1
			d.Assets[0].SourceByteLength = &length
		}},
		// The derivation's result must be the bytes the deck pinned, or this
		// build and the deck disagree and nothing is served.
		{"wrong-derived-digest", func(d *pptxpatch.NativePPTXDeck) {
			d.Assets[0].SHA256 = strings.Repeat("0", 64)
		}},
		{"wrong-derived-length", func(d *pptxpatch.NativePPTXDeck) {
			length := rasterLength + 1
			d.Assets[0].ByteLength = &length
		}},
		{"unknown-transform", func(d *pptxpatch.NativePPTXDeck) {
			unknown := pptxpatch.NativeAssetSourceTransform("somethingElseV1")
			d.Assets[0].SourceTransform = &unknown
		}},
		{"missing-source-length", func(d *pptxpatch.NativePPTXDeck) { d.Assets[0].SourceByteLength = nil }},
		// Without a transform the two digests must agree, which is the rule
		// every asset without a transform has always followed.
		{"transform-dropped", func(d *pptxpatch.NativePPTXDeck) {
			d.Assets[0].SourceTransform = nil
			d.Assets[0].SourceByteLength = nil
		}},
	} {
		t.Run(test.name, func(t *testing.T) {
			bad := makeDeck()
			test.mutate(&bad)
			if err := attachPPTXPreviewImages(context.Background(), data, &bad, 0); err == nil {
				t.Fatal("accepted a derived asset that was not pinned at both ends")
			}
			if bad.Assets[0].DataBase64 != nil {
				t.Fatal("bytes were attached despite the refusal")
			}
		})
	}
}

// A metafile the decoder declines must not be served as anything. A deck can
// only reach this path by claiming a raster the source cannot produce.
func TestPPTXPreviewImageRefusesASourceTransformThatCannotRun(t *testing.T) {
	source := []byte("not a metafile at all")
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	part, err := writer.Create("ppt/media/control.wmf")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := part.Write(source); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	id := "asset:control"
	sourceLength, rasterLength := int64(len(source)), int64(64)
	digest := fmt.Sprintf("%x", sha256.Sum256(source))
	transform := pptxpatch.NativeAssetSourceTransformWmfRasterV1
	deck := pptxpatch.NativePPTXDeck{
		Slides: []pptxpatch.NativeSlide{{Elements: []pptxpatch.NativeElement{{Kind: pptxpatch.NativeElementKindPicture, AssetID: &id}}}},
		Assets: []pptxpatch.NativeAsset{{
			ID: id, ContentType: "image/png", SHA256: strings.Repeat("a", 64), ByteLength: &rasterLength,
			SourceTransform: &transform, SourceByteLength: &sourceLength,
			Source: &pptxpatch.NativeSourceAnchor{PartName: "ppt/media/control.wmf", FingerprintSHA256: digest},
		}},
	}
	if err := attachPPTXPreviewImages(context.Background(), buffer.Bytes(), &deck, 0); err == nil {
		t.Fatal("accepted a source transform that cannot run on its source part")
	}
	if deck.Assets[0].DataBase64 != nil {
		t.Fatal("bytes were attached despite the refusal")
	}
}

// controlSnapshotMetafile is the smallest metafile in the shape PowerPoint
// writes for a control: anisotropic window, one solid fill, EOF.
func controlSnapshotMetafile(t *testing.T) []byte {
	t.Helper()
	record := func(out []byte, function uint16, payload []byte) []byte {
		head := make([]byte, 6)
		binary.LittleEndian.PutUint32(head, uint32(len(payload)/2+3))
		binary.LittleEndian.PutUint16(head[4:], function)
		return append(append(out, head...), payload...)
	}
	words := func(values ...uint16) []byte {
		out := make([]byte, len(values)*2)
		for index, value := range values {
			binary.LittleEndian.PutUint16(out[index*2:], value)
		}
		return out
	}
	var records []byte
	records = record(records, 0x0103, words(8))                    // META_SETMAPMODE, MM_ANISOTROPIC
	records = record(records, 0x020C, words(6, 12))                // META_SETWINDOWEXT 12x6
	records = record(records, 0x02FC, words(0, 0x3412, 0x0056, 0)) // META_CREATEBRUSHINDIRECT, solid
	records = record(records, 0x012D, words(0))                    // META_SELECTOBJECT
	records = record(records, 0x061D, append(words(0x0021, 0x00F0), words(6, 12, 0, 0)...))
	records = record(records, 0x0000, nil) // META_EOF
	header := make([]byte, 18)
	binary.LittleEndian.PutUint16(header[0:], 1)
	binary.LittleEndian.PutUint16(header[2:], 9)
	binary.LittleEndian.PutUint16(header[4:], 0x0300)
	binary.LittleEndian.PutUint32(header[6:], uint32(18+len(records))/2)
	binary.LittleEndian.PutUint16(header[10:], 1)
	return append(header, records...)
}
