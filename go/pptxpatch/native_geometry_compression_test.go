package pptxpatch

import (
	"bytes"
	"compress/gzip"
	"os"
	"testing"
)

func TestNativePresetCompressionPreservesPinnedBytes(t *testing.T) {
	original, err := os.ReadFile("presetdata/preset-shapes.xml")
	if err != nil {
		t.Fatal(err)
	}
	decoded, err := decodeNativePresetCatalog(nativePresetCatalogGzip)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(decoded, original) {
		t.Fatal("compressed catalog changes original resource")
	}
	// The deterministic generator retains no timestamp, filename or OS metadata.
	if len(nativePresetCatalogGzip) > 65536 || !bytes.Equal(nativePresetCatalogGzip[:10], []byte{31, 139, 8, 0, 0, 0, 0, 0, 2, 255}) {
		t.Fatal("noncanonical gzip metadata or budget")
	}
}
func TestNativePresetCompressionRefusesMalformedOrUnboundedData(t *testing.T) {
	compress := func(value []byte) []byte {
		var b bytes.Buffer
		w := gzip.NewWriter(&b)
		_, _ = w.Write(value)
		if err := w.Close(); err != nil {
			t.Fatal(err)
		}
		return b.Bytes()
	}
	original, err := os.ReadFile("presetdata/preset-shapes.xml")
	if err != nil {
		t.Fatal(err)
	}
	changed := bytes.Clone(original)
	changed[100] ^= 1
	crc := bytes.Clone(nativePresetCatalogGzip)
	crc[len(crc)-8] ^= 1
	for name, data := range map[string][]byte{
		"truncated": nativePresetCatalogGzip[:len(nativePresetCatalogGzip)-1], "crc": crc,
		"trailing":     append(bytes.Clone(nativePresetCatalogGzip), 0),
		"concatenated": append(bytes.Clone(nativePresetCatalogGzip), nativePresetCatalogGzip...),
		"fingerprint":  compress(changed), "inflation": compress(make([]byte, nativePresetCatalogXMLSize+1)),
		"compressedBudget": make([]byte, 65537),
	} {
		t.Run(name, func(t *testing.T) {
			if result, err := decodeNativePresetCatalog(data); err == nil || result != nil {
				t.Fatal("invalid compressed source accepted")
			}
		})
	}
}
