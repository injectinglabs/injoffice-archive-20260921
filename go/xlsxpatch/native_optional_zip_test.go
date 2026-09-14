package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"hash/crc32"
	"io"
	"strings"
	"testing"
)

func withOptionalPackageThumbnail(parts map[string]string) map[string]string {
	out := make(map[string]string, len(parts)+2)
	for name, value := range parts {
		out[name] = value
	}
	out["[Content_Types].xml"] = strings.Replace(out["[Content_Types].xml"],
		`<Default Extension="rels"`,
		`<Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="rels"`,
		1)
	out["_rels/.rels"] = strings.Replace(out["_rels/.rels"],
		`</Relationships>`,
		`<Relationship Id="rThumb" Type="`+relTypePackageThumbnail+`" Target="docProps/thumbnail.wmf"/></Relationships>`,
		1)
	return out
}

type nativeRawZIPEntry struct {
	name         string
	body         []byte
	method       uint16
	uncompressed uint64
	crc          uint32
	flags        uint16
}

func buildNativeZIPWithRaw(t *testing.T, parts map[string]string, raw ...nativeRawZIPEntry) []byte {
	t.Helper()
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	for i := 0; i < len(names); i++ {
		for j := i + 1; j < len(names); j++ {
			if names[j] < names[i] {
				names[i], names[j] = names[j], names[i]
			}
		}
	}
	var buf bytes.Buffer
	writer := zip.NewWriter(&buf)
	for _, name := range names {
		entry, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write([]byte(parts[name])); err != nil {
			t.Fatal(err)
		}
	}
	for _, item := range raw {
		header := &zip.FileHeader{
			Name:               item.name,
			Method:             item.method,
			CRC32:              item.crc,
			CompressedSize64:   uint64(len(item.body)),
			UncompressedSize64: item.uncompressed,
			Flags:              item.flags,
		}
		entry, err := writer.CreateRaw(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write(item.body); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func truncatedDeflateZIPEntry(name string) nativeRawZIPEntry {
	return nativeRawZIPEntry{name: name, method: zip.Deflate}
}

func storeZIPEntry(name string, body []byte, uncompressed uint64, flags uint16) nativeRawZIPEntry {
	return nativeRawZIPEntry{
		name:         name,
		body:         body,
		method:       zip.Store,
		uncompressed: uncompressed,
		crc:          crc32.ChecksumIEEE(body),
		flags:        flags,
	}
}

func TestNativeOptionalThumbnailZIPName(t *testing.T) {
	if !nativeOptionalThumbnailZIPName("docProps/thumbnail.wmf") || !nativeOptionalThumbnailZIPName("DocProps/Thumbnail.JPEG") {
		t.Fatal("package thumbnail names were rejected")
	}
	for _, name := range []string{
		"docProps/thumbnail",
		"docProps/thumbnail.",
		"xl/media/thumbnail.wmf",
		"docProps/core.xml",
		"xl/worksheets/sheet1.xml",
		"docProps/thumbnail.wmf/extra",
	} {
		if nativeOptionalThumbnailZIPName(name) {
			t.Fatalf("non-thumbnail name %q was accepted", name)
		}
	}
}

func TestExtractNativeWorkbookOmitsTruncatedOptionalThumbnail(t *testing.T) {
	parts := withOptionalPackageThumbnail(nativeWorkbookFixture(false))
	data := buildNativeZIPWithRaw(t, parts, truncatedDeflateZIPEntry("docProps/thumbnail.wmf"))
	workbook, err := ExtractNativeWorkbookV1(data)
	if err != nil {
		t.Fatal(err)
	}
	_ = findNativeCell(t, workbook, "7", "A1")
	for _, part := range workbook.PassthroughParts {
		if asciiEqualFold(part.PartName, "docProps/thumbnail.wmf") {
			t.Fatalf("truncated thumbnail bytes were invented as passthrough: %#v", part)
		}
	}
	found := false
	for _, item := range workbook.Unsupported {
		if item.Code == "DRAWING_OR_MEDIA_CONTENT" && item.PartName != nil && *item.PartName == "docProps/thumbnail.wmf" {
			found = true
			if item.Capability != "drawings" || item.Preservation != "preserve-exact" || !strings.Contains(item.Message, "optional package thumbnail omitted") {
				t.Fatalf("omitted thumbnail diagnostic mismatch: %#v", item)
			}
		}
	}
	if !found {
		t.Fatalf("omitted thumbnail was not inventoried: %#v", workbook.Unsupported)
	}
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		t.Fatal(err)
	}
	if _, ok := pkg.files["docProps/thumbnail.wmf"]; ok {
		t.Fatal("truncated thumbnail was stored in the package file map")
	}
}

func TestExtractNativeWorkbookOmitsSizeMismatchedOptionalThumbnail(t *testing.T) {
	parts := withOptionalPackageThumbnail(nativeWorkbookFixture(false))
	payload := []byte("not-a-real-wmf")
	data := buildNativeZIPWithRaw(t, parts, storeZIPEntry("docProps/thumbnail.wmf", payload, uint64(len(payload)+8), 0))
	workbook, err := ExtractNativeWorkbookV1(data)
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, item := range workbook.Unsupported {
		if item.Code == "DRAWING_OR_MEDIA_CONTENT" && item.PartName != nil && *item.PartName == "docProps/thumbnail.wmf" && strings.Contains(item.Message, "optional package thumbnail omitted") {
			found = true
		}
	}
	if !found {
		t.Fatalf("size-mismatched thumbnail was not omitted with a diagnostic: %#v", workbook.Unsupported)
	}
}

func TestExtractNativeWorkbookKeepsCompleteOptionalThumbnail(t *testing.T) {
	parts := withOptionalPackageThumbnail(nativeWorkbookFixture(false))
	payload := []byte("complete-wmf-bytes")
	data := buildNativeZIPWithRaw(t, parts, storeZIPEntry("docProps/thumbnail.wmf", payload, uint64(len(payload)), 0))
	workbook, err := ExtractNativeWorkbookV1(data)
	if err != nil {
		t.Fatal(err)
	}
	found := false
	for _, part := range workbook.PassthroughParts {
		if part.PartName == "docProps/thumbnail.wmf" {
			found = true
			if part.ByteLength == nil || *part.ByteLength != int64(len(payload)) || part.SHA256 != nativeWorkbookDigest(payload) {
				t.Fatalf("complete thumbnail fingerprint mismatch: %#v", part)
			}
		}
	}
	if !found {
		t.Fatal("complete thumbnail was omitted from passthrough")
	}
	for _, item := range workbook.Unsupported {
		if item.PartName != nil && *item.PartName == "docProps/thumbnail.wmf" && strings.Contains(item.Message, "optional package thumbnail omitted") {
			t.Fatalf("complete thumbnail was treated as truncated: %#v", item)
		}
	}
}

func TestExtractNativeWorkbookRefusesTruncatedWorksheet(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	delete(parts, "Sheets/s1.xml")
	data := buildNativeZIPWithRaw(t, parts, truncatedDeflateZIPEntry("Sheets/s1.xml"))
	workbook, err := ExtractNativeWorkbookV1(data)
	if err == nil || workbook != nil || !strings.Contains(err.Error(), `read ZIP entry "Sheets/s1.xml"`) || !strings.Contains(err.Error(), "unexpected EOF") {
		t.Fatalf("truncated worksheet was accepted: workbook=%#v err=%v", workbook, err)
	}
}

func TestExtractNativeWorkbookRefusesTruncatedThumbnailWithoutProof(t *testing.T) {
	t.Run("missing thumbnail relationship", func(t *testing.T) {
		parts := nativeWorkbookFixture(false)
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"],
			`<Default Extension="rels"`,
			`<Default Extension="wmf" ContentType="image/x-wmf"/><Default Extension="rels"`,
			1)
		data := buildNativeZIPWithRaw(t, parts, truncatedDeflateZIPEntry("docProps/thumbnail.wmf"))
		if workbook, err := ExtractNativeWorkbookV1(data); err == nil || workbook != nil || !strings.Contains(err.Error(), "unexpected EOF") {
			t.Fatalf("unproven thumbnail skip: workbook=%#v err=%v", workbook, err)
		}
	})
	t.Run("non-image content type", func(t *testing.T) {
		parts := nativeWorkbookFixture(false)
		parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"],
			`</Relationships>`,
			`<Relationship Id="rThumb" Type="`+relTypePackageThumbnail+`" Target="docProps/thumbnail.wmf"/></Relationships>`,
			1)
		data := buildNativeZIPWithRaw(t, parts, truncatedDeflateZIPEntry("docProps/thumbnail.wmf"))
		if workbook, err := ExtractNativeWorkbookV1(data); err == nil || workbook != nil || !strings.Contains(err.Error(), "unexpected EOF") {
			t.Fatalf("non-image thumbnail skip: workbook=%#v err=%v", workbook, err)
		}
	})
	t.Run("worksheet content type override", func(t *testing.T) {
		parts := withOptionalPackageThumbnail(nativeWorkbookFixture(false))
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"],
			`</Types>`,
			`<Override PartName="/docProps/thumbnail.wmf" ContentType="`+nativeWorksheetContentType+`"/></Types>`,
			1)
		data := buildNativeZIPWithRaw(t, parts, truncatedDeflateZIPEntry("docProps/thumbnail.wmf"))
		if workbook, err := ExtractNativeWorkbookV1(data); err == nil || workbook != nil || !strings.Contains(err.Error(), "unexpected EOF") {
			t.Fatalf("required-type thumbnail skip: workbook=%#v err=%v", workbook, err)
		}
	})
}

func TestExtractNativeWorkbookRefusesEncryptedAndUnsupportedThumbnailCompression(t *testing.T) {
	parts := withOptionalPackageThumbnail(nativeWorkbookFixture(false))
	payload := []byte("wmf")
	encrypted := buildNativeZIPWithRaw(t, parts, storeZIPEntry("docProps/thumbnail.wmf", payload, uint64(len(payload)), 0x1))
	if workbook, err := ExtractNativeWorkbookV1(encrypted); err == nil || workbook != nil || !strings.Contains(err.Error(), `encrypted ZIP entry "docProps/thumbnail.wmf"`) {
		t.Fatalf("encrypted thumbnail was accepted: workbook=%#v err=%v", workbook, err)
	}
	unsupported := buildNativeZIPWithRaw(t, parts, nativeRawZIPEntry{
		name:         "docProps/thumbnail.wmf",
		body:         payload,
		method:       99,
		uncompressed: uint64(len(payload)),
		crc:          crc32.ChecksumIEEE(payload),
	})
	if workbook, err := ExtractNativeWorkbookV1(unsupported); err == nil || workbook != nil || !strings.Contains(err.Error(), `unsupported compression method 99`) {
		t.Fatalf("unsupported thumbnail compression was accepted: workbook=%#v err=%v", workbook, err)
	}
}

func TestExtractNativeWorkbookRefusesSizeMismatchedWorksheet(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	payload := []byte(parts["Sheets/s1.xml"])
	delete(parts, "Sheets/s1.xml")
	data := buildNativeZIPWithRaw(t, parts, storeZIPEntry("Sheets/s1.xml", payload, uint64(len(payload)+4), 0))
	if workbook, err := ExtractNativeWorkbookV1(data); err == nil || workbook != nil {
		t.Fatalf("size-mismatched worksheet was accepted: %#v", workbook)
	} else if !strings.Contains(err.Error(), "Sheets/s1.xml") {
		t.Fatalf("size-mismatched worksheet error = %v", err)
	}
}

func TestNativeZIPTruncationError(t *testing.T) {
	if !nativeZIPTruncationError(io.ErrUnexpectedEOF) || !nativeZIPTruncationError(io.EOF) {
		t.Fatal("EOF truncation errors were not recognized")
	}
	if nativeZIPTruncationError(nil) || nativeZIPTruncationError(io.ErrClosedPipe) {
		t.Fatal("non-truncation errors were treated as skippable")
	}
}
