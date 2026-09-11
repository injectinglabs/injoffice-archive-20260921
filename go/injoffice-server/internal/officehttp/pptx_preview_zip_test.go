package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"encoding/binary"
	"errors"
	"fmt"
	"math/rand"
	"strings"
	"testing"
)

func previewZIP(t *testing.T, headers []zip.FileHeader, payload []byte) []byte {
	t.Helper()
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for _, header := range headers {
		entry, err := writer.CreateHeader(&header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = entry.Write(payload); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return buffer.Bytes()
}

func TestPPTXPreviewZIPExpandedBudgetsBeforeExtraction(t *testing.T) {
	// Repeat a 16 KiB seeded block: a real, valid Deflate ZIP below 8 MiB,
	// with ratio below the generic extractor's 200:1 ceiling.
	block := make([]byte, 16<<10)
	_, _ = rand.New(rand.NewSource(42)).Read(block)
	for _, tc := range []struct {
		name           string
		parts, repeats int
		want           string
	}{
		{"at-total-limit", 4, 512, ""},
		{"expanded-total", 5, 448, "expanded total exceeds 32 MiB"},
		{"expanded-part", 1, 513, "expanded part exceeds 8 MiB"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			headers := make([]zip.FileHeader, tc.parts)
			for i := range headers {
				headers[i] = zip.FileHeader{Name: fmt.Sprintf("ppt/media/%d.bin", i), Method: zip.Deflate}
			}
			data := previewZIP(t, headers, bytes.Repeat(block, tc.repeats))
			if len(data) >= 8<<20 {
				t.Fatalf("fixture is not below upload cap: %d", len(data))
			}
			archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
			if err != nil {
				t.Fatal(err)
			}
			for _, part := range archive.File {
				if part.UncompressedSize64 > part.CompressedSize64*200 {
					t.Fatal("fixture exceeded generic compression-ratio ceiling")
				}
			}
			err = preflightPPTXPreviewZIP(context.Background(), data)
			if tc.want == "" {
				if err != nil {
					t.Fatal(err)
				}
				return
			}
			if err == nil || !strings.Contains(err.Error(), tc.want) {
				t.Fatalf("preflight error = %v, want %s", err, tc.want)
			}
			// The input lacks OPC metadata deliberately: this budget error must occur
			// before the full extractor's missing-content-types refusal.
			_, err = pptxPreviewInput(context.Background(), data, 0, PPTXPreviewOptions{})
			if err == nil || !strings.Contains(err.Error(), tc.want) {
				t.Fatalf("full input skipped preflight: %v", err)
			}
		})
	}
}

func TestPPTXPreviewZIPMetadataRefusals(t *testing.T) {
	for _, tc := range []struct {
		name      string
		names     []string
		encrypted bool
	}{
		{name: "duplicate", names: []string{"ppt/a.xml", "ppt/a.xml"}},
		{name: "caseAlias", names: []string{"ppt/A.xml", "ppt/a.xml"}},
		{name: "percentAlias", names: []string{"ppt/a.xml", "ppt/%61.xml"}},
		{name: "traversal", names: []string{"ppt/../a.xml"}},
		{name: "encodedTraversal", names: []string{"ppt/%2E%2E/a.xml"}},
		{name: "malformedEscape", names: []string{"ppt/%GG.xml"}},
		{name: "encrypted", names: []string{"ppt/a.xml"}, encrypted: true},
	} {
		t.Run(tc.name, func(t *testing.T) {
			headers := make([]zip.FileHeader, len(tc.names))
			for i, name := range tc.names {
				headers[i] = zip.FileHeader{Name: name, Method: zip.Store}
				if tc.encrypted {
					headers[i].Flags = 1
				}
			}
			data := previewZIP(t, headers, []byte("bounded"))
			if err := preflightPPTXPreviewZIP(context.Background(), data); err == nil {
				t.Fatal("unsafe metadata admitted")
			}
		})
	}
	headers := make([]zip.FileHeader, pptxPreviewMaxEntries+1)
	for i := range headers {
		headers[i] = zip.FileHeader{Name: fmt.Sprintf("parts/%d", i), Method: zip.Store}
	}
	if err := preflightPPTXPreviewZIP(context.Background(), previewZIP(t, headers, nil)); err == nil || !strings.Contains(err.Error(), "2048") {
		t.Fatalf("entry cap: %v", err)
	}
	if err := preflightPPTXPreviewZIP(context.Background(), []byte("not a ZIP")); err == nil {
		t.Fatal("malformed ZIP admitted")
	}
	data := previewZIP(t, []zip.FileHeader{{Name: "ppt/a.xml", Method: zip.Store}}, []byte("bounded"))
	unsupported := append([]byte(nil), data...)
	central := bytes.Index(unsupported, []byte{'P', 'K', 1, 2})
	if central < 0 {
		t.Fatal("fixture central directory missing")
	}
	unsupported[8], unsupported[central+10] = 99, 99
	if err := preflightPPTXPreviewZIP(context.Background(), unsupported); err == nil || !strings.Contains(err.Error(), "unsupported-compression") {
		t.Fatalf("unsupported method: %v", err)
	}
	// Corrupt the local header while retaining a parseable central directory.
	data[0] = 0
	if err := preflightPPTXPreviewZIP(context.Background(), data); err == nil {
		t.Fatal("malformed local header admitted")
	}
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	if err := preflightPPTXPreviewZIP(ctx, data); !errors.Is(err, context.Canceled) {
		t.Fatalf("cancellation lost: %v", err)
	}
}

func TestPPTXPreviewZIPDirectoryRecordsAreNotPartAliases(t *testing.T) {
	for _, names := range [][]string{
		{"ppt/", "ppt/", "ppt/a.xml"},
		{"PPT/", "%70pt/", "ppt", "ppt/a.xml"},
	} {
		headers := make([]zip.FileHeader, len(names))
		for i, name := range names {
			headers[i] = zip.FileHeader{Name: name, Method: zip.Store}
		}
		if err := preflightPPTXPreviewZIP(context.Background(), previewZIP(t, headers, nil)); err != nil {
			t.Fatalf("harmless directory records refused: %v", err)
		}
	}
	for _, name := range []string{"../", "ppt/%2E%2E/", "/ppt/", "ppt//"} {
		if err := preflightPPTXPreviewZIP(context.Background(), previewZIP(t, []zip.FileHeader{{Name: name, Method: zip.Store}}, nil)); err == nil {
			t.Fatalf("unsafe directory admitted: %s", name)
		}
	}
	data := previewZIP(t, []zip.FileHeader{{Name: "ppt/", Method: zip.Store}}, nil)
	central := bytes.Index(data, []byte{'P', 'K', 1, 2})
	if central < 0 {
		t.Fatal("central directory missing")
	}
	binary.LittleEndian.PutUint32(data[central+24:central+28], pptxPreviewMaxPartBytes+1)
	if err := preflightPPTXPreviewZIP(context.Background(), data); err == nil || !strings.Contains(err.Error(), "expanded part") {
		t.Fatalf("directory declaration bypassed expanded budget: %v", err)
	}
}
