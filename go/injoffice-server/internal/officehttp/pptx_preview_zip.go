package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"errors"
	"fmt"
	"net/url"
	"strings"
	"unicode"
	"unicode/utf8"
)

const (
	pptxPreviewMaxEntries       = 2048
	pptxPreviewMaxPartBytes     = 8 << 20
	pptxPreviewMaxExpandedBytes = 32 << 20
)

// Metadata-only admission before the generic extractor allocates expanded
// parts. CRC, XML and OPC graph validity remain the extractor's authority.
// The entry cap is checked after archive/zip parses the bounded 8 MiB input;
// it is not a strict cap on that parser's temporary metadata allocations.
// Neither this preflight nor synchronous extraction is OS-level isolation.
func preflightPPTXPreviewZIP(ctx context.Context, data []byte) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	if len(data) == 0 || len(data) > 8<<20 {
		return errors.New("PPTX preview ZIP must be 1 byte–8 MiB")
	}
	archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if contextErr := ctx.Err(); contextErr != nil {
		return contextErr
	}
	if err != nil {
		return fmt.Errorf("PPTX preview ZIP metadata: %w", err)
	}
	if len(archive.File) == 0 || len(archive.File) > pptxPreviewMaxEntries {
		return errors.New("PPTX preview ZIP must contain 1–2048 entries")
	}
	aliases := make(map[string]bool, len(archive.File))
	var total uint64
	for _, part := range archive.File {
		if err := ctx.Err(); err != nil {
			return err
		}
		name := part.Name
		directory := part.FileInfo().IsDir()
		if directory {
			name = strings.TrimSuffix(name, "/")
		}
		alias, err := pptxPreviewPartAlias(name)
		if err != nil {
			return err
		}
		if !directory && aliases[alias] {
			return errors.New("PPTX preview ZIP contains duplicate or case/percent-aliased entries")
		}
		if !directory {
			aliases[alias] = true
		}
		if part.UncompressedSize64 > pptxPreviewMaxPartBytes {
			return errors.New("PPTX preview ZIP expanded part exceeds 8 MiB")
		}
		if part.UncompressedSize64 > pptxPreviewMaxExpandedBytes-total {
			return errors.New("PPTX preview ZIP expanded total exceeds 32 MiB")
		}
		total += part.UncompressedSize64
		// Safe directory records are not OPC parts and the extractor skips them.
		// Keep their declared sizes in admission budgets, but do not let their
		// spelling collide with actual parts or other directory records.
		if directory {
			continue
		}
		if part.Flags&1 != 0 || (part.Method != zip.Store && part.Method != zip.Deflate) {
			return errors.New("PPTX preview ZIP contains encrypted or unsupported-compression entries")
		}
		offset, err := part.DataOffset()
		if err != nil || offset < 0 || offset > int64(len(data)) || part.CompressedSize64 > uint64(int64(len(data))-offset) {
			return errors.New("PPTX preview ZIP has invalid local entry bounds")
		}
		if part.UncompressedSize64 > 0 && (part.CompressedSize64 == 0 || part.UncompressedSize64 > part.CompressedSize64*200) {
			return errors.New("PPTX preview ZIP entry exceeds the extractor compression-ratio limit")
		}
	}
	return ctx.Err()
}

// Match the extractor's safe OPC spelling and ASCII case/percent aliases.
// This does not authorize filesystem access: bytes stay inside the upload.
func pptxPreviewPartAlias(name string) (string, error) {
	fail := func() (string, error) { return "", errors.New("PPTX preview ZIP contains an unsafe part name") }
	if name == "" || len(name) > 1024 || strings.HasPrefix(name, "/") || strings.HasSuffix(name, "/") || strings.Contains(name, "\\") || !utf8.ValidString(name) {
		return fail()
	}
	segments := strings.Split(name, "/")
	for i, segment := range segments {
		if segment == "" {
			return fail()
		}
		for j := 0; j < len(segment); j++ {
			b := segment[j]
			if b == '?' || b == '#' || b <= 0x20 || b == 0x7f {
				return fail()
			}
			if b == '%' {
				upperHex := func(b byte) bool { return b >= '0' && b <= '9' || b >= 'A' && b <= 'F' }
				if j+2 >= len(segment) || !upperHex(segment[j+1]) || !upperHex(segment[j+2]) {
					return fail()
				}
				j += 2
			}
		}
		decoded, err := url.PathUnescape(segment)
		if err != nil || !utf8.ValidString(decoded) || decoded == "" || decoded == "." || decoded == ".." || strings.HasSuffix(decoded, ".") || strings.ContainsAny(decoded, "/\\") {
			return fail()
		}
		for _, r := range decoded {
			if r < 0x20 || r == 0x7f || !strings.Contains(segment, "%") && unicode.IsSpace(r) {
				return fail()
			}
		}
		folded := []byte(decoded)
		for j, b := range folded {
			if b >= 'A' && b <= 'Z' {
				folded[j] = b + ('a' - 'A')
			}
		}
		segments[i] = string(folded)
	}
	return strings.Join(segments, "/"), nil
}
