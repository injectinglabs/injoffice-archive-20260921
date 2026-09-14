package xlsxpatch

import (
	"errors"
	"fmt"
	"io"
	"strings"
)

const relTypePackageThumbnail = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail"

type nativePendingZIPTruncation struct {
	name string
	err  error
}

func nativeWorkbookZIPReadError(name string, readErr, closeErr error, actual int, declared uint64) error {
	if readErr != nil {
		return fmt.Errorf("xlsxpatch: native extract: read ZIP entry %q: %w", name, readErr)
	}
	if closeErr != nil {
		return fmt.Errorf("xlsxpatch: native extract: close ZIP entry %q: %w", name, closeErr)
	}
	if uint64(actual) != declared {
		return fmt.Errorf("xlsxpatch: native extract: ZIP entry %q actual length %d does not match declared uncompressed size %d", name, actual, declared)
	}
	return nil
}

func nativeZIPTruncationError(err error) bool {
	return err != nil && (errors.Is(err, io.ErrUnexpectedEOF) || errors.Is(err, io.EOF))
}

func nativeZIPTruncationCandidate(readErr, closeErr error, actual int, declared uint64) bool {
	if nativeZIPTruncationError(readErr) || nativeZIPTruncationError(closeErr) {
		return true
	}
	return uint64(actual) != declared
}

func nativeOptionalThumbnailZIPName(name string) bool {
	key, err := canonicalOPCPartKey(name)
	if err != nil {
		return false
	}
	lower := asciiLower(key)
	const prefix = "docprops/thumbnail."
	if !strings.HasPrefix(lower, prefix) {
		return false
	}
	ext := lower[len(prefix):]
	return ext != "" && !strings.ContainsAny(ext, "/\\")
}

func nativeThumbnailImageContentType(contentType string) bool {
	lower := asciiLower(contentType)
	return strings.HasPrefix(lower, "image/") && len(lower) > len("image/") && !strings.ContainsAny(lower, " \t\r\n")
}

func nativeRequiredWorkbookPartContentType(contentType string) bool {
	lower := asciiLower(contentType)
	switch {
	case asciiEqualFold(contentType, nativeWorkbookContentType),
		asciiEqualFold(contentType, nativeWorksheetContentType),
		asciiEqualFold(contentType, nativeSharedStringsType),
		asciiEqualFold(contentType, nativeRelationshipsType):
		return true
	case strings.Contains(lower, "spreadsheetml"):
		return true
	default:
		return false
	}
}

func nativeRefuseRequiredZIPTruncations(pending []nativePendingZIPTruncation) error {
	for _, item := range pending {
		if !nativeOptionalThumbnailZIPName(item.name) {
			return item.err
		}
	}
	return nil
}

func nativeOptionalThumbnailProven(name string, files map[string][]byte, index *opcPackageIndex, contentTypes nativeContentTypeRegistry) bool {
	if !nativeOptionalThumbnailZIPName(name) {
		return false
	}
	pkg := &nativeWorkbookPackage{index: index, files: files, contentTypes: contentTypes}
	contentType, err := pkg.contentType(name)
	if err != nil || !nativeThumbnailImageContentType(contentType) || nativeRequiredWorkbookPartContentType(contentType) {
		return false
	}
	relsPart, _, found := index.lookupSpelling("_rels/.rels")
	if !found {
		return false
	}
	relsData, ok := files[relsPart]
	if !ok || len(relsData) == 0 {
		return false
	}
	relationships, relErr := parseRoutingRelationships(relsData)
	if relErr != nil {
		return false
	}
	partKey, keyErr := canonicalOPCPartKey(name)
	if keyErr != nil {
		return false
	}
	matched := false
	for _, relationship := range relationships {
		switch relationship.targetMode {
		case "", "Internal":
		default:
			continue
		}
		resolved, resolveErr := resolveNativeRelationshipTarget("", "", relationship.target)
		if resolveErr != nil {
			continue
		}
		resolvedKey, resolvedErr := canonicalOPCPartKey(resolved)
		if resolvedErr != nil || resolvedKey != partKey {
			continue
		}
		if relationship.relType != relTypePackageThumbnail {
			return false
		}
		matched = true
	}
	return matched
}

func nativeOmitProvenOptionalZIPTruncations(pending []nativePendingZIPTruncation, files map[string][]byte, index *opcPackageIndex, contentTypes nativeContentTypeRegistry) ([]nativePendingZIPTruncation, error) {
	if len(pending) == 0 {
		return nil, nil
	}
	omitted := make([]nativePendingZIPTruncation, 0, len(pending))
	for _, item := range pending {
		if !nativeOptionalThumbnailProven(item.name, files, index, contentTypes) {
			return nil, item.err
		}
		omitted = append(omitted, item)
	}
	return omitted, nil
}

func (extractor *nativeWorkbookExtractor) noteOmittedOptionalZIPParts() error {
	for _, item := range extractor.pkg.omittedOptionalZIP {
		message := "optional package thumbnail omitted because ZIP entry is truncated or size-mismatched; original package remains authoritative and thumbnail bytes were not invented"
		if err := extractor.addUnsupported("DRAWING_OR_MEDIA_CONTENT", "drawings", "workbook", item.name, "", message); err != nil {
			return err
		}
	}
	return nil
}
