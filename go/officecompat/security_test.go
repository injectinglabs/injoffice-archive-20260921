package officecompat_test

import (
	"encoding/binary"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
)

func TestStructuralInspectRejectsUnsafeAndAmbiguousPartNames(t *testing.T) {
	unsafeNames := []string{
		"/word/document.xml",
		"word/../document.xml",
		"word/./document.xml",
		"word//document.xml",
		`word\document.xml`,
		"word/document.xml.",
		"word/%2E%2E/document.xml",
		"word/%2e%2e/document.xml",
		"word/%2Fdocument.xml",
		"word/%5Cdocument.xml",
		"word/%00document.xml",
		"word/%25document.xml",
		"C:/word/document.xml",
		"scheme:word/document.xml",
		"word/%3Adocument.xml",
		"word/\u00a0document.xml",
		"word/%C2%A0document.xml",
		"word/%E2%80%8Bdocument.xml",
		"word/document.xml?query",
		"word/document.xml#fragment",
		"word/doc<ument.xml",
		"word/doc>ument.xml",
		"word/doc\"ument.xml",
		"word/doc[ument.xml",
		"word/doc]ument.xml",
		"word/doc^ument.xml",
		"word/doc`ument.xml",
		"word/doc{ument.xml",
		"word/doc|ument.xml",
		"word/doc}ument.xml",
		"word/caf\u00e9.xml",
		"word/%09document.xml",
		"word/doc%3Cument.xml",
		"word/doc%3Eument.xml",
		"word/doc%22ument.xml",
		"word/doc%5Bument.xml",
		"word/doc%5Dument.xml",
		"word/doc%5Eument.xml",
		"word/doc%60ument.xml",
		"word/doc%7Bument.xml",
		"word/doc%7Cument.xml",
		"word/doc%7Dument.xml",
		"word/document%2E",
	}
	for _, name := range unsafeNames {
		t.Run(name, func(t *testing.T) {
			entries := minimalOPCRoots()
			entries[name] = []byte("payload")
			if _, err := officecompat.Inspect(buildPackage(t, entries, false, false)); err == nil {
				t.Fatalf("unsafe part name %q was accepted", name)
			}
		})
	}

	for _, names := range [][]string{
		{"word/document.xml", "WORD/document.xml"},
		{"word/document.xml", "word/%64ocument.xml"},
	} {
		entries := minimalOPCRoots()
		for _, name := range names {
			entries[name] = []byte(name)
		}
		if _, err := officecompat.Inspect(buildPackage(t, entries, false, false)); err == nil {
			t.Fatalf("case/escape-equivalent names were accepted: %v", names)
		}
	}

	entries := minimalOPCRoots()
	entries["word/../"] = nil
	if _, err := officecompat.Inspect(buildPackage(t, entries, false, true)); err == nil {
		t.Fatal("unsafe ZIP directory was accepted")
	}

	encodedContentTypes := map[string][]byte{
		"%5BContent_Types%5D.xml": []byte(contentTypes),
		"_rels/.rels":             []byte(emptyRootRels),
	}
	if _, err := officecompat.Inspect(buildPackage(t, encodedContentTypes, false, true)); err == nil {
		t.Fatal("percent-encoded alias of reserved [Content_Types].xml was accepted")
	}

	validPChar := minimalOPCRoots()
	validPChar["custom/a!$&'()+,;=@_~-.bin"] = []byte("payload")
	validPChar["custom/caf%C3%A9.bin"] = []byte("percent-encoded UTF-8 payload")
	if _, err := officecompat.Inspect(buildPackage(t, validPChar, false, true)); err != nil {
		t.Fatalf("valid RFC 3986 pchar part was rejected: %v", err)
	}
}

func TestStructuralInspectRejectsEncryptedAndUnsupportedZIPEntries(t *testing.T) {
	valid := buildPackage(t, minimalOPCRoots(), false, true)
	encrypted := mutateZIPHeaders(t, valid, func(flags, method uint16) (uint16, uint16) {
		return flags | 0x1, method
	})
	if _, err := officecompat.Inspect(encrypted); err == nil || !strings.Contains(err.Error(), "encrypted") {
		t.Fatalf("encrypted ZIP result = %v", err)
	}

	unsupported := mutateZIPHeaders(t, valid, func(flags, _ uint16) (uint16, uint16) {
		return flags, 99
	})
	if _, err := officecompat.Inspect(unsupported); err == nil || !strings.Contains(err.Error(), "unsupported compression") {
		t.Fatalf("unsupported ZIP method result = %v", err)
	}
}

func TestStructuralInspectPreservesCaseFoldedRootRouting(t *testing.T) {
	entries := map[string][]byte{
		"[Content_Types].xml": []byte(contentTypes),
		"_RELS/.RELS":         []byte(emptyRootRels),
	}
	inventory, err := officecompat.Inspect(buildPackage(t, entries, false, true))
	if err != nil {
		t.Fatal(err)
	}
	if len(inventory.Parts) != 2 || inventory.Parts[1].Name != "_RELS/.RELS" {
		t.Fatalf("exact case-preserved root relationship bytes were not retained: %+v", inventory.Parts)
	}
}

func TestStructuralInspectWithLimitsFailsBeforeUnboundedExpansion(t *testing.T) {
	entries := map[string][]byte{
		"[Content_Types].xml": []byte("x"),
		"_rels/.rels":         []byte("y"),
		"custom/data.bin":     []byte("12345"),
	}
	data := buildPackage(t, entries, false, true)

	t.Run("package bytes", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxPackageBytes = uint64(len(data) - 1)
		if _, err := officecompat.InspectWithLimits(data, limits); err == nil {
			t.Fatal("package-byte ceiling was ignored")
		}
	})
	t.Run("part count", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxParts = 2
		if _, err := officecompat.InspectWithLimits(data, limits); err == nil {
			t.Fatal("part-count ceiling was ignored")
		}
	})
	t.Run("part bytes", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxPartBytes = 4
		limits.MaxXMLBytes = 4
		if _, err := officecompat.InspectWithLimits(data, limits); err == nil {
			t.Fatal("part-byte ceiling was ignored")
		}
	})
	t.Run("expanded bytes", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxPartBytes = 5
		limits.MaxXMLBytes = 5
		limits.MaxExpandedBytes = 6
		if _, err := officecompat.InspectWithLimits(data, limits); err == nil {
			t.Fatal("expanded-byte ceiling was ignored")
		}
	})
	t.Run("compression ratio", func(t *testing.T) {
		bombEntries := minimalOPCRoots()
		bombEntries["custom/bomb.bin"] = make([]byte, 64*1024)
		bomb := buildPackage(t, bombEntries, false, false)
		limits := officecompat.DefaultLimits()
		limits.MaxCompressionRatio = 2
		limits.CompressionRatioSlack = 0
		if _, err := officecompat.InspectWithLimits(bomb, limits); err == nil || !strings.Contains(err.Error(), "compression ratio") {
			t.Fatalf("compression-ratio result = %v", err)
		}
	})
	t.Run("invalid envelope", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxXMLTokens = 0
		if _, err := officecompat.InspectWithLimits(data, limits); err == nil {
			t.Fatal("invalid limits were accepted")
		}
	})
}

func TestStructuralMutablePartCapabilitiesAreExactAndLive(t *testing.T) {
	beforeEntries := minimalOPCRoots()
	beforeEntries["word/document.xml"] = []byte("before")
	afterEntries := cloneEntries(beforeEntries)
	afterEntries["word/document.xml"] = []byte("after")
	before := buildPackage(t, beforeEntries, false, true)
	after := buildPackage(t, afterEntries, false, true)

	if _, err := officecompat.CompareUntouchedParts(before, after, []string{"word/document.xml"}); err != nil {
		t.Fatalf("exact live capability failed: %v", err)
	}
	for _, capabilities := range [][]string{
		{"word/missing.xml"},
		{"WORD/document.xml"},
		{"word/document.xml", "word/document.xml"},
		{"word/document.xml", "WORD/document.xml"},
		{"word/%64ocument.xml"},
	} {
		if _, err := officecompat.CompareUntouchedParts(before, after, capabilities); err == nil {
			t.Fatalf("unsafe/stale capabilities were accepted: %v", capabilities)
		}
	}

	withAddition := cloneEntries(afterEntries)
	withAddition["word/new.xml"] = []byte("new")
	afterAddition := buildPackage(t, withAddition, false, true)
	if _, err := officecompat.CompareUntouchedParts(after, afterAddition, []string{"word/new.xml"}); err != nil {
		t.Fatalf("exact addition capability failed: %v", err)
	}
}

func minimalOPCRoots() map[string][]byte {
	return map[string][]byte{
		"[Content_Types].xml": []byte(contentTypes),
		"_rels/.rels":         []byte(emptyRootRels),
	}
}

func mutateZIPHeaders(t *testing.T, source []byte, mutate func(flags, method uint16) (uint16, uint16)) []byte {
	t.Helper()
	data := append([]byte(nil), source...)
	modified := 0
	for offset := 0; offset+12 <= len(data); offset++ {
		signature := binary.LittleEndian.Uint32(data[offset:])
		var flagsOffset, methodOffset int
		switch signature {
		case 0x04034b50:
			flagsOffset, methodOffset = offset+6, offset+8
		case 0x02014b50:
			flagsOffset, methodOffset = offset+8, offset+10
		default:
			continue
		}
		flags := binary.LittleEndian.Uint16(data[flagsOffset:])
		method := binary.LittleEndian.Uint16(data[methodOffset:])
		flags, method = mutate(flags, method)
		binary.LittleEndian.PutUint16(data[flagsOffset:], flags)
		binary.LittleEndian.PutUint16(data[methodOffset:], method)
		modified++
	}
	if modified == 0 {
		t.Fatal("test ZIP contained no local or central headers")
	}
	return data
}
