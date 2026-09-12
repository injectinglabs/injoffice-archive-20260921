package pptxpatch

import (
	"bytes"
	"os"
	"path/filepath"
	"testing"
)

func TestNativeMutationRunPropertySchemaOrder(t *testing.T) {
	for _, drawing := range []string{nsDrawingTransitional, nsDrawingStrict} {
		encoded, err := encodeNativeParagraphs(nativeMutationParagraphs("ordered"), nativeExtractDialect{drawing: drawing})
		if err != nil {
			t.Fatal(err)
		}
		// CT_TextCharacterProperties requires fill before latin/ea/cs font slots.
		fill, font := bytes.Index(encoded, []byte("<a:solidFill>")), bytes.Index(encoded, []byte("<a:latin "))
		if fill < 0 || font < 0 || fill > font {
			t.Fatal("run property sequence violates DrawingML schema")
		}
	}
}

func TestNativeOptionalSchemaValidationMutationFixture(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_SDK_FIXTURE_DIR")
	if dir == "" {
		t.Skip("optional local SDK package validation")
	}
	input, err := os.ReadFile(filepath.Join(dir, "slots-en-US-false.pptx"))
	if err != nil {
		t.Fatal(err)
	}
	deck, err := ExtractNativePPTX(input, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	e := deck.Slides[0].Elements[0]
	paragraphs := nativeMutationParagraphs("Schema valid replacement")
	result, err := ApplyNativePPTXMutations(input, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{{OperationID: "replace", Kind: NativePPTXReplaceText, ElementID: e.ID, ExpectedFingerprintSHA256: e.Source.FingerprintSHA256, Paragraphs: &paragraphs}}})
	if err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "mutation-schema-order.pptx"), result, 0600); err != nil {
		t.Fatal(err)
	}
}
