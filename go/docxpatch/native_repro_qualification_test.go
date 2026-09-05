package docxpatch

import (
	"bytes"
	"crypto/sha256"
	"fmt"
	"os"
	"testing"
)

// TestNativeDOCXCrossRuntimeReproQualification binds the JavaScript
// qualification's extracted-document boundary to a real, deterministic OPC
// package owned by the officecompat corpus. It intentionally checks semantics,
// not ZIP metadata or a second renderer.
func TestNativeDOCXCrossRuntimeReproQualification(t *testing.T) {
	packageBytes, err := os.ReadFile("../officecompat/corpus/generated/packages/docx-inline-png-page-paint.docx")
	if err != nil {
		t.Fatal(err)
	}
	first, err := ExtractNativeDocumentV1(packageBytes)
	if err != nil {
		t.Fatal(err)
	}
	firstJSON, err := EncodeNativeDocumentV1(first)
	if err != nil {
		t.Fatal(err)
	}
	second, err := ExtractNativeDocumentV1(bytes.Clone(packageBytes))
	if err != nil {
		t.Fatal(err)
	}
	secondJSON, err := EncodeNativeDocumentV1(second)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(firstJSON, secondJSON) {
		t.Fatal("identical corpus bytes produced different canonical native DOCX JSON")
	}
	const expected = "b0ce1ea47a0055316093f732085bf212e7688192252c5faa2c08b398c4df4251"
	actual := fmt.Sprintf("%x", sha256.Sum256(firstJSON))
	if actual != expected {
		t.Fatalf("canonical native DOCX semantic digest = %s, want %s", actual, expected)
	}
}
