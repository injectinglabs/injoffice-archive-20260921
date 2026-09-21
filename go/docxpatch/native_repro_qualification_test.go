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
	const expected = "8cd7d0eddc556ea1f4d0a407e4819d1ec13b13edf6e06fb3a24090b1d621dce4"
	actual := fmt.Sprintf("%x", sha256.Sum256(firstJSON))
	if actual != expected {
		t.Fatalf("canonical native DOCX semantic digest = %s, want %s", actual, expected)
	}
}
