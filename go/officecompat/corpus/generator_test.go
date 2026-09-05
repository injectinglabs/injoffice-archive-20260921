package corpus

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestGenerateIsDeterministicAndCheckDetectsDrift(t *testing.T) {
	root := t.TempDir()
	specs := filepath.Join(root, "specs", "docx")
	if err := os.MkdirAll(specs, 0o755); err != nil {
		t.Fatal(err)
	}
	spec := `{
  "protocol": "injoffice-office-fixture-spec/v1",
  "id": "docx-transitional-minimal",
  "format": "docx",
  "dialect": "transitional",
  "outcome": "accepted",
  "provenance": {"kind":"generated","generator":"officecompat test","source":"hand-authored OOXML","license":"CC0-1.0"},
  "coverage": ["common-text", "transitional"],
  "parts": [
    {"name":"word/document.xml","text":"<document>native</document>"},
    {"name":"[Content_Types].xml","text":"<Types/>"},
    {"name":"_rels/.rels","method":"store","text":"<Relationships/>"}
  ],
  "expectation": {"protocol":"injoffice-expected-native-json/v1","fixtureId":"docx-transitional-minimal","format":"docx","dialect":"transitional","outcome":"accepted","native":{"z":2,"a":1}}
}`
	if err := os.WriteFile(filepath.Join(specs, "minimal.json"), []byte(spec), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := Generate(root); err != nil {
		t.Fatal(err)
	}
	packagePath := filepath.Join(root, "generated", "packages", "docx-transitional-minimal.docx")
	first, err := os.ReadFile(packagePath)
	if err != nil {
		t.Fatal(err)
	}
	if err := Generate(root); err != nil {
		t.Fatal(err)
	}
	second, err := os.ReadFile(packagePath)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, second) {
		t.Fatal("fixture bytes changed across identical generations")
	}
	reader, err := zip.NewReader(bytes.NewReader(first), int64(len(first)))
	if err != nil {
		t.Fatal(err)
	}
	if len(reader.File) != 3 || reader.File[0].Name != "[Content_Types].xml" || reader.File[2].Name != "word/document.xml" {
		t.Fatalf("parts are not deterministically ordered: %+v", reader.File)
	}
	if err := Check(root); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(packagePath, append(first, 0), 0o644); err != nil {
		t.Fatal(err)
	}
	if err := Check(root); err == nil || !strings.Contains(err.Error(), "stale") {
		t.Fatalf("expected stale output failure, got %v", err)
	}
}

func TestGenerateRejectsUnboundedAndInvalidExpectations(t *testing.T) {
	tests := []struct {
		name string
		spec string
		want string
	}{
		{"repeat", refusedSpec(`{"name":"bomb.bin","repeat":{"text":"0123456789","count":500000}}`), "repeat payload exceeds"},
		{"identity", strings.Replace(refusedSpec(`{"name":"safe.bin","text":"x"}`), `"fixtureId":"xlsx-strict-refused"`, `"fixtureId":"wrong"`, 1), "expectation envelope"},
		{"accepted-with-refusal", strings.Replace(refusedSpec(`{"name":"safe.bin","text":"x"}`), `"outcome":"refused"`, `"outcome":"accepted"`, 2), "accepted fixture requires"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			root := t.TempDir()
			if err := os.MkdirAll(filepath.Join(root, "specs", "xlsx"), 0o755); err != nil {
				t.Fatal(err)
			}
			if err := os.WriteFile(filepath.Join(root, "specs", "xlsx", "case.json"), []byte(test.spec), 0o644); err != nil {
				t.Fatal(err)
			}
			if err := Generate(root); err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("want %q, got %v", test.want, err)
			}
		})
	}
}

func TestManifestIsCanonicalAndBounded(t *testing.T) {
	root := t.TempDir()
	if err := Generate(root); err != nil {
		t.Fatal(err)
	}
	data, err := os.ReadFile(filepath.Join(root, "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var manifest Manifest
	if err := json.Unmarshal(data, &manifest); err != nil {
		t.Fatal(err)
	}
	if manifest.Protocol != ManifestProtocol || manifest.Generator != GeneratorVersion || manifest.Fixtures == nil {
		t.Fatalf("unexpected empty manifest: %+v", manifest)
	}
	if !bytes.HasSuffix(data, []byte("\n")) {
		t.Fatal("canonical JSON must end with newline")
	}
}

func refusedSpec(part string) string {
	return `{"protocol":"injoffice-office-fixture-spec/v1","id":"xlsx-strict-refused","format":"xlsx","dialect":"strict","outcome":"refused","provenance":{"kind":"generated","generator":"test","source":"test","license":"CC0-1.0"},"coverage":["refusal"],"parts":[` + part + `],"expectation":{"protocol":"injoffice-expected-native-json/v1","fixtureId":"xlsx-strict-refused","format":"xlsx","dialect":"strict","outcome":"refused","refusal":{"class":"invalid-package","contains":"refused"}}}`
}
