package officecompat_test

import (
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

func TestStructuralQualifyCheckedInAcceptedCorpus(t *testing.T) {
	manifestBytes, err := os.ReadFile(filepath.Join("corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var manifest corpus.Manifest
	if err := json.Unmarshal(manifestBytes, &manifest); err != nil {
		t.Fatal(err)
	}
	accepted := manifest
	accepted.Fixtures = nil
	for _, fixture := range manifest.Fixtures {
		if fixture.Outcome == "accepted" {
			accepted.Fixtures = append(accepted.Fixtures, fixture)
		}
	}
	if len(accepted.Fixtures) == 0 {
		t.Fatal("checked-in corpus has no accepted native Office fixtures")
	}

	qualifications, err := officecompat.QualifyCorpusManifest("corpus", accepted)
	if err != nil {
		t.Fatal(err)
	}
	if len(qualifications) != len(accepted.Fixtures) {
		t.Fatalf("qualified %d fixtures, want %d", len(qualifications), len(accepted.Fixtures))
	}
	for index, qualification := range qualifications {
		fixture := accepted.Fixtures[index]
		if qualification.FixtureID != fixture.ID || len(qualification.Inventory.Parts) != fixture.Parts || len(qualification.XMLParts) == 0 {
			t.Fatalf("checked-in fixture %q lacks exact structural evidence: %+v", fixture.ID, qualification)
		}
	}
}

func TestStructuralQualifyCorpusManifestProtocol(t *testing.T) {
	root := t.TempDir()
	entries := minimalOPCRoots()
	entries["word/document.xml"] = []byte(`<w:document xmlns:w="urn:word"><w:body><w:p/></w:body></w:document>`)
	packageBytes := buildPackage(t, entries, false, true)
	expectedBytes := []byte("{}\n")
	packageRelative := "generated/packages/docx-transitional-structural.docx"
	expectedRelative := "generated/expected/docx-transitional-structural.json"
	for relative, data := range map[string][]byte{packageRelative: packageBytes, expectedRelative: expectedBytes} {
		filename := filepath.Join(root, filepath.FromSlash(relative))
		if err := os.MkdirAll(filepath.Dir(filename), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filename, data, 0o644); err != nil {
			t.Fatal(err)
		}
	}
	expanded := 0
	for _, payload := range entries {
		expanded += len(payload)
	}
	manifest := corpus.Manifest{
		Protocol:  corpus.ManifestProtocol,
		Generator: corpus.GeneratorVersion,
		Fixtures: []corpus.FixtureRecord{{
			ID: "docx-transitional-structural", Format: "docx", Dialect: "transitional", Outcome: "accepted",
			Package: packageRelative, Expected: expectedRelative,
			SHA256: sha256String(packageBytes), ExpectedSHA256: sha256String(expectedBytes),
			Bytes: len(packageBytes), ExpandedBytes: expanded, Parts: len(entries),
			Coverage:   []string{"structural"},
			Provenance: corpus.Provenance{Kind: "generated", Generator: "test", Source: "in-memory test", License: "CC0-1.0"},
		}},
	}
	first, err := officecompat.QualifyCorpusManifest(root, manifest)
	if err != nil {
		t.Fatal(err)
	}
	second, err := officecompat.QualifyCorpusManifest(root, manifest)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("corpus qualification is nondeterministic:\nfirst:  %+v\nsecond: %+v", first, second)
	}
	if len(first) != 1 || first[0].FixtureID != manifest.Fixtures[0].ID || len(first[0].Inventory.Parts) != len(entries) || len(first[0].XMLParts) != len(entries) {
		t.Fatalf("unexpected corpus qualification: %+v", first)
	}
	for _, part := range first[0].XMLParts {
		if part.SHA256 == "" || part.StructuralSHA256 == "" {
			t.Fatalf("part lacks lexical/structural evidence: %+v", part)
		}
	}

	badDigest := manifest
	badDigest.Fixtures = append([]corpus.FixtureRecord(nil), manifest.Fixtures...)
	badDigest.Fixtures[0].SHA256 = strings.Repeat("0", 64)
	if _, err := officecompat.QualifyCorpusManifest(root, badDigest); err == nil {
		t.Fatal("manifest package digest drift was accepted")
	}

	escaped := manifest
	escaped.Fixtures = append([]corpus.FixtureRecord(nil), manifest.Fixtures...)
	escaped.Fixtures[0].Package = "../escape.docx"
	if _, err := officecompat.QualifyCorpusManifest(root, escaped); err == nil {
		t.Fatal("manifest path traversal was accepted")
	}

	symlinked := manifest
	symlinked.Fixtures = append([]corpus.FixtureRecord(nil), manifest.Fixtures...)
	symlinkPath := filepath.Join(root, "generated", "packages", "symlink.docx")
	if err := os.Symlink(filepath.Join(root, filepath.FromSlash(packageRelative)), symlinkPath); err != nil {
		t.Skipf("symbolic links unavailable: %v", err)
	}
	symlinked.Fixtures[0].Package = "generated/packages/symlink.docx"
	if _, err := officecompat.QualifyCorpusManifest(root, symlinked); err == nil {
		t.Fatal("symbolic-link corpus package was accepted")
	}
}

func TestStructuralQualifyCorpusManifestEnforcesAggregateBudgetsAndUniquePaths(t *testing.T) {
	root := t.TempDir()
	first := writeStructuralCorpusFixture(t, root, "docx-a", "a")
	second := writeStructuralCorpusFixture(t, root, "docx-b", "b")
	manifest := corpus.Manifest{Protocol: corpus.ManifestProtocol, Generator: corpus.GeneratorVersion, Fixtures: []corpus.FixtureRecord{first, second}}

	t.Run("cumulative package bytes", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxPackageBytes = uint64(first.Bytes + 1)
		if _, err := officecompat.QualifyCorpusManifestWithLimits(root, manifest, limits); err == nil || !strings.Contains(err.Error(), "cumulative") {
			t.Fatalf("cumulative package-byte result = %v", err)
		}
	})
	t.Run("cumulative parts", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxParts = first.Parts + 1
		if _, err := officecompat.QualifyCorpusManifestWithLimits(root, manifest, limits); err == nil || !strings.Contains(err.Error(), "cumulative") {
			t.Fatalf("cumulative part result = %v", err)
		}
	})
	t.Run("cumulative expanded bytes", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxExpandedBytes = uint64(first.ExpandedBytes + 1)
		limits.MaxPartBytes = uint64(first.ExpandedBytes)
		limits.MaxXMLBytes = uint64(first.ExpandedBytes)
		if _, err := officecompat.QualifyCorpusManifestWithLimits(root, manifest, limits); err == nil || !strings.Contains(err.Error(), "cumulative") {
			t.Fatalf("cumulative expanded-byte result = %v", err)
		}
	})
	t.Run("cumulative XML tokens", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxXMLTokens = 18
		if _, err := officecompat.QualifyCorpusManifestWithLimits(root, manifest, limits); err == nil || !strings.Contains(err.Error(), string(officecompat.XMLFailureResourceLimit)) {
			t.Fatalf("cumulative XML-token result = %v", err)
		}
	})
	t.Run("cumulative XML attributes", func(t *testing.T) {
		limits := officecompat.DefaultLimits()
		limits.MaxXMLAttributes = 9
		if _, err := officecompat.QualifyCorpusManifestWithLimits(root, manifest, limits); err == nil || !strings.Contains(err.Error(), string(officecompat.XMLFailureResourceLimit)) {
			t.Fatalf("cumulative XML-attribute result = %v", err)
		}
	})
	t.Run("duplicate package path", func(t *testing.T) {
		duplicate := manifest
		duplicate.Fixtures = append([]corpus.FixtureRecord(nil), manifest.Fixtures...)
		duplicate.Fixtures[1].Package = duplicate.Fixtures[0].Package
		if _, err := officecompat.QualifyCorpusManifest(root, duplicate); err == nil || !strings.Contains(err.Error(), "duplicates") {
			t.Fatalf("duplicate package-path result = %v", err)
		}
	})
	t.Run("duplicate expectation path", func(t *testing.T) {
		duplicate := manifest
		duplicate.Fixtures = append([]corpus.FixtureRecord(nil), manifest.Fixtures...)
		duplicate.Fixtures[1].Expected = duplicate.Fixtures[0].Expected
		if _, err := officecompat.QualifyCorpusManifest(root, duplicate); err == nil || !strings.Contains(err.Error(), "duplicates") {
			t.Fatalf("duplicate expectation-path result = %v", err)
		}
	})
	t.Run("cumulative expectation bytes", func(t *testing.T) {
		large := manifest
		large.Fixtures = append([]corpus.FixtureRecord(nil), manifest.Fixtures...)
		packageBudget := first.Bytes + second.Bytes
		expectation := []byte(strings.Repeat("x", packageBudget/2+1))
		for index := range large.Fixtures {
			filename := filepath.Join(root, filepath.FromSlash(large.Fixtures[index].Expected))
			if err := os.WriteFile(filename, expectation, 0o644); err != nil {
				t.Fatal(err)
			}
			large.Fixtures[index].ExpectedSHA256 = sha256String(expectation)
		}
		limits := officecompat.DefaultLimits()
		limits.MaxPackageBytes = uint64(packageBudget)
		if _, err := officecompat.QualifyCorpusManifestWithLimits(root, large, limits); err == nil || !strings.Contains(err.Error(), "expectations") {
			t.Fatalf("cumulative expectation-byte result = %v", err)
		}
	})
}

func writeStructuralCorpusFixture(t *testing.T, root, id, suffix string) corpus.FixtureRecord {
	t.Helper()
	entries := minimalOPCRoots()
	entries["word/document.xml"] = []byte(`<root/>`)
	packageBytes := buildPackage(t, entries, false, true)
	expectedBytes := []byte("{}\n")
	packageRelative := "generated/packages/" + suffix + ".docx"
	expectedRelative := "generated/expected/" + suffix + ".json"
	for relative, data := range map[string][]byte{packageRelative: packageBytes, expectedRelative: expectedBytes} {
		filename := filepath.Join(root, filepath.FromSlash(relative))
		if err := os.MkdirAll(filepath.Dir(filename), 0o755); err != nil {
			t.Fatal(err)
		}
		if err := os.WriteFile(filename, data, 0o644); err != nil {
			t.Fatal(err)
		}
	}
	expanded := 0
	for _, data := range entries {
		expanded += len(data)
	}
	return corpus.FixtureRecord{
		ID: id, Format: "docx", Dialect: "transitional", Outcome: "accepted",
		Package: packageRelative, Expected: expectedRelative,
		SHA256: sha256String(packageBytes), ExpectedSHA256: sha256String(expectedBytes),
		Bytes: len(packageBytes), ExpandedBytes: expanded, Parts: len(entries),
		Coverage: []string{"structural"}, Provenance: corpus.Provenance{Kind: "generated", Generator: "test", Source: "in-memory test", License: "CC0-1.0"},
	}
}

func sha256String(data []byte) string {
	digest := sha256.Sum256(data)
	return hex.EncodeToString(digest[:])
}
