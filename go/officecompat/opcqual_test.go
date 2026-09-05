package officecompat_test

import (
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

const qualifiedContentTypes = `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`

const qualifiedRootRelationships = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rOffice" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`

const qualifiedDocumentRelationships = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/evidence" TargetMode="External"/><Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.bin"/></Relationships>`

func TestStructuralQualifyOPCProducesDeterministicContentTypeAndRelationshipEvidence(t *testing.T) {
	entries := qualifiedOPCEntries()
	data := buildPackage(t, entries, true, true)
	first, err := officecompat.QualifyOPC(data)
	if err != nil {
		t.Fatal(err)
	}
	second, err := officecompat.QualifyOPC(data)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(first, second) {
		t.Fatalf("OPC qualification is nondeterministic:\nfirst:  %+v\nsecond: %+v", first, second)
	}
	if len(first.Parts) != 4 || len(first.Relationships) != 3 {
		t.Fatalf("qualification evidence is incomplete: %+v", first)
	}
	if got := first.Relationships[0]; got.RelationshipsPart != "_rels/.rels" || got.SourcePart != "" || got.ResolvedPart != "word/document.xml" {
		t.Fatalf("root relationship did not resolve to the exact package spelling: %+v", got)
	}
	if got := first.Relationships[1]; got.ID != "rExternal" || got.TargetMode != "External" || got.ResolvedPart != "" {
		t.Fatalf("external relationship was not retained without package resolution: %+v", got)
	}
	if got := first.Relationships[2]; got.ID != "rImage" || got.SourcePart != "word/document.xml" || got.ResolvedPart != "word/media/pixel.bin" {
		t.Fatalf("part-relative relationship did not resolve exactly: %+v", got)
	}
}

func TestStructuralQualifyOPCRejectsAmbiguousOrDanglingPackageSemantics(t *testing.T) {
	tests := []struct {
		name     string
		mutate   func(map[string][]byte)
		contains string
	}{
		{
			name: "dangling internal target",
			mutate: func(entries map[string][]byte) {
				entries["word/_rels/document.xml.rels"] = []byte(strings.Replace(qualifiedDocumentRelationships, "media/pixel.bin", "media/missing.bin", 1))
			},
			contains: "dangling internal target",
		},
		{
			name: "duplicate relationship id",
			mutate: func(entries map[string][]byte) {
				entries["word/_rels/document.xml.rels"] = []byte(strings.Replace(qualifiedDocumentRelationships, "rImage", "rExternal", 1))
			},
			contains: "duplicate Id",
		},
		{
			name: "orphan relationship source",
			mutate: func(entries map[string][]byte) {
				entries["word/_rels/missing.xml.rels"] = []byte(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
			},
			contains: "absent source part",
		},
		{
			name: "relative external target",
			mutate: func(entries map[string][]byte) {
				entries["word/_rels/document.xml.rels"] = []byte(strings.Replace(qualifiedDocumentRelationships, "https://example.invalid/evidence", "relative/location", 1))
			},
			contains: "non-absolute target",
		},
		{
			name: "undeclared content type",
			mutate: func(entries map[string][]byte) {
				entries["word/media/pixel.dat"] = entries["word/media/pixel.bin"]
				delete(entries, "word/media/pixel.bin")
				entries["word/_rels/document.xml.rels"] = []byte(strings.Replace(qualifiedDocumentRelationships, "pixel.bin", "pixel.dat", 1))
			},
			contains: "no effective content type",
		},
		{
			name: "override for absent part",
			mutate: func(entries map[string][]byte) {
				entries["[Content_Types].xml"] = []byte(strings.Replace(qualifiedContentTypes, "</Types>", `<Override PartName="/word/absent.xml" ContentType="application/xml"/></Types>`, 1))
			},
			contains: "targets absent part",
		},
		{
			name: "encoded traversal",
			mutate: func(entries map[string][]byte) {
				entries["word/_rels/document.xml.rels"] = []byte(strings.Replace(qualifiedDocumentRelationships, "media/pixel.bin", "%2E%2E/pixel.bin", 1))
			},
			contains: "encoded traversal",
		},
	}

	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			entries := qualifiedOPCEntries()
			test.mutate(entries)
			_, err := officecompat.QualifyOPC(buildPackage(t, entries, false, true))
			if err == nil || !strings.Contains(err.Error(), test.contains) {
				t.Fatalf("qualification result = %v, want marker %q", err, test.contains)
			}
		})
	}
}

func TestStructuralQualifyCheckedInAcceptedCorpusOPCGraphs(t *testing.T) {
	manifestData, err := os.ReadFile(filepath.Join("corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var manifest corpus.Manifest
	if err := json.Unmarshal(manifestData, &manifest); err != nil {
		t.Fatal(err)
	}
	qualified := map[string]int{"docx": 0, "pptx": 0, "xlsx": 0}
	for _, fixture := range manifest.Fixtures {
		if fixture.Outcome != "accepted" {
			continue
		}
		packageData, err := os.ReadFile(filepath.Join("corpus", filepath.FromSlash(fixture.Package)))
		if err != nil {
			t.Fatal(err)
		}
		evidence, err := officecompat.QualifyOPC(packageData)
		if err != nil {
			t.Fatalf("accepted fixture %q has an invalid OPC graph: %v", fixture.ID, err)
		}
		if len(evidence.Parts) != fixture.Parts-1 || len(evidence.Relationships) == 0 {
			t.Fatalf("accepted fixture %q lacks complete OPC evidence: %+v", fixture.ID, evidence)
		}
		qualified[fixture.Format]++
	}
	for format, count := range qualified {
		if count == 0 {
			t.Errorf("accepted corpus contains no qualified %s graph", format)
		}
	}
}

func TestStructuralQualifyOPCEnforcesCumulativeXMLBudget(t *testing.T) {
	data := buildPackage(t, qualifiedOPCEntries(), false, true)
	limits := officecompat.DefaultLimits()
	limits.MaxXMLTokens = 10
	if _, err := officecompat.QualifyOPCWithLimits(data, limits); err == nil || !strings.Contains(err.Error(), "token budget") {
		t.Fatalf("cumulative OPC XML token result = %v", err)
	}
}

func qualifiedOPCEntries() map[string][]byte {
	return map[string][]byte{
		"[Content_Types].xml":          []byte(qualifiedContentTypes),
		"_rels/.rels":                  []byte(qualifiedRootRelationships),
		"word/_rels/document.xml.rels": []byte(qualifiedDocumentRelationships),
		"word/document.xml":            []byte(`<w:document xmlns:w="urn:test"><w:body/></w:document>`),
		"word/media/pixel.bin":         {0x00, 0x01, 0x02},
	}
}
