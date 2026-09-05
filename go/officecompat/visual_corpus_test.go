package officecompat_test

import (
	"image/color"
	"image/png"
	"reflect"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

func TestVisualCorpusFixturesConsumesCanonicalManifest(t *testing.T) {
	accepted := visualFixtureRecord("docx-transitional-text", "docx", "transitional", "accepted")
	accepted.Coverage = []string{"common-text", "transitional"}
	refused := visualFixtureRecord("xlsx-strict-refusal", "xlsx", "strict", "refused")
	manifest := corpus.Manifest{
		Protocol:  corpus.ManifestProtocol,
		Generator: corpus.GeneratorVersion,
		Fixtures:  []corpus.FixtureRecord{accepted, refused},
	}

	fixtures, err := officecompat.VisualCorpusFixtures(manifest)
	if err != nil {
		t.Fatal(err)
	}
	if len(fixtures) != 1 || !reflect.DeepEqual(fixtures[0], accepted) {
		t.Fatalf("unexpected visual fixture selection: %+v", fixtures)
	}
	fixtures[0].Coverage[0] = "mutated"
	if manifest.Fixtures[0].Coverage[0] != "common-text" {
		t.Fatal("visual fixture selection must not alias manifest coverage")
	}

	empty, err := officecompat.VisualCorpusFixtures(corpus.Manifest{
		Protocol: corpus.ManifestProtocol, Generator: corpus.GeneratorVersion, Fixtures: []corpus.FixtureRecord{},
	})
	if err != nil || empty == nil || len(empty) != 0 {
		t.Fatalf("canonical empty manifest should remain a non-nil empty selection: %#v, %v", empty, err)
	}
}

func TestVisualCorpusFixturesRejectsManifestDriftAndUnsafeRecords(t *testing.T) {
	first := visualFixtureRecord("docx-transitional-text", "docx", "transitional", "accepted")
	second := visualFixtureRecord("pptx-strict-shape", "pptx", "strict", "accepted")
	valid := corpus.Manifest{Protocol: corpus.ManifestProtocol, Generator: corpus.GeneratorVersion, Fixtures: []corpus.FixtureRecord{first, second}}

	tests := []struct {
		name      string
		mutate    func(*corpus.Manifest)
		wantError string
	}{
		{name: "protocol", mutate: func(value *corpus.Manifest) { value.Protocol = "v2" }, wantError: "visual corpus protocol"},
		{name: "generator", mutate: func(value *corpus.Manifest) { value.Generator = "other" }, wantError: "visual corpus generator"},
		{name: "missing fixtures", mutate: func(value *corpus.Manifest) { value.Fixtures = nil }, wantError: "fixtures must be present"},
		{name: "unordered", mutate: func(value *corpus.Manifest) {
			value.Fixtures[0], value.Fixtures[1] = value.Fixtures[1], value.Fixtures[0]
		}, wantError: "not strictly ordered"},
		{name: "duplicate", mutate: func(value *corpus.Manifest) { value.Fixtures[1] = value.Fixtures[0] }, wantError: "not strictly ordered"},
		{name: "unsafe id", mutate: func(value *corpus.Manifest) { value.Fixtures[0].ID = "../escape" }, wantError: "fixture id"},
		{name: "format", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Format = "html" }, wantError: "unsupported format"},
		{name: "dialect", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Dialect = "approximate" }, wantError: "unsupported dialect"},
		{name: "outcome", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Outcome = "rendered" }, wantError: "unsupported outcome"},
		{name: "package traversal", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Package = "../escape.docx" }, wantError: "invalid package path"},
		{name: "expectation mismatch", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Expected = "generated/expected/other.json" }, wantError: "invalid expectation path"},
		{name: "package hash", mutate: func(value *corpus.Manifest) { value.Fixtures[0].SHA256 = strings.Repeat("A", 64) }, wantError: "invalid SHA-256"},
		{name: "expected hash", mutate: func(value *corpus.Manifest) { value.Fixtures[0].ExpectedSHA256 = "00" }, wantError: "invalid SHA-256"},
		{name: "resource inventory", mutate: func(value *corpus.Manifest) { value.Fixtures[0].ExpandedBytes = 0 }, wantError: "invalid resource inventory"},
		{name: "missing coverage", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Coverage = nil }, wantError: "non-empty sorted unique list"},
		{name: "empty coverage", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Coverage = []string{""} }, wantError: "non-empty sorted unique list"},
		{name: "unsorted coverage", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Coverage = []string{"z", "a"} }, wantError: "non-empty sorted unique list"},
		{name: "duplicate coverage", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Coverage = []string{"native", "native"} }, wantError: "non-empty sorted unique list"},
		{name: "provenance kind", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Provenance.Kind = "downloaded" }, wantError: "invalid generated provenance"},
		{name: "provenance generator", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Provenance.Generator = "" }, wantError: "invalid generated provenance"},
		{name: "provenance source", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Provenance.Source = "" }, wantError: "invalid generated provenance"},
		{name: "provenance license", mutate: func(value *corpus.Manifest) { value.Fixtures[0].Provenance.License = "" }, wantError: "invalid generated provenance"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			manifest := cloneVisualManifest(valid)
			test.mutate(&manifest)
			_, err := officecompat.VisualCorpusFixtures(manifest)
			if err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("expected %q, got %v", test.wantError, err)
			}
		})
	}
}

func TestVisualCompareFixturePNGBindsExactCorpusIdentity(t *testing.T) {
	fixture := visualFixtureRecord("pptx-strict-shape", "pptx", "strict", "accepted")
	fixture.Coverage = []string{"shape", "strict"}
	pixels := solidPixels(1, color.NRGBA{A: 255})
	reference := encodePNG(t, 1, 1, pixels, png.DefaultCompression)
	candidate := encodePNG(t, 1, 1, pixels, png.DefaultCompression)

	report, err := officecompat.CompareFixturePNG(fixture, reference, candidate, visualTestLimits, officecompat.VisualTolerance{})
	if err != nil {
		t.Fatal(err)
	}
	if !report.Matches() || !reflect.DeepEqual(report.Fixture, fixture) || report.Visual.ReferencePixelSHA256 == "" {
		t.Fatalf("visual evidence lost corpus identity: %+v", report)
	}
	report.Fixture.Coverage[0] = "mutated"
	if fixture.Coverage[0] != "shape" {
		t.Fatal("fixture report must not alias caller coverage")
	}
}

func TestVisualCompareFixturePNGRejectsRefusedCorpusRecord(t *testing.T) {
	fixture := visualFixtureRecord("xlsx-strict-refusal", "xlsx", "strict", "refused")
	_, err := officecompat.CompareFixturePNG(fixture, nil, nil, visualTestLimits, officecompat.VisualTolerance{})
	if err == nil || !strings.Contains(err.Error(), "is refused and has no visual render") {
		t.Fatalf("unexpected error: %v", err)
	}
}

func visualFixtureRecord(id, format, dialect, outcome string) corpus.FixtureRecord {
	return corpus.FixtureRecord{
		ID:             id,
		Format:         format,
		Dialect:        dialect,
		Outcome:        outcome,
		Package:        "generated/packages/" + id + "." + format,
		Expected:       "generated/expected/" + id + ".json",
		SHA256:         strings.Repeat("1", 64),
		ExpectedSHA256: strings.Repeat("2", 64),
		Bytes:          100,
		ExpandedBytes:  200,
		Parts:          3,
		Coverage:       []string{"native"},
		Provenance: corpus.Provenance{
			Kind: "generated", Generator: "test", Source: "unit test", License: "CC0-1.0",
		},
	}
}

func cloneVisualManifest(value corpus.Manifest) corpus.Manifest {
	cloned := value
	cloned.Fixtures = append([]corpus.FixtureRecord(nil), value.Fixtures...)
	for index := range cloned.Fixtures {
		cloned.Fixtures[index].Coverage = append([]string(nil), cloned.Fixtures[index].Coverage...)
	}
	return cloned
}
