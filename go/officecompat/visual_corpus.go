package officecompat

import (
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"regexp"
	"strings"

	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

var visualFixtureIDPattern = regexp.MustCompile(`^[a-z0-9]+(?:-[a-z0-9]+)*$`)

// VisualFixtureReport binds visual evidence to the exact native corpus record
// whose package was rendered. It preserves the corpus protocol rather than
// introducing a parallel fixture identity or manifest.
type VisualFixtureReport struct {
	Fixture corpus.FixtureRecord `json:"fixture"`
	Visual  VisualReport         `json:"visual"`
}

// Matches reports whether the fixture's visual comparison satisfies its
// recorded tolerance.
func (report VisualFixtureReport) Matches() bool {
	return report.Visual.Matches()
}

// VisualCorpusFixtures validates the corpus identity and returns independent
// copies of accepted fixture records in canonical manifest order. Refused
// fixtures remain qualification inputs for native refusal tests but cannot
// produce visual reference images and are therefore omitted.
func VisualCorpusFixtures(manifest corpus.Manifest) ([]corpus.FixtureRecord, error) {
	if manifest.Protocol != corpus.ManifestProtocol {
		return nil, fmt.Errorf("officecompat: visual corpus protocol must be %q", corpus.ManifestProtocol)
	}
	if manifest.Generator != corpus.GeneratorVersion {
		return nil, fmt.Errorf("officecompat: visual corpus generator must be %q", corpus.GeneratorVersion)
	}
	if manifest.Fixtures == nil {
		return nil, fmt.Errorf("officecompat: visual corpus fixtures must be present")
	}

	accepted := make([]corpus.FixtureRecord, 0, len(manifest.Fixtures))
	previousID := ""
	for index, fixture := range manifest.Fixtures {
		if err := validateVisualFixtureRecord(fixture, false); err != nil {
			return nil, fmt.Errorf("officecompat: visual corpus fixture %d: %w", index, err)
		}
		if fixture.ID <= previousID {
			return nil, fmt.Errorf("officecompat: visual corpus fixtures are not strictly ordered by id at %q", fixture.ID)
		}
		previousID = fixture.ID
		if fixture.Outcome == "accepted" {
			accepted = append(accepted, cloneFixtureRecord(fixture))
		}
	}
	return accepted, nil
}

// CompareFixturePNG compares externally supplied renders and binds the stable
// report to an accepted corpus FixtureRecord. Rendering and artifact lookup
// remain the responsibility of dedicated qualification providers/runners.
func CompareFixturePNG(fixture corpus.FixtureRecord, reference, candidate []byte, limits VisualLimits, tolerance VisualTolerance) (VisualFixtureReport, error) {
	if err := validateVisualFixtureRecord(fixture, true); err != nil {
		return VisualFixtureReport{}, fmt.Errorf("officecompat: visual fixture: %w", err)
	}
	report, err := ComparePNG(reference, candidate, limits, tolerance)
	if err != nil {
		return VisualFixtureReport{}, err
	}
	return VisualFixtureReport{Fixture: cloneFixtureRecord(fixture), Visual: report}, nil
}

func validateVisualFixtureRecord(fixture corpus.FixtureRecord, requireAccepted bool) error {
	if !visualFixtureIDPattern.MatchString(fixture.ID) {
		return fmt.Errorf("fixture id %q is invalid", fixture.ID)
	}
	if fixture.Format != "docx" && fixture.Format != "pptx" && fixture.Format != "xlsx" {
		return fmt.Errorf("fixture %q has unsupported format %q", fixture.ID, fixture.Format)
	}
	if fixture.Dialect != "strict" && fixture.Dialect != "transitional" {
		return fmt.Errorf("fixture %q has unsupported dialect %q", fixture.ID, fixture.Dialect)
	}
	if fixture.Outcome != "accepted" && fixture.Outcome != "refused" {
		return fmt.Errorf("fixture %q has unsupported outcome %q", fixture.ID, fixture.Outcome)
	}
	if requireAccepted && fixture.Outcome != "accepted" {
		return fmt.Errorf("fixture %q is refused and has no visual render", fixture.ID)
	}
	wantPackage := "generated/packages/" + fixture.ID + "." + fixture.Format
	if fixture.Package != wantPackage {
		return fmt.Errorf("fixture %q has invalid package path %q", fixture.ID, fixture.Package)
	}
	wantExpected := "generated/expected/" + fixture.ID + ".json"
	if fixture.Expected != wantExpected {
		return fmt.Errorf("fixture %q has invalid expectation path %q", fixture.ID, fixture.Expected)
	}
	if !canonicalSHA256(fixture.SHA256) || !canonicalSHA256(fixture.ExpectedSHA256) {
		return fmt.Errorf("fixture %q has invalid SHA-256", fixture.ID)
	}
	if fixture.Bytes <= 0 || fixture.ExpandedBytes <= 0 || fixture.Parts <= 0 {
		return fmt.Errorf("fixture %q has invalid resource inventory", fixture.ID)
	}
	if !sortedUniqueVisualMetadata(fixture.Coverage) {
		return fmt.Errorf("fixture %q coverage must be a non-empty sorted unique list", fixture.ID)
	}
	if fixture.Provenance.Kind != "generated" || fixture.Provenance.Generator == "" || fixture.Provenance.Source == "" || fixture.Provenance.License == "" {
		return fmt.Errorf("fixture %q has invalid generated provenance", fixture.ID)
	}
	return nil
}

func canonicalSHA256(value string) bool {
	if len(value) != sha256.Size*2 || strings.ToLower(value) != value {
		return false
	}
	_, err := hex.DecodeString(value)
	return err == nil
}

func cloneFixtureRecord(fixture corpus.FixtureRecord) corpus.FixtureRecord {
	fixture.Coverage = append([]string(nil), fixture.Coverage...)
	return fixture
}

func sortedUniqueVisualMetadata(values []string) bool {
	if len(values) == 0 {
		return false
	}
	for index, value := range values {
		if value == "" || (index > 0 && values[index-1] >= value) {
			return false
		}
	}
	return true
}
