// Package corpus generates the bounded, deterministic native Office fixture
// corpus. Declarative specs are the authority; checked-in OOXML packages,
// expected native JSON, and the manifest are reproducible outputs.
package corpus

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/base64"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"os"
	"path/filepath"
	"regexp"
	"sort"
	"strings"
	"time"
)

const (
	SpecProtocol        = "injoffice-office-fixture-spec/v1"
	ExpectationProtocol = "injoffice-expected-native-json/v1"
	ManifestProtocol    = "injoffice-office-compat-corpus/v1"
	GeneratorVersion    = "officecompat-corpusgen/v1"

	maxSpecs              = 128
	maxPartsPerFixture    = 256
	maxPartExpandedBytes  = 4 << 20
	maxTotalExpandedBytes = 12 << 20
	maxExpectedBytes      = 4 << 20
)

var fixtureIDPattern = regexp.MustCompile(`^[a-z0-9]+(?:-[a-z0-9]+)*$`)

type Provenance struct {
	Kind      string `json:"kind"`
	Generator string `json:"generator"`
	Source    string `json:"source"`
	License   string `json:"license"`
}

type RepeatPayload struct {
	Text  string `json:"text"`
	Count int    `json:"count"`
}

type PartSpec struct {
	Name   string         `json:"name"`
	Method string         `json:"method,omitempty"`
	Text   *string        `json:"text,omitempty"`
	Base64 *string        `json:"base64,omitempty"`
	Repeat *RepeatPayload `json:"repeat,omitempty"`
}

type Refusal struct {
	Class    string `json:"class"`
	Contains string `json:"contains"`
}

// Expectation is the common sidecar protocol. Native remains the exact JSON
// emitted by the format's native contract; the shared envelope does not
// normalize, reconstruct, or otherwise reinterpret it.
type Expectation struct {
	Protocol  string          `json:"protocol"`
	FixtureID string          `json:"fixtureId"`
	Format    string          `json:"format"`
	Dialect   string          `json:"dialect"`
	Outcome   string          `json:"outcome"`
	Native    json.RawMessage `json:"native,omitempty"`
	Refusal   *Refusal        `json:"refusal,omitempty"`
}

type Spec struct {
	Protocol    string      `json:"protocol"`
	ID          string      `json:"id"`
	Format      string      `json:"format"`
	Dialect     string      `json:"dialect"`
	Outcome     string      `json:"outcome"`
	Provenance  Provenance  `json:"provenance"`
	Coverage    []string    `json:"coverage"`
	Parts       []PartSpec  `json:"parts"`
	Expectation Expectation `json:"expectation"`
}

type FixtureRecord struct {
	ID             string     `json:"id"`
	Format         string     `json:"format"`
	Dialect        string     `json:"dialect"`
	Outcome        string     `json:"outcome"`
	Package        string     `json:"package"`
	Expected       string     `json:"expected"`
	SHA256         string     `json:"sha256"`
	ExpectedSHA256 string     `json:"expectedSha256"`
	Bytes          int        `json:"bytes"`
	ExpandedBytes  int        `json:"expandedBytes"`
	Parts          int        `json:"parts"`
	Coverage       []string   `json:"coverage"`
	Provenance     Provenance `json:"provenance"`
}

type Manifest struct {
	Protocol  string          `json:"protocol"`
	Generator string          `json:"generator"`
	Fixtures  []FixtureRecord `json:"fixtures"`
}

type generatedFile struct {
	path string
	data []byte
}

// Generate writes every derived package, expectation sidecar, and the
// canonical manifest beneath root.
func Generate(root string) error {
	files, err := render(root)
	if err != nil {
		return err
	}
	for _, file := range files {
		if err := os.MkdirAll(filepath.Dir(file.path), 0o755); err != nil {
			return fmt.Errorf("corpus: create output directory: %w", err)
		}
		if err := os.WriteFile(file.path, file.data, 0o644); err != nil {
			return fmt.Errorf("corpus: write %s: %w", file.path, err)
		}
	}
	return removeStaleOutputs(root, files)
}

// Check fails when generated outputs are absent, stale, or contain files no
// longer described by a spec.
func Check(root string) error {
	files, err := render(root)
	if err != nil {
		return err
	}
	wanted := make(map[string]bool, len(files))
	for _, file := range files {
		wanted[filepath.Clean(file.path)] = true
		actual, readErr := os.ReadFile(file.path)
		if readErr != nil {
			return fmt.Errorf("corpus: generated file %s: %w", file.path, readErr)
		}
		if !bytes.Equal(actual, file.data) {
			return fmt.Errorf("corpus: generated file %s is stale", file.path)
		}
	}
	return walkGenerated(root, func(path string) error {
		if !wanted[filepath.Clean(path)] {
			return fmt.Errorf("corpus: stale generated file %s", path)
		}
		return nil
	})
}

func render(root string) ([]generatedFile, error) {
	specs, err := loadSpecs(filepath.Join(root, "specs"))
	if err != nil {
		return nil, err
	}
	manifest := Manifest{Protocol: ManifestProtocol, Generator: GeneratorVersion, Fixtures: make([]FixtureRecord, 0, len(specs))}
	files := make([]generatedFile, 0, len(specs)*2+1)
	for _, spec := range specs {
		packageBytes, expanded, err := buildPackage(spec)
		if err != nil {
			return nil, fmt.Errorf("corpus: fixture %s: %w", spec.ID, err)
		}
		expectedBytes, err := marshalCanonical(spec.Expectation)
		if err != nil {
			return nil, fmt.Errorf("corpus: fixture %s expectation: %w", spec.ID, err)
		}
		if len(expectedBytes) > maxExpectedBytes {
			return nil, fmt.Errorf("corpus: fixture %s expectation exceeds %d bytes", spec.ID, maxExpectedBytes)
		}

		extension := "." + spec.Format
		packageRel := filepath.ToSlash(filepath.Join("generated", "packages", spec.ID+extension))
		expectedRel := filepath.ToSlash(filepath.Join("generated", "expected", spec.ID+".json"))
		files = append(files,
			generatedFile{path: filepath.Join(root, filepath.FromSlash(packageRel)), data: packageBytes},
			generatedFile{path: filepath.Join(root, filepath.FromSlash(expectedRel)), data: expectedBytes},
		)
		manifest.Fixtures = append(manifest.Fixtures, FixtureRecord{
			ID: spec.ID, Format: spec.Format, Dialect: spec.Dialect, Outcome: spec.Outcome,
			Package: packageRel, Expected: expectedRel,
			SHA256: sha256Hex(packageBytes), ExpectedSHA256: sha256Hex(expectedBytes),
			Bytes: len(packageBytes), ExpandedBytes: expanded, Parts: len(spec.Parts),
			Coverage: append([]string(nil), spec.Coverage...), Provenance: spec.Provenance,
		})
	}
	manifestBytes, err := marshalCanonical(manifest)
	if err != nil {
		return nil, err
	}
	files = append(files, generatedFile{path: filepath.Join(root, "manifest.json"), data: manifestBytes})
	sort.Slice(files, func(i, j int) bool { return files[i].path < files[j].path })
	return files, nil
}

func loadSpecs(dir string) ([]Spec, error) {
	var paths []string
	err := filepath.WalkDir(dir, func(path string, entry fs.DirEntry, walkErr error) error {
		if walkErr != nil {
			if errors.Is(walkErr, os.ErrNotExist) && path == dir {
				return nil
			}
			return walkErr
		}
		if !entry.IsDir() && strings.HasSuffix(entry.Name(), ".json") {
			paths = append(paths, path)
		}
		return nil
	})
	if err != nil && !errors.Is(err, os.ErrNotExist) {
		return nil, fmt.Errorf("corpus: discover specs: %w", err)
	}
	if len(paths) > maxSpecs {
		return nil, fmt.Errorf("corpus: spec count exceeds %d", maxSpecs)
	}
	sort.Strings(paths)
	specs := make([]Spec, 0, len(paths))
	seen := map[string]bool{}
	for _, path := range paths {
		data, err := os.ReadFile(path)
		if err != nil {
			return nil, fmt.Errorf("corpus: read spec %s: %w", path, err)
		}
		var spec Spec
		decoder := json.NewDecoder(bytes.NewReader(data))
		decoder.DisallowUnknownFields()
		if err := decoder.Decode(&spec); err != nil {
			return nil, fmt.Errorf("corpus: decode spec %s: %w", path, err)
		}
		if err := decoder.Decode(&struct{}{}); !errors.Is(err, io.EOF) {
			return nil, fmt.Errorf("corpus: spec %s has trailing JSON", path)
		}
		if err := validateSpec(spec); err != nil {
			return nil, fmt.Errorf("corpus: spec %s: %w", path, err)
		}
		if seen[spec.ID] {
			return nil, fmt.Errorf("corpus: duplicate fixture id %q", spec.ID)
		}
		seen[spec.ID] = true
		specs = append(specs, spec)
	}
	sort.Slice(specs, func(i, j int) bool { return specs[i].ID < specs[j].ID })
	return specs, nil
}

func validateSpec(spec Spec) error {
	if spec.Protocol != SpecProtocol {
		return fmt.Errorf("protocol must be %q", SpecProtocol)
	}
	if !fixtureIDPattern.MatchString(spec.ID) {
		return fmt.Errorf("invalid fixture id %q", spec.ID)
	}
	if spec.Format != "docx" && spec.Format != "pptx" && spec.Format != "xlsx" {
		return fmt.Errorf("invalid format %q", spec.Format)
	}
	if spec.Dialect != "strict" && spec.Dialect != "transitional" {
		return fmt.Errorf("invalid dialect %q", spec.Dialect)
	}
	if spec.Outcome != "accepted" && spec.Outcome != "refused" {
		return fmt.Errorf("invalid outcome %q", spec.Outcome)
	}
	if spec.Provenance.Kind != "generated" || spec.Provenance.Generator == "" || spec.Provenance.Source == "" || spec.Provenance.License == "" {
		return errors.New("generated provenance requires generator, source, and license")
	}
	if len(spec.Coverage) == 0 || !sortedUniqueNonempty(spec.Coverage) {
		return errors.New("coverage must be a non-empty sorted unique list")
	}
	if len(spec.Parts) == 0 || len(spec.Parts) > maxPartsPerFixture {
		return fmt.Errorf("part count must be 1..%d", maxPartsPerFixture)
	}
	if err := validateExpectation(spec); err != nil {
		return err
	}
	return nil
}

func validateExpectation(spec Spec) error {
	expectation := spec.Expectation
	if expectation.Protocol != ExpectationProtocol || expectation.FixtureID != spec.ID || expectation.Format != spec.Format || expectation.Dialect != spec.Dialect || expectation.Outcome != spec.Outcome {
		return errors.New("expectation envelope does not match fixture identity")
	}
	if spec.Outcome == "accepted" {
		trimmed := bytes.TrimSpace(expectation.Native)
		if len(trimmed) == 0 || bytes.Equal(trimmed, []byte("null")) || trimmed[0] != '{' || expectation.Refusal != nil {
			return errors.New("accepted fixture requires native object and no refusal")
		}
		return nil
	}
	if len(bytes.TrimSpace(expectation.Native)) != 0 || expectation.Refusal == nil || expectation.Refusal.Class == "" || expectation.Refusal.Contains == "" {
		return errors.New("refused fixture requires class/contains and no native object")
	}
	return nil
}

func buildPackage(spec Spec) ([]byte, int, error) {
	type materializedPart struct {
		name   string
		method uint16
		data   []byte
		order  int
	}
	parts := make([]materializedPart, 0, len(spec.Parts))
	total := 0
	for index, part := range spec.Parts {
		if err := validatePartName(part.Name); err != nil {
			return nil, 0, err
		}
		data, err := materialize(part)
		if err != nil {
			return nil, 0, fmt.Errorf("part %q: %w", part.Name, err)
		}
		if len(data) > maxPartExpandedBytes || total > maxTotalExpandedBytes-len(data) {
			return nil, 0, errors.New("expanded fixture exceeds corpus bounds")
		}
		total += len(data)
		method := uint16(zip.Deflate)
		if part.Method == "store" {
			method = zip.Store
		} else if part.Method != "" && part.Method != "deflate" {
			return nil, 0, fmt.Errorf("part %q has invalid method %q", part.Name, part.Method)
		}
		parts = append(parts, materializedPart{name: part.Name, method: method, data: data, order: index})
	}
	sort.SliceStable(parts, func(i, j int) bool {
		if parts[i].name == parts[j].name {
			return parts[i].order < parts[j].order
		}
		return parts[i].name < parts[j].name
	})
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	fixedTime := time.Date(1980, time.January, 1, 0, 0, 0, 0, time.UTC)
	for _, part := range parts {
		header := &zip.FileHeader{Name: part.name, Method: part.method}
		header.SetModTime(fixedTime)
		header.SetMode(0o644)
		entry, err := writer.CreateHeader(header)
		if err != nil {
			return nil, 0, err
		}
		if _, err := entry.Write(part.data); err != nil {
			return nil, 0, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, 0, err
	}
	return output.Bytes(), total, nil
}

func materialize(part PartSpec) ([]byte, error) {
	choices := 0
	if part.Text != nil {
		choices++
	}
	if part.Base64 != nil {
		choices++
	}
	if part.Repeat != nil {
		choices++
	}
	if choices != 1 {
		return nil, errors.New("exactly one of text, base64, or repeat is required")
	}
	if part.Text != nil {
		return []byte(*part.Text), nil
	}
	if part.Base64 != nil {
		decoded, err := base64.StdEncoding.DecodeString(*part.Base64)
		if err != nil {
			return nil, fmt.Errorf("invalid base64: %w", err)
		}
		return decoded, nil
	}
	if part.Repeat.Count < 1 || len(part.Repeat.Text) == 0 || len(part.Repeat.Text) > maxPartExpandedBytes/part.Repeat.Count {
		return nil, errors.New("repeat payload exceeds corpus bounds")
	}
	return []byte(strings.Repeat(part.Repeat.Text, part.Repeat.Count)), nil
}

func validatePartName(name string) error {
	if name == "" || strings.HasPrefix(name, "/") || strings.Contains(name, `\`) || strings.ContainsRune(name, '\x00') {
		return fmt.Errorf("unsafe OPC part name %q", name)
	}
	clean := filepath.ToSlash(filepath.Clean(name))
	if clean != name || clean == "." || clean == ".." || strings.HasPrefix(clean, "../") {
		return fmt.Errorf("unsafe OPC part name %q", name)
	}
	return nil
}

func marshalCanonical(value any) ([]byte, error) {
	var output bytes.Buffer
	encoder := json.NewEncoder(&output)
	encoder.SetEscapeHTML(false)
	encoder.SetIndent("", "  ")
	if err := encoder.Encode(value); err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}

func sha256Hex(data []byte) string {
	digest := sha256.Sum256(data)
	return hex.EncodeToString(digest[:])
}

func sortedUniqueNonempty(values []string) bool {
	for index, value := range values {
		if value == "" || (index > 0 && values[index-1] >= value) {
			return false
		}
	}
	return true
}

func removeStaleOutputs(root string, files []generatedFile) error {
	wanted := make(map[string]bool, len(files))
	for _, file := range files {
		wanted[filepath.Clean(file.path)] = true
	}
	return walkGenerated(root, func(path string) error {
		if wanted[filepath.Clean(path)] {
			return nil
		}
		if err := os.Remove(path); err != nil {
			return fmt.Errorf("corpus: remove stale generated file %s: %w", path, err)
		}
		return nil
	})
}

func walkGenerated(root string, visit func(string) error) error {
	for _, relative := range []string{filepath.Join("generated", "packages"), filepath.Join("generated", "expected")} {
		dir := filepath.Join(root, relative)
		err := filepath.WalkDir(dir, func(path string, entry fs.DirEntry, walkErr error) error {
			if walkErr != nil {
				return walkErr
			}
			if !entry.IsDir() {
				return visit(path)
			}
			return nil
		})
		if err != nil && !errors.Is(err, os.ErrNotExist) {
			return err
		}
	}
	return nil
}
