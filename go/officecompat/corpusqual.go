package officecompat

import (
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"io"
	"os"
	"path"
	"path/filepath"
	"regexp"
	"strings"

	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
)

const maxCorpusExpectedBytes = 4 * 1024 * 1024

var corpusSHA256Pattern = regexp.MustCompile(`^[0-9a-f]{64}$`)

// XMLPartQualification is the bounded structural validation of one corpus XML
// part. SHA256 is the authoritative lexical payload digest; StructuralSHA256
// is additional deterministic diagnostic evidence.
type XMLPartQualification struct {
	Name             string `json:"name"`
	SHA256           string `json:"sha256"`
	StructuralSHA256 string `json:"structural_sha256"`
}

// CorpusQualification binds a corpus manifest record to its exact package
// inventory and namespace-aware XML structural fingerprints.
type CorpusQualification struct {
	FixtureID string                 `json:"fixture_id"`
	Format    string                 `json:"format"`
	Dialect   string                 `json:"dialect"`
	Outcome   string                 `json:"outcome"`
	Inventory Inventory              `json:"inventory"`
	XMLParts  []XMLPartQualification `json:"xml_parts"`
}

// QualifyCorpusManifest consumes the corpus.Manifest/FixtureRecord protocol
// directly. It verifies manifest hashes and declared resource metadata before
// inspecting packages, then requires every .xml/.rels part to produce a
// bounded structural fingerprint. Specs and generated fixtures remain owned by
// the corpus package and are never modified.
func QualifyCorpusManifest(root string, manifest corpus.Manifest) ([]CorpusQualification, error) {
	return QualifyCorpusManifestWithLimits(root, manifest, DefaultLimits())
}

// QualifyCorpusManifestWithLimits is QualifyCorpusManifest with an explicit
// package and XML resource envelope.
func QualifyCorpusManifestWithLimits(root string, manifest corpus.Manifest, limits Limits) ([]CorpusQualification, error) {
	if err := limits.validate(); err != nil {
		return nil, err
	}
	if manifest.Protocol != corpus.ManifestProtocol || manifest.Generator != corpus.GeneratorVersion {
		return nil, fmt.Errorf("officecompat: unsupported corpus manifest protocol or generator")
	}
	if len(manifest.Fixtures) > limits.MaxParts {
		return nil, fmt.Errorf("officecompat: corpus fixture count exceeds %d", limits.MaxParts)
	}
	root = filepath.Clean(root)
	fixturePaths, err := preflightCorpusManifest(root, manifest, limits)
	if err != nil {
		return nil, err
	}
	results := make([]CorpusQualification, 0, len(manifest.Fixtures))
	remainingXMLTokens := limits.MaxXMLTokens
	remainingXMLAttributes := limits.MaxXMLAttributes
	for index, fixture := range manifest.Fixtures {
		packageBytes, err := readExactCorpusFile(fixturePaths[index].packagePath, uint64(fixture.Bytes), limits.MaxPackageBytes)
		if err != nil {
			return nil, fmt.Errorf("officecompat: corpus fixture %q package: %w", fixture.ID, err)
		}
		if digestBytes(packageBytes) != fixture.SHA256 {
			return nil, fmt.Errorf("officecompat: corpus fixture %q package digest mismatch", fixture.ID)
		}
		if err := verifyCorpusFileDigest(fixturePaths[index].expectedPath, fixture.ExpectedSHA256, maxCorpusExpectedBytes); err != nil {
			return nil, fmt.Errorf("officecompat: corpus fixture %q expectation: %w", fixture.ID, err)
		}

		pkg, err := inspectPackage(packageBytes, limits)
		if err != nil {
			return nil, fmt.Errorf("officecompat: corpus fixture %q: %w", fixture.ID, err)
		}
		if len(pkg.inventory.Parts) != fixture.Parts {
			return nil, fmt.Errorf("officecompat: corpus fixture %q part count %d does not match manifest %d", fixture.ID, len(pkg.inventory.Parts), fixture.Parts)
		}
		var expanded uint64
		xmlParts := make([]XMLPartQualification, 0)
		for _, part := range pkg.inventory.Parts {
			expanded += part.Size
			if !isXMLPart(part.Name) {
				continue
			}
			partLimits := limits
			partLimits.MaxXMLTokens = remainingXMLTokens
			partLimits.MaxXMLAttributes = remainingXMLAttributes
			fingerprint := fingerprintXML(pkg.files[part.Name].Open, part.Size, partLimits)
			if fingerprint.failure != "" {
				return nil, fmt.Errorf("officecompat: corpus fixture %q XML part %q failed structural validation: %s", fixture.ID, part.Name, fingerprint.failure)
			}
			remainingXMLTokens -= fingerprint.tokens
			remainingXMLAttributes -= fingerprint.attributes
			xmlParts = append(xmlParts, XMLPartQualification{Name: part.Name, SHA256: part.SHA256, StructuralSHA256: fingerprint.digest})
		}
		if expanded != uint64(fixture.ExpandedBytes) {
			return nil, fmt.Errorf("officecompat: corpus fixture %q expanded bytes %d do not match manifest %d", fixture.ID, expanded, fixture.ExpandedBytes)
		}
		results = append(results, CorpusQualification{
			FixtureID: fixture.ID,
			Format:    fixture.Format, Dialect: fixture.Dialect, Outcome: fixture.Outcome,
			Inventory: pkg.inventory, XMLParts: xmlParts,
		})
	}
	return results, nil
}

type corpusFixturePaths struct {
	packagePath  string
	expectedPath string
}

// preflightCorpusManifest validates all declared aggregate work and every path
// before the first package or expectation is materialized. Existing Limits are
// corpus-wide ceilings here as well as per-package ceilings.
func preflightCorpusManifest(root string, manifest corpus.Manifest, limits Limits) ([]corpusFixturePaths, error) {
	paths := make([]corpusFixturePaths, len(manifest.Fixtures))
	usedPaths := make(map[string]string, len(manifest.Fixtures)*2)
	previousID := ""
	var packageBytes uint64
	var expandedBytes uint64
	var expectedBytes uint64
	var parts int
	for index, fixture := range manifest.Fixtures {
		if fixture.ID == "" || index > 0 && fixture.ID <= previousID {
			return nil, fmt.Errorf("officecompat: corpus fixtures must have strictly sorted unique IDs")
		}
		previousID = fixture.ID
		if fixture.Format != "docx" && fixture.Format != "pptx" && fixture.Format != "xlsx" {
			return nil, fmt.Errorf("officecompat: corpus fixture %q has invalid format", fixture.ID)
		}
		if fixture.Dialect != "strict" && fixture.Dialect != "transitional" {
			return nil, fmt.Errorf("officecompat: corpus fixture %q has invalid dialect", fixture.ID)
		}
		if fixture.Outcome != "accepted" && fixture.Outcome != "refused" {
			return nil, fmt.Errorf("officecompat: corpus fixture %q has invalid outcome", fixture.ID)
		}
		if fixture.Bytes <= 0 || uint64(fixture.Bytes) > limits.MaxPackageBytes || fixture.ExpandedBytes <= 0 || uint64(fixture.ExpandedBytes) > limits.MaxExpandedBytes || fixture.Parts <= 0 || fixture.Parts > limits.MaxParts {
			return nil, fmt.Errorf("officecompat: corpus fixture %q exceeds declared resource limits", fixture.ID)
		}
		if packageBytes > limits.MaxPackageBytes-uint64(fixture.Bytes) || expandedBytes > limits.MaxExpandedBytes-uint64(fixture.ExpandedBytes) || parts > limits.MaxParts-fixture.Parts {
			return nil, fmt.Errorf("officecompat: corpus exceeds cumulative declared resource limits")
		}
		packageBytes += uint64(fixture.Bytes)
		expandedBytes += uint64(fixture.ExpandedBytes)
		parts += fixture.Parts
		if !corpusSHA256Pattern.MatchString(fixture.SHA256) || !corpusSHA256Pattern.MatchString(fixture.ExpectedSHA256) {
			return nil, fmt.Errorf("officecompat: corpus fixture %q has invalid digest metadata", fixture.ID)
		}

		packagePath, err := corpusFilePath(root, fixture.Package)
		if err != nil {
			return nil, fmt.Errorf("officecompat: corpus fixture %q package path: %w", fixture.ID, err)
		}
		expectedPath, err := corpusFilePath(root, fixture.Expected)
		if err != nil {
			return nil, fmt.Errorf("officecompat: corpus fixture %q expectation path: %w", fixture.ID, err)
		}
		if path.Ext(fixture.Package) != "."+fixture.Format || path.Ext(fixture.Expected) != ".json" {
			return nil, fmt.Errorf("officecompat: corpus fixture %q paths do not match declared format", fixture.ID)
		}
		for _, candidate := range []struct {
			relative string
			role     string
		}{{fixture.Package, "package"}, {fixture.Expected, "expectation"}} {
			key := asciiLower(candidate.relative)
			if previous, duplicate := usedPaths[key]; duplicate {
				return nil, fmt.Errorf("officecompat: corpus fixture %q %s path duplicates %s", fixture.ID, candidate.role, previous)
			}
			usedPaths[key] = fixture.ID + " " + candidate.role
		}
		expectedSize, err := regularCorpusFileSize(expectedPath, maxCorpusExpectedBytes)
		if err != nil {
			return nil, fmt.Errorf("officecompat: corpus fixture %q expectation: %w", fixture.ID, err)
		}
		if expectedSize > limits.MaxPackageBytes || expectedBytes > limits.MaxPackageBytes-expectedSize {
			return nil, fmt.Errorf("officecompat: corpus expectations exceed %d cumulative bytes", limits.MaxPackageBytes)
		}
		expectedBytes += expectedSize
		paths[index] = corpusFixturePaths{packagePath: packagePath, expectedPath: expectedPath}
	}
	return paths, nil
}

func corpusFilePath(root, relative string) (string, error) {
	if relative == "" || strings.Contains(relative, "\\") || strings.HasPrefix(relative, "/") || path.Clean(relative) != relative || relative == "." || relative == ".." || strings.HasPrefix(relative, "../") {
		return "", fmt.Errorf("unsafe relative path %q", relative)
	}
	joined := filepath.Join(root, filepath.FromSlash(relative))
	relativeToRoot, err := filepath.Rel(root, joined)
	if err != nil || relativeToRoot == ".." || strings.HasPrefix(relativeToRoot, ".."+string(filepath.Separator)) {
		return "", fmt.Errorf("path escapes corpus root")
	}
	return joined, nil
}

func readExactCorpusFile(filename string, expectedSize, maximum uint64) ([]byte, error) {
	if expectedSize == 0 || expectedSize > maximum || expectedSize >= uint64(1<<63-1) {
		return nil, fmt.Errorf("declared file size exceeds limit")
	}
	linkInfo, err := os.Lstat(filename)
	if err != nil {
		return nil, err
	}
	if linkInfo.Mode()&os.ModeSymlink != 0 || !linkInfo.Mode().IsRegular() {
		return nil, fmt.Errorf("file must be regular and not a symbolic link")
	}
	file, err := os.Open(filename)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	info, err := file.Stat()
	if err != nil {
		return nil, err
	}
	if !info.Mode().IsRegular() || info.Size() < 0 || uint64(info.Size()) != expectedSize {
		return nil, fmt.Errorf("file size does not match manifest")
	}
	data, err := io.ReadAll(io.LimitReader(file, int64(expectedSize)+1))
	if err != nil {
		return nil, err
	}
	if uint64(len(data)) != expectedSize {
		return nil, fmt.Errorf("file changed while being read")
	}
	return data, nil
}

func verifyCorpusFileDigest(filename, expectedDigest string, maximum uint64) error {
	linkInfo, err := os.Lstat(filename)
	if err != nil {
		return err
	}
	if linkInfo.Mode()&os.ModeSymlink != 0 || !linkInfo.Mode().IsRegular() {
		return fmt.Errorf("file must be regular and not a symbolic link")
	}
	file, err := os.Open(filename)
	if err != nil {
		return err
	}
	defer file.Close()
	info, err := file.Stat()
	if err != nil {
		return err
	}
	if !info.Mode().IsRegular() || info.Size() <= 0 || uint64(info.Size()) > maximum {
		return fmt.Errorf("file size exceeds limit")
	}
	digest := sha256.New()
	written, err := io.Copy(digest, io.LimitReader(file, int64(maximum)+1))
	if err != nil {
		return err
	}
	if written != info.Size() || uint64(written) > maximum || hex.EncodeToString(digest.Sum(nil)) != expectedDigest {
		return fmt.Errorf("digest mismatch")
	}
	return nil
}

func regularCorpusFileSize(filename string, maximum uint64) (uint64, error) {
	info, err := os.Lstat(filename)
	if err != nil {
		return 0, err
	}
	if info.Mode()&os.ModeSymlink != 0 || !info.Mode().IsRegular() || info.Size() <= 0 || uint64(info.Size()) > maximum {
		return 0, fmt.Errorf("file must be regular, nonempty, and within the size limit")
	}
	return uint64(info.Size()), nil
}

func digestBytes(data []byte) string {
	digest := sha256.Sum256(data)
	return hex.EncodeToString(digest[:])
}
