// Package officecompat provides format-neutral preservation checks for Office
// Open XML packages. It deliberately works at the OPC part boundary: callers
// name the parts an operation may mutate, and every other part must retain the
// same uncompressed payload fingerprint.
package officecompat

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"fmt"
	"io"
	"net/url"
	"sort"
	"strings"
	"unicode"
	"unicode/utf8"
)

const (
	contentTypesPart = "[Content_Types].xml"
	rootRelsPart     = "_rels/.rels"
)

// Part is a stable content-level fingerprint of one OPC part. ZIP compression,
// entry order, and timestamps are intentionally excluded: those details may be
// rewritten without changing the Office part itself.
type Part struct {
	Name   string `json:"name"`
	Size   uint64 `json:"size"`
	CRC32  uint32 `json:"crc32"`
	SHA256 string `json:"sha256"`
}

// Inventory is a deterministic list of OPC parts plus a package fingerprint.
// Fingerprint changes when a part is added, removed, renamed, or its payload
// changes, but not when ZIP metadata or compression changes.
type Inventory struct {
	Parts       []Part `json:"parts"`
	Fingerprint string `json:"fingerprint"`
}

// Change records a part whose payload differs between packages.
type Change struct {
	Before Part           `json:"before"`
	After  Part           `json:"after"`
	XML    *XMLComparison `json:"xml,omitempty"`
}

// PreservationReport describes every mutation outside the caller's exact
// mutable-part allowlist.
type PreservationReport struct {
	Missing []Part   `json:"missing,omitempty"`
	Changed []Change `json:"changed,omitempty"`
	Added   []Part   `json:"added,omitempty"`
}

// OK reports whether every mutation was confined to the mutable allowlist.
func (report PreservationReport) OK() bool {
	return len(report.Missing) == 0 && len(report.Changed) == 0 && len(report.Added) == 0
}

// PreservationError is returned by RequireUntouchedParts when an operation
// changes an OPC part outside its declared mutation boundary.
type PreservationError struct {
	Report PreservationReport
}

func (err *PreservationError) Error() string {
	return fmt.Sprintf(
		"officecompat: mutation escaped allowlist: %d missing, %d changed, %d added parts",
		len(err.Report.Missing), len(err.Report.Changed), len(err.Report.Added),
	)
}

// Inspect fingerprints every non-directory part in an OPC package. It rejects
// ambiguous duplicate names, unsafe/non-canonical paths, malformed ZIP data,
// and ZIP files that do not contain the two required OPC root parts.
func Inspect(data []byte) (Inventory, error) {
	return InspectWithLimits(data, DefaultLimits())
}

// InspectWithLimits is Inspect with an explicit fail-closed resource envelope.
func InspectWithLimits(data []byte, limits Limits) (Inventory, error) {
	pkg, err := inspectPackage(data, limits)
	if err != nil {
		return Inventory{}, err
	}
	return pkg.inventory, nil
}

type inspectedPackage struct {
	inventory Inventory
	byName    map[string]Part
	files     map[string]*zip.File
}

func inspectPackage(data []byte, limits Limits) (inspectedPackage, error) {
	if err := limits.validate(); err != nil {
		return inspectedPackage{}, err
	}
	if len(data) == 0 || uint64(len(data)) > limits.MaxPackageBytes {
		return inspectedPackage{}, fmt.Errorf("officecompat: OPC package size must be 1..%d bytes", limits.MaxPackageBytes)
	}
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return inspectedPackage{}, fmt.Errorf("officecompat: open OPC package: %w", err)
	}
	if len(zr.File) == 0 || len(zr.File) > limits.MaxParts {
		return inspectedPackage{}, fmt.Errorf("officecompat: OPC package exceeds %d entries", limits.MaxParts)
	}
	seen := make(map[string]bool, len(zr.File))
	aliases := make(map[string]string, len(zr.File))
	files := make(map[string]*zip.File, len(zr.File))
	parts := make([]Part, 0, len(zr.File))
	var declaredTotal uint64
	var actualTotal uint64
	for _, file := range zr.File {
		if file.FileInfo().IsDir() || strings.HasSuffix(file.Name, "/") {
			if err := validateDirectoryName(file.Name); err != nil {
				return inspectedPackage{}, err
			}
			continue
		}
		alias, err := canonicalPartAlias(file.Name)
		if err != nil {
			return inspectedPackage{}, err
		}
		if seen[file.Name] {
			return inspectedPackage{}, fmt.Errorf("officecompat: duplicate OPC part %q", file.Name)
		}
		seen[file.Name] = true
		if previous, collision := aliases[alias]; collision {
			return inspectedPackage{}, fmt.Errorf("officecompat: case/percent-equivalent OPC parts %q and %q", previous, file.Name)
		}
		aliases[alias] = file.Name
		if file.Flags&0x1 != 0 {
			return inspectedPackage{}, fmt.Errorf("officecompat: encrypted OPC part %q is unsupported", file.Name)
		}
		if file.Method != zip.Store && file.Method != zip.Deflate {
			return inspectedPackage{}, fmt.Errorf("officecompat: OPC part %q uses unsupported compression method %d", file.Name, file.Method)
		}
		if file.UncompressedSize64 > limits.MaxPartBytes {
			return inspectedPackage{}, fmt.Errorf("officecompat: part %q exceeds %d expanded bytes", file.Name, limits.MaxPartBytes)
		}
		if exceedsCompressionRatioWithLimits(file.UncompressedSize64, file.CompressedSize64, limits.MaxCompressionRatio, limits.CompressionRatioSlack) {
			return inspectedPackage{}, fmt.Errorf("officecompat: part %q exceeds compression ratio limit", file.Name)
		}
		if file.UncompressedSize64 > limits.MaxExpandedBytes-declaredTotal {
			return inspectedPackage{}, fmt.Errorf("officecompat: OPC package exceeds %d expanded bytes", limits.MaxExpandedBytes)
		}
		declaredTotal += file.UncompressedSize64

		rc, err := file.Open()
		if err != nil {
			return inspectedPackage{}, fmt.Errorf("officecompat: open part %q: %w", file.Name, err)
		}
		digest := sha256.New()
		readBytes, readErr := io.Copy(digest, io.LimitReader(rc, int64(limits.MaxPartBytes)+1))
		closeErr := rc.Close()
		if readErr != nil {
			return inspectedPackage{}, fmt.Errorf("officecompat: read part %q: %w", file.Name, readErr)
		}
		if closeErr != nil {
			return inspectedPackage{}, fmt.Errorf("officecompat: close part %q: %w", file.Name, closeErr)
		}
		if readBytes < 0 || uint64(readBytes) > limits.MaxPartBytes {
			return inspectedPackage{}, fmt.Errorf("officecompat: expanded size mismatch for part %q", file.Name)
		}
		if uint64(readBytes) != file.UncompressedSize64 {
			return inspectedPackage{}, fmt.Errorf("officecompat: expanded size mismatch for part %q", file.Name)
		}
		if uint64(readBytes) > limits.MaxExpandedBytes-actualTotal {
			return inspectedPackage{}, fmt.Errorf("officecompat: OPC package exceeds %d expanded bytes", limits.MaxExpandedBytes)
		}
		actualTotal += uint64(readBytes)
		parts = append(parts, Part{
			Name:   file.Name,
			Size:   uint64(readBytes),
			CRC32:  file.CRC32,
			SHA256: hex.EncodeToString(digest.Sum(nil)),
		})
		files[file.Name] = file
	}

	for _, required := range []string{contentTypesPart, rootRelsPart} {
		requiredAlias, _ := canonicalPartAlias(required)
		if _, exists := aliases[requiredAlias]; !exists {
			return inspectedPackage{}, fmt.Errorf("officecompat: missing required OPC part %q", required)
		}
	}

	sort.Slice(parts, func(i, j int) bool { return parts[i].Name < parts[j].Name })
	inventory := Inventory{Parts: parts, Fingerprint: inventoryFingerprint(parts)}
	return inspectedPackage{inventory: inventory, byName: indexParts(parts), files: files}, nil
}

// CompareUntouchedParts compares two OPC packages. mutableParts is an exact
// allowlist of parts that may be added, removed, or changed by the operation.
// Renaming a part is represented as one removal and one addition.
func CompareUntouchedParts(before, after []byte, mutableParts []string) (PreservationReport, error) {
	report, _, err := compareDeclaredPartMutation(before, after, mutableParts)
	return report, err
}

// compareDeclaredPartMutation reports both mutations outside the exact
// allowlist and whether at least one allowlisted part changed at the content
// level. ZIP order, timestamps, compression, and other container metadata do
// not count as an Office mutation.
func compareDeclaredPartMutation(before, after []byte, mutableParts []string) (PreservationReport, bool, error) {
	return comparePackages(before, after, mutableParts, DefaultLimits(), true)
}

// ComparePackages reports every lexical package-part change and, for changed
// XML parts, the bounded structural comparison. Exact payload bytes remain the
// preservation authority even when XML is structurally equivalent.
func ComparePackages(before, after []byte) (PreservationReport, error) {
	return ComparePackagesWithLimits(before, after, DefaultLimits())
}

// ComparePackagesWithLimits is ComparePackages with explicit resource limits.
func ComparePackagesWithLimits(before, after []byte, limits Limits) (PreservationReport, error) {
	report, _, err := comparePackages(before, after, nil, limits, false)
	return report, err
}

// CompareUntouchedPartsWithLimits is CompareUntouchedParts with explicit
// package and XML resource limits.
func CompareUntouchedPartsWithLimits(before, after []byte, mutableParts []string, limits Limits) (PreservationReport, error) {
	report, _, err := comparePackages(before, after, mutableParts, limits, true)
	return report, err
}

func comparePackages(before, after []byte, mutableParts []string, limits Limits, enforceAllowlist bool) (PreservationReport, bool, error) {
	beforePackage, err := inspectPackage(before, limits)
	if err != nil {
		return PreservationReport{}, false, fmt.Errorf("officecompat: inspect original: %w", err)
	}
	afterPackage, err := inspectPackage(after, limits)
	if err != nil {
		return PreservationReport{}, false, fmt.Errorf("officecompat: inspect candidate: %w", err)
	}

	allowed := map[string]bool{}
	if enforceAllowlist {
		allowed, err = validateMutableParts(mutableParts, beforePackage.byName, afterPackage.byName)
		if err != nil {
			return PreservationReport{}, false, err
		}
	}

	var report PreservationReport
	mutableChanged := false
	for _, part := range beforePackage.inventory.Parts {
		candidate, exists := afterPackage.byName[part.Name]
		if allowed[part.Name] {
			if !exists || !samePartPayload(part, candidate) {
				mutableChanged = true
			}
			continue
		}
		if !exists {
			report.Missing = append(report.Missing, part)
			continue
		}
		if !samePartPayload(part, candidate) {
			change := Change{Before: part, After: candidate}
			if isXMLPart(part.Name) {
				comparison := compareXMLFiles(beforePackage.files[part.Name], afterPackage.files[part.Name], part.SHA256, candidate.SHA256, limits)
				change.XML = &comparison
			}
			report.Changed = append(report.Changed, change)
		}
	}
	for _, part := range afterPackage.inventory.Parts {
		if allowed[part.Name] {
			if _, existed := beforePackage.byName[part.Name]; !existed {
				mutableChanged = true
			}
			continue
		}
		if _, existed := beforePackage.byName[part.Name]; !existed {
			report.Added = append(report.Added, part)
		}
	}
	return report, mutableChanged, nil
}

// RequireUntouchedParts fails closed when CompareUntouchedParts finds any
// mutation outside the exact mutable-part allowlist.
func RequireUntouchedParts(before, after []byte, mutableParts []string) error {
	report, err := CompareUntouchedParts(before, after, mutableParts)
	if err != nil {
		return err
	}
	if !report.OK() {
		return &PreservationError{Report: report}
	}
	return nil
}

func validatePartName(name string) error {
	_, err := canonicalPartAlias(name)
	return err
}

func validateDirectoryName(name string) error {
	if !strings.HasSuffix(name, "/") || name == "/" {
		return fmt.Errorf("officecompat: unsafe or non-canonical OPC directory name %q", name)
	}
	trimmed := strings.TrimSuffix(name, "/")
	if asciiLower(trimmed) == asciiLower(contentTypesPart) {
		return fmt.Errorf("officecompat: unsafe or non-canonical OPC directory name %q", name)
	}
	_, err := canonicalPartAlias(trimmed)
	if err != nil {
		return fmt.Errorf("officecompat: unsafe or non-canonical OPC directory name %q", name)
	}
	return nil
}

func canonicalPartAlias(name string) (string, error) {
	if name == "" || len(name) > 4096 || !utf8.ValidString(name) || strings.HasPrefix(name, "/") || strings.HasSuffix(name, "/") || strings.Contains(name, "\\") || strings.Contains(name, "//") {
		return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
	}
	// OPC reserves this one root part name. Square brackets are not RFC 3986
	// pchar characters and remain forbidden everywhere else (and in escaped
	// aliases of the reserved spelling).
	if asciiLower(name) == asciiLower(contentTypesPart) {
		return asciiLower(contentTypesPart), nil
	}
	segments := strings.Split(name, "/")
	for index, segment := range segments {
		if segment == "" {
			return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
		}
		for _, character := range segment {
			if unicode.IsSpace(character) {
				return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
			}
		}
		for offset := 0; offset < len(segment); offset++ {
			character := segment[offset]
			if character != '%' {
				if !isOPCPChar(character) {
					return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
				}
				continue
			}
			if offset+2 >= len(segment) || !isUpperHex(segment[offset+1]) || !isUpperHex(segment[offset+2]) {
				return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
			}
			offset += 2
		}
		decoded, err := url.PathUnescape(segment)
		if err != nil || !utf8.ValidString(decoded) || decoded == "" || decoded == "." || decoded == ".." || strings.HasSuffix(decoded, ".") || strings.ContainsAny(decoded, `/\\?#%:`) {
			return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
		}
		for _, character := range decoded {
			if unicode.IsControl(character) || unicode.IsSpace(character) || unicode.In(character, unicode.Cf, unicode.Zl, unicode.Zp) {
				return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
			}
			if character <= 0x7f && !isOPCPChar(byte(character)) {
				return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
			}
		}
		segments[index] = asciiLower(decoded)
	}
	alias := strings.Join(segments, "/")
	if alias == asciiLower(contentTypesPart) {
		return "", fmt.Errorf("officecompat: unsafe or non-canonical OPC part name %q", name)
	}
	return alias, nil
}

func isOPCPChar(value byte) bool {
	return value >= 'a' && value <= 'z' || value >= 'A' && value <= 'Z' || value >= '0' && value <= '9' ||
		strings.ContainsRune("-._~!$&'()*+,;=@", rune(value))
}

func isUpperHex(value byte) bool {
	return value >= '0' && value <= '9' || value >= 'A' && value <= 'F'
}

func asciiLower(value string) string {
	data := []byte(value)
	for index, character := range data {
		if character >= 'A' && character <= 'Z' {
			data[index] = character + ('a' - 'A')
		}
	}
	return string(data)
}

func validateMutableParts(names []string, before, after map[string]Part) (map[string]bool, error) {
	allowed := make(map[string]bool, len(names))
	aliases := make(map[string]string, len(names))
	for _, name := range names {
		alias, err := canonicalPartAlias(name)
		if err != nil {
			return nil, fmt.Errorf("officecompat: invalid mutable part: %w", err)
		}
		if allowed[name] {
			return nil, fmt.Errorf("officecompat: duplicate mutable part %q", name)
		}
		if previous, collision := aliases[alias]; collision {
			return nil, fmt.Errorf("officecompat: case/percent-equivalent mutable parts %q and %q", previous, name)
		}
		if _, existsBefore := before[name]; !existsBefore {
			if _, existsAfter := after[name]; !existsAfter {
				return nil, fmt.Errorf("officecompat: mutable part %q is absent from both packages", name)
			}
		}
		allowed[name] = true
		aliases[alias] = name
	}
	return allowed, nil
}

func exceedsCompressionRatioWithLimits(expanded, compressed, ratio, slack uint64) bool {
	if compressed > (^uint64(0)-slack)/ratio {
		return false
	}
	return expanded > compressed*ratio+slack
}

func indexParts(parts []Part) map[string]Part {
	indexed := make(map[string]Part, len(parts))
	for _, part := range parts {
		indexed[part.Name] = part
	}
	return indexed
}

func samePartPayload(left, right Part) bool {
	return left.Size == right.Size && left.CRC32 == right.CRC32 && left.SHA256 == right.SHA256
}

func exceedsCompressionRatio(uncompressed, compressed uint64) bool {
	return exceedsCompressionRatioWithLimits(uncompressed, compressed, MaxCompressionRatio, defaultCompressionRatioSlack)
}

func inventoryFingerprint(parts []Part) string {
	digest := sha256.New()
	var length [8]byte
	for _, part := range parts {
		binary.BigEndian.PutUint64(length[:], uint64(len(part.Name)))
		digest.Write(length[:])
		digest.Write([]byte(part.Name))
		binary.BigEndian.PutUint64(length[:], part.Size)
		digest.Write(length[:])
		partDigest, _ := hex.DecodeString(part.SHA256)
		digest.Write(partDigest)
	}
	return hex.EncodeToString(digest.Sum(nil))
}
