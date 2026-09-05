package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/xml"
	"fmt"
	"io"
	"sort"
	"strings"
)

// PreservationKind identifies office content that partial spreadsheet models
// commonly lose during a parse-and-regenerate save.
type PreservationKind string

const (
	PreserveChart     PreservationKind = "chart"
	PreservePivot     PreservationKind = "pivot"
	PreserveDrawing   PreservationKind = "drawing"
	PreserveMedia     PreservationKind = "media"
	PreserveEmbedding PreservationKind = "embedding"
)

// PreservationPart is either a real OPC part or a synthetic relationship
// (names beginning with @). Hash is SHA-256 of the real part.
type PreservationPart struct {
	Name   string           `json:"name"`
	Kind   PreservationKind `json:"kind"`
	Size   int64            `json:"size"`
	SHA256 string           `json:"sha256"`
}

type PreservationChange struct {
	Before PreservationPart `json:"before"`
	After  PreservationPart `json:"after"`
}

// PreservationReport describes protected content lost or altered by a full
// workbook replacement. Extra content in the candidate is allowed.
type PreservationReport struct {
	Missing []PreservationPart   `json:"missing,omitempty"`
	Changed []PreservationChange `json:"changed,omitempty"`
}

func (report PreservationReport) OK() bool {
	return len(report.Missing) == 0 && len(report.Changed) == 0
}

// PreservationError is returned by RequirePreservation with the full report.
type PreservationError struct {
	Report PreservationReport
}

func (err *PreservationError) Error() string {
	return fmt.Sprintf("xlsxpatch: replacement would lose %d and alter %d protected inventory items", len(err.Report.Missing), len(err.Report.Changed))
}

// PreservationInventory fingerprints charts, native pivots, DrawingML/VML,
// media, controls, and embedded objects. It also records relationships that
// point at those parts, preventing a replacement from retaining orphaned bytes
// while dropping the worksheet/workbook link that makes them visible.
func PreservationInventory(data []byte) ([]PreservationPart, error) {
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: preservation inventory: %w", err)
	}
	files := make(map[string]*zip.File, len(zr.File))
	for _, file := range zr.File {
		if _, duplicate := files[file.Name]; duplicate {
			return nil, fmt.Errorf("xlsxpatch: preservation inventory: duplicate entry %q", file.Name)
		}
		files[file.Name] = file
	}
	read := func(file *zip.File) ([]byte, error) {
		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		defer rc.Close()
		return io.ReadAll(rc)
	}

	var inventory []PreservationPart
	for name, file := range files {
		kind, protected := preservationKind(name)
		if !protected {
			continue
		}
		content, err := read(file)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: preservation inventory: read %q: %w", name, err)
		}
		digest := sha256.Sum256(content)
		inventory = append(inventory, PreservationPart{Name: name, Kind: kind, Size: int64(len(content)), SHA256: hex.EncodeToString(digest[:])})
	}

	// Ignore rId numbering and record semantic owner/type/target edges. This
	// tolerates harmless id renumbering while detecting disconnected parts.
	var semantic []PreservationPart
	for name, file := range files {
		if !strings.HasSuffix(name, ".rels") {
			continue
		}
		content, err := read(file)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: preservation inventory: read %q: %w", name, err)
		}
		rels, err := parsePackageRelationships(string(content))
		if err != nil {
			return nil, err
		}
		owner, base := relationshipOwner(name)
		for _, rel := range rels {
			external, modeErr := packageRelationshipIsExternal(rel)
			if modeErr != nil {
				return nil, fmt.Errorf("xlsxpatch: preservation inventory: relationship %q in %q: %w", rel.ID, name, modeErr)
			}
			if external {
				continue
			}
			target, resolveErr := resolveRelPath(base, rel.Target)
			if resolveErr != nil {
				return nil, fmt.Errorf("xlsxpatch: preservation inventory: relationship %q in %q: %w", rel.ID, name, resolveErr)
			}
			kind, protected := preservationKind(target)
			if !protected {
				continue
			}
			semantic = append(semantic, syntheticInventory("@rel", owner, rel.Type, target, kind))
		}
		if ownerFile, ok := files[owner]; ok {
			ownerXML, readErr := read(ownerFile)
			if readErr != nil {
				return nil, fmt.Errorf("xlsxpatch: preservation inventory: read %q: %w", owner, readErr)
			}
			references, refErr := protectedOwnerReferences(owner, base, ownerXML, rels)
			if refErr != nil {
				return nil, refErr
			}
			semantic = append(semantic, references...)
		}
	}
	inventory = append(inventory, uniquifySynthetic(semantic)...)
	sort.Slice(inventory, func(i, j int) bool { return inventory[i].Name < inventory[j].Name })
	return inventory, nil
}

// ComparePreservation compares an original workbook to a replacement.
// allowChanged contains exact real part names that an intentional edit may
// change or remove. Semantic edges targeting an allowed part are allowed too.
func ComparePreservation(original, candidate []byte, allowChanged []string) (PreservationReport, error) {
	before, err := PreservationInventory(original)
	if err != nil {
		return PreservationReport{}, err
	}
	after, err := PreservationInventory(candidate)
	if err != nil {
		return PreservationReport{}, err
	}
	afterByName := make(map[string]PreservationPart, len(after))
	for _, part := range after {
		afterByName[part.Name] = part
	}
	allowed := map[string]bool{}
	for _, name := range allowChanged {
		allowed[name] = true
	}
	isAllowed := func(name string) bool {
		if allowed[name] {
			return true
		}
		for target := range allowed {
			if strings.HasPrefix(name, "@") && strings.Contains(name, "|"+target) {
				return true
			}
		}
		return false
	}
	var report PreservationReport
	for _, part := range before {
		if isAllowed(part.Name) {
			continue
		}
		other, ok := afterByName[part.Name]
		if !ok {
			report.Missing = append(report.Missing, part)
			continue
		}
		if part.SHA256 != other.SHA256 || part.Size != other.Size {
			report.Changed = append(report.Changed, PreservationChange{Before: part, After: other})
		}
	}
	return report, nil
}

// RequirePreservation is the fail-closed replacement guard.
func RequirePreservation(original, candidate []byte, allowChanged []string) error {
	report, err := ComparePreservation(original, candidate, allowChanged)
	if err != nil {
		return err
	}
	if !report.OK() {
		return &PreservationError{Report: report}
	}
	return nil
}

func preservationKind(name string) (PreservationKind, bool) {
	switch {
	case strings.HasPrefix(name, "xl/charts/"):
		return PreserveChart, true
	case strings.HasPrefix(name, "xl/pivotTables/"), strings.HasPrefix(name, "xl/pivotCache/"), strings.HasPrefix(name, "xl/slicers/"), strings.HasPrefix(name, "xl/slicerCaches/"), strings.HasPrefix(name, "xl/timelines/"), strings.HasPrefix(name, "xl/timelineCaches/"):
		return PreservePivot, true
	case strings.HasPrefix(name, "xl/drawings/"), strings.HasPrefix(name, "xl/vmlDrawings/"), strings.HasPrefix(name, "xl/ctrlProps/"), strings.HasPrefix(name, "xl/activeX/"):
		return PreserveDrawing, true
	case strings.HasPrefix(name, "xl/media/"):
		return PreserveMedia, true
	case strings.HasPrefix(name, "xl/embeddings/"):
		return PreserveEmbedding, true
	default:
		return "", false
	}
}

func relationshipOwner(relsPart string) (owner, base string) {
	if relsPart == "_rels/.rels" {
		return "/", ""
	}
	marker := "/_rels/"
	index := strings.LastIndex(relsPart, marker)
	if index < 0 {
		lastSlash := strings.LastIndex(relsPart, "/")
		if lastSlash < 0 {
			return relsPart, ""
		}
		return relsPart, relsPart[:lastSlash]
	}
	directory := relsPart[:index]
	file := strings.TrimSuffix(relsPart[index+len(marker):], ".rels")
	return directory + "/" + file, directory
}

func syntheticInventory(prefix, owner, relType, target string, kind PreservationKind) PreservationPart {
	name := prefix + "/" + owner + "|" + relType + "|" + target
	digest := sha256.Sum256([]byte(name))
	return PreservationPart{Name: name, Kind: kind, Size: int64(len(name)), SHA256: hex.EncodeToString(digest[:])}
}

var protectedReferenceElements = map[string]bool{
	"drawing": true, "legacyDrawing": true, "legacyDrawingHF": true,
	"pivotTablePart": true, "pivotCache": true, "oleObject": true,
	"control": true, "picture": true,
}

// protectedOwnerReferences proves that a relationship is actually referenced
// by its workbook/worksheet owner. Keeping an orphan relationship and part is
// not preservation: Excel will no longer display the object.
func protectedOwnerReferences(owner, base string, ownerXML []byte, rels []packageRelationship) ([]PreservationPart, error) {
	byID := map[string]packageRelationship{}
	for _, rel := range rels {
		byID[rel.ID] = rel
	}
	dec := xml.NewDecoder(bytes.NewReader(ownerXML))
	var out []PreservationPart
	for {
		token, err := dec.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: preservation inventory: parse owner %q: %w", owner, err)
		}
		start, ok := token.(xml.StartElement)
		if !ok || !protectedReferenceElements[start.Name.Local] {
			continue
		}
		rel, ok := byID[attrVal(start, "id")]
		if !ok {
			continue
		}
		external, modeErr := packageRelationshipIsExternal(rel)
		if modeErr != nil {
			return nil, fmt.Errorf("xlsxpatch: preservation inventory: referenced relationship %q from %q: %w", rel.ID, owner, modeErr)
		}
		if external {
			continue
		}
		target, resolveErr := resolveRelPath(base, rel.Target)
		if resolveErr != nil {
			return nil, fmt.Errorf("xlsxpatch: preservation inventory: referenced relationship %q from %q: %w", rel.ID, owner, resolveErr)
		}
		kind, protected := preservationKind(target)
		if !protected {
			continue
		}
		out = append(out, syntheticInventory("@ref", owner, start.Name.Local, target, kind))
	}
	return out, nil
}

func uniquifySynthetic(parts []PreservationPart) []PreservationPart {
	sort.Slice(parts, func(i, j int) bool { return parts[i].Name < parts[j].Name })
	counts := map[string]int{}
	for index := range parts {
		base := parts[index].Name
		counts[base]++
		parts[index].Name = fmt.Sprintf("%s#%d", base, counts[base])
		digest := sha256.Sum256([]byte(parts[index].Name))
		parts[index].Size = int64(len(parts[index].Name))
		parts[index].SHA256 = hex.EncodeToString(digest[:])
	}
	return parts
}
