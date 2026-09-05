package officecompat

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"path"
	"sort"
	"strings"
	"unicode"
)

const (
	contentTypesNamespace  = "http://schemas.openxmlformats.org/package/2006/content-types"
	relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships"
)

// QualifiedOPCPart records the effective content type for one package part.
// The list returned by QualifyOPC is sorted by the exact package spelling.
type QualifiedOPCPart struct {
	Name        string `json:"name"`
	ContentType string `json:"content_type"`
}

// QualifiedOPCRelationship records one relationship after resolving an
// internal target against its source part. External targets deliberately have
// an empty ResolvedPart and retain their exact target URI.
type QualifiedOPCRelationship struct {
	RelationshipsPart string `json:"relationships_part"`
	SourcePart        string `json:"source_part,omitempty"`
	ID                string `json:"id"`
	Type              string `json:"type"`
	Target            string `json:"target"`
	TargetMode        string `json:"target_mode,omitempty"`
	ResolvedPart      string `json:"resolved_part,omitempty"`
}

// OPCQualification is deterministic package-level evidence that content-type
// declarations and the complete relationship graph resolve without ambiguity.
type OPCQualification struct {
	Parts         []QualifiedOPCPart         `json:"parts"`
	Relationships []QualifiedOPCRelationship `json:"relationships"`
}

// QualifyOPC validates OPC content types and relationships in addition to the
// ZIP/package-name checks performed by Inspect. It rejects dangling internal
// targets, orphan relationship parts, duplicate IDs/declarations, unsafe
// targets, and parts without an effective content type.
func QualifyOPC(data []byte) (OPCQualification, error) {
	return QualifyOPCWithLimits(data, DefaultLimits())
}

// QualifyOPCWithLimits is QualifyOPC with an explicit cumulative XML budget.
func QualifyOPCWithLimits(data []byte, limits Limits) (OPCQualification, error) {
	pkg, err := inspectPackage(data, limits)
	if err != nil {
		return OPCQualification{}, err
	}
	remaining := opcXMLBudget{tokens: limits.MaxXMLTokens, attributes: limits.MaxXMLAttributes}
	contentTypesFile, ok := fileByAlias(pkg, asciiLower(contentTypesPart))
	if !ok {
		return OPCQualification{}, fmt.Errorf("officecompat: content-types part is unavailable")
	}
	defaults, overrides, err := parseContentTypes(contentTypesFile, limits, &remaining)
	if err != nil {
		return OPCQualification{}, fmt.Errorf("officecompat: qualify %s: %w", contentTypesFile.Name, err)
	}

	actualByAlias := make(map[string]string, len(pkg.inventory.Parts))
	for _, part := range pkg.inventory.Parts {
		alias, aliasErr := canonicalPartAlias(part.Name)
		if aliasErr != nil {
			return OPCQualification{}, aliasErr
		}
		actualByAlias[alias] = part.Name
	}
	for alias := range overrides {
		if _, exists := actualByAlias[alias]; !exists {
			return OPCQualification{}, fmt.Errorf("officecompat: content-type override targets absent part %q", alias)
		}
	}

	qualification := OPCQualification{}
	for _, part := range pkg.inventory.Parts {
		if asciiLower(part.Name) == asciiLower(contentTypesPart) {
			continue
		}
		alias, _ := canonicalPartAlias(part.Name)
		contentType := overrides[alias]
		if contentType == "" {
			extension := asciiLower(path.Ext(part.Name))
			if extension != "" {
				contentType = defaults[strings.TrimPrefix(extension, ".")]
			}
		}
		if contentType == "" {
			return OPCQualification{}, fmt.Errorf("officecompat: OPC part %q has no effective content type", part.Name)
		}
		if strings.HasSuffix(asciiLower(part.Name), ".rels") && contentType != "application/vnd.openxmlformats-package.relationships+xml" {
			return OPCQualification{}, fmt.Errorf("officecompat: relationship part %q has invalid content type %q", part.Name, contentType)
		}
		qualification.Parts = append(qualification.Parts, QualifiedOPCPart{Name: part.Name, ContentType: contentType})
	}

	for _, part := range pkg.inventory.Parts {
		if !strings.HasSuffix(asciiLower(part.Name), ".rels") {
			continue
		}
		sourceAlias, sourcePart, sourceErr := relationshipSource(part.Name, actualByAlias)
		if sourceErr != nil {
			return OPCQualification{}, sourceErr
		}
		relationshipFile := pkg.files[part.Name]
		relationships, parseErr := parseRelationships(relationshipFile, limits, &remaining)
		if parseErr != nil {
			return OPCQualification{}, fmt.Errorf("officecompat: qualify relationship part %q: %w", part.Name, parseErr)
		}
		for _, relationship := range relationships {
			record := QualifiedOPCRelationship{
				RelationshipsPart: part.Name,
				SourcePart:        sourcePart,
				ID:                relationship.ID,
				Type:              relationship.Type,
				Target:            relationship.Target,
				TargetMode:        relationship.TargetMode,
			}
			if relationship.TargetMode != "External" {
				resolvedAlias, resolveErr := resolveRelationshipTarget(sourceAlias, relationship.Target)
				if resolveErr != nil {
					return OPCQualification{}, fmt.Errorf("officecompat: relationship %q in %q: %w", relationship.ID, part.Name, resolveErr)
				}
				resolved, exists := actualByAlias[resolvedAlias]
				if !exists {
					return OPCQualification{}, fmt.Errorf("officecompat: relationship %q in %q has dangling internal target %q", relationship.ID, part.Name, relationship.Target)
				}
				record.ResolvedPart = resolved
			}
			qualification.Relationships = append(qualification.Relationships, record)
		}
	}

	sort.Slice(qualification.Parts, func(i, j int) bool { return qualification.Parts[i].Name < qualification.Parts[j].Name })
	sort.Slice(qualification.Relationships, func(i, j int) bool {
		left, right := qualification.Relationships[i], qualification.Relationships[j]
		if left.RelationshipsPart != right.RelationshipsPart {
			return left.RelationshipsPart < right.RelationshipsPart
		}
		return left.ID < right.ID
	})
	return qualification, nil
}

type opcXMLBudget struct {
	tokens     uint64
	attributes int
}

func parseContentTypes(file *zip.File, limits Limits, budget *opcXMLBudget) (map[string]string, map[string]string, error) {
	payload, err := readQualifiedXML(file, limits)
	if err != nil {
		return nil, nil, err
	}
	decoder := xml.NewDecoder(bytes.NewReader(payload))
	decoder.Strict = true
	defaults := map[string]string{}
	overrides := map[string]string{}
	depth := 0
	rootSeen := false
	for {
		token, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			break
		}
		if tokenErr != nil {
			return nil, nil, tokenErr
		}
		if err := consumeOPCTokenBudget(token, depth, limits, budget); err != nil {
			return nil, nil, err
		}
		switch value := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if rootSeen || value.Name.Space != contentTypesNamespace || value.Name.Local != "Types" {
					return nil, nil, fmt.Errorf("invalid content-types root element")
				}
				rootSeen = true
				if _, err := attributes(value.Attr, nil); err != nil {
					return nil, nil, err
				}
				continue
			}
			if depth != 2 || value.Name.Space != contentTypesNamespace {
				return nil, nil, fmt.Errorf("unexpected nested content-types element %q", value.Name.Local)
			}
			attrs, err := attributes(value.Attr, map[string]bool{"Extension": true, "PartName": true, "ContentType": true})
			if err != nil {
				return nil, nil, err
			}
			switch value.Name.Local {
			case "Default":
				extension, contentType := asciiLower(attrs["Extension"]), attrs["ContentType"]
				if extension == "" || strings.Contains(extension, ".") || contentType == "" || defaults[extension] != "" {
					return nil, nil, fmt.Errorf("invalid or duplicate content-type default")
				}
				defaults[extension] = contentType
			case "Override":
				partName, contentType := attrs["PartName"], attrs["ContentType"]
				if !strings.HasPrefix(partName, "/") || strings.HasPrefix(partName, "//") || contentType == "" {
					return nil, nil, fmt.Errorf("invalid content-type override")
				}
				alias, aliasErr := canonicalPartAlias(strings.TrimPrefix(partName, "/"))
				if aliasErr != nil || overrides[alias] != "" {
					return nil, nil, fmt.Errorf("invalid or duplicate content-type override %q", partName)
				}
				overrides[alias] = contentType
			default:
				return nil, nil, fmt.Errorf("unexpected content-types element %q", value.Name.Local)
			}
		case xml.EndElement:
			depth--
		case xml.CharData:
			if strings.TrimSpace(string(value)) != "" {
				return nil, nil, fmt.Errorf("content-types XML contains non-whitespace text")
			}
		case xml.Directive:
			return nil, nil, fmt.Errorf("content-types XML directives are unsupported")
		case xml.ProcInst:
			if !strings.EqualFold(value.Target, "xml") {
				return nil, nil, fmt.Errorf("content-types XML processing instruction is unsupported")
			}
		}
	}
	if !rootSeen || depth != 0 {
		return nil, nil, fmt.Errorf("content-types root is missing or unbalanced")
	}
	return defaults, overrides, nil
}

type rawRelationship struct {
	ID         string
	Type       string
	Target     string
	TargetMode string
}

func parseRelationships(file *zip.File, limits Limits, budget *opcXMLBudget) ([]rawRelationship, error) {
	payload, err := readQualifiedXML(file, limits)
	if err != nil {
		return nil, err
	}
	decoder := xml.NewDecoder(bytes.NewReader(payload))
	decoder.Strict = true
	depth := 0
	rootSeen := false
	seenIDs := map[string]bool{}
	var relationships []rawRelationship
	for {
		token, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			break
		}
		if tokenErr != nil {
			return nil, tokenErr
		}
		if err := consumeOPCTokenBudget(token, depth, limits, budget); err != nil {
			return nil, err
		}
		switch value := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if rootSeen || value.Name.Space != relationshipsNamespace || value.Name.Local != "Relationships" {
					return nil, fmt.Errorf("invalid relationships root element")
				}
				rootSeen = true
				if _, err := attributes(value.Attr, nil); err != nil {
					return nil, err
				}
				continue
			}
			if depth != 2 || value.Name.Space != relationshipsNamespace || value.Name.Local != "Relationship" {
				return nil, fmt.Errorf("unexpected relationships element %q", value.Name.Local)
			}
			attrs, err := attributes(value.Attr, map[string]bool{"Id": true, "Type": true, "Target": true, "TargetMode": true})
			if err != nil {
				return nil, err
			}
			relationship := rawRelationship{ID: attrs["Id"], Type: attrs["Type"], Target: attrs["Target"], TargetMode: attrs["TargetMode"]}
			if relationship.ID == "" || relationship.Type == "" || relationship.Target == "" || seenIDs[relationship.ID] {
				return nil, fmt.Errorf("relationship has missing fields or duplicate Id %q", relationship.ID)
			}
			typeURI, typeErr := url.Parse(relationship.Type)
			if typeErr != nil || !typeURI.IsAbs() {
				return nil, fmt.Errorf("relationship %q has non-absolute Type", relationship.ID)
			}
			if relationship.TargetMode != "" && relationship.TargetMode != "Internal" && relationship.TargetMode != "External" {
				return nil, fmt.Errorf("relationship %q has invalid TargetMode %q", relationship.ID, relationship.TargetMode)
			}
			if relationship.TargetMode == "External" {
				targetURI, targetErr := url.Parse(relationship.Target)
				if targetErr != nil || !targetURI.IsAbs() {
					return nil, fmt.Errorf("external relationship %q has non-absolute target", relationship.ID)
				}
			}
			seenIDs[relationship.ID] = true
			relationships = append(relationships, relationship)
		case xml.EndElement:
			depth--
		case xml.CharData:
			if strings.TrimSpace(string(value)) != "" {
				return nil, fmt.Errorf("relationships XML contains non-whitespace text")
			}
		case xml.Directive:
			return nil, fmt.Errorf("relationships XML directives are unsupported")
		case xml.ProcInst:
			if !strings.EqualFold(value.Target, "xml") {
				return nil, fmt.Errorf("relationships XML processing instruction is unsupported")
			}
		}
	}
	if !rootSeen || depth != 0 {
		return nil, fmt.Errorf("relationships root is missing or unbalanced")
	}
	return relationships, nil
}

func readQualifiedXML(file *zip.File, limits Limits) ([]byte, error) {
	if file.UncompressedSize64 > limits.MaxXMLBytes {
		return nil, fmt.Errorf("XML part exceeds %d bytes", limits.MaxXMLBytes)
	}
	reader, err := file.Open()
	if err != nil {
		return nil, err
	}
	payload, readErr := io.ReadAll(io.LimitReader(reader, int64(limits.MaxXMLBytes)+1))
	closeErr := reader.Close()
	if readErr != nil {
		return nil, readErr
	}
	if closeErr != nil {
		return nil, closeErr
	}
	if uint64(len(payload)) != file.UncompressedSize64 || uint64(len(payload)) > limits.MaxXMLBytes {
		return nil, fmt.Errorf("XML expanded size mismatch")
	}
	return payload, nil
}

func consumeOPCTokenBudget(token xml.Token, depth int, limits Limits, budget *opcXMLBudget) error {
	if budget.tokens == 0 {
		return fmt.Errorf("OPC XML token budget exceeded")
	}
	budget.tokens--
	start, ok := token.(xml.StartElement)
	if !ok {
		return nil
	}
	if depth+1 > limits.MaxXMLDepth {
		return fmt.Errorf("OPC XML depth exceeds %d", limits.MaxXMLDepth)
	}
	if len(start.Attr) > budget.attributes {
		return fmt.Errorf("OPC XML attribute budget exceeded")
	}
	budget.attributes -= len(start.Attr)
	return nil
}

func attributes(input []xml.Attr, allowed map[string]bool) (map[string]string, error) {
	result := map[string]string{}
	seen := map[string]bool{}
	for _, attribute := range input {
		if attribute.Name.Space == "xmlns" || attribute.Name.Space == "" && attribute.Name.Local == "xmlns" {
			continue
		}
		if attribute.Name.Space != "" || allowed == nil || !allowed[attribute.Name.Local] || seen[attribute.Name.Local] {
			return nil, fmt.Errorf("unexpected or duplicate attribute %q", attribute.Name.Local)
		}
		seen[attribute.Name.Local] = true
		result[attribute.Name.Local] = attribute.Value
	}
	return result, nil
}

func fileByAlias(pkg inspectedPackage, wanted string) (*zip.File, bool) {
	for name, file := range pkg.files {
		alias, err := canonicalPartAlias(name)
		if err == nil && alias == wanted {
			return file, true
		}
	}
	return nil, false
}

func relationshipSource(relationshipsPart string, actualByAlias map[string]string) (string, string, error) {
	alias, err := canonicalPartAlias(relationshipsPart)
	if err != nil {
		return "", "", err
	}
	if alias == asciiLower(rootRelsPart) {
		return "", "", nil
	}
	directory, filename := path.Split(relationshipsPart)
	directory = strings.TrimSuffix(directory, "/")
	parent, relsDirectory := path.Split(directory)
	if !strings.EqualFold(relsDirectory, "_rels") || !strings.HasSuffix(strings.ToLower(filename), ".rels") {
		return "", "", fmt.Errorf("officecompat: relationship part %q is not in an OPC relationship location", relationshipsPart)
	}
	sourceSpelling := path.Join(parent, filename[:len(filename)-len(".rels")])
	sourceAlias, aliasErr := canonicalPartAlias(sourceSpelling)
	if aliasErr != nil {
		return "", "", fmt.Errorf("officecompat: relationship part %q has invalid source mapping", relationshipsPart)
	}
	source, exists := actualByAlias[sourceAlias]
	if !exists {
		return "", "", fmt.Errorf("officecompat: relationship part %q has absent source part", relationshipsPart)
	}
	return sourceAlias, source, nil
}

func resolveRelationshipTarget(sourceAlias, target string) (string, error) {
	if strings.Contains(target, "\\") || strings.Contains(target, "//") {
		return "", fmt.Errorf("unsafe internal target %q", target)
	}
	parsed, err := url.Parse(target)
	if err != nil || parsed.IsAbs() || parsed.Host != "" || parsed.RawQuery != "" {
		return "", fmt.Errorf("unsafe internal target %q", target)
	}
	for _, character := range parsed.Fragment {
		if unicode.IsControl(character) || unicode.IsSpace(character) {
			return "", fmt.Errorf("unsafe internal target fragment in %q", target)
		}
	}
	var segments []string
	if !strings.HasPrefix(target, "/") && sourceAlias != "" {
		sourceDirectory := path.Dir(sourceAlias)
		if sourceDirectory != "." {
			segments = append(segments, strings.Split(sourceDirectory, "/")...)
		}
	}
	rawPath := strings.TrimPrefix(parsed.EscapedPath(), "/")
	if rawPath == "" {
		return "", fmt.Errorf("empty internal target")
	}
	for _, rawSegment := range strings.Split(rawPath, "/") {
		if rawSegment == "" {
			return "", fmt.Errorf("unsafe internal target %q", target)
		}
		decoded, decodeErr := url.PathUnescape(rawSegment)
		if decodeErr != nil || decoded == "" || strings.ContainsAny(decoded, "/\\?#%:") {
			return "", fmt.Errorf("unsafe internal target %q", target)
		}
		if decoded == "." || decoded == ".." {
			if decoded != rawSegment {
				return "", fmt.Errorf("encoded traversal in internal target %q", target)
			}
			if decoded == "." {
				continue
			}
			if len(segments) == 0 {
				return "", fmt.Errorf("internal target escapes package root")
			}
			segments = segments[:len(segments)-1]
			continue
		}
		for _, character := range decoded {
			if unicode.IsControl(character) || unicode.IsSpace(character) {
				return "", fmt.Errorf("unsafe internal target %q", target)
			}
		}
		segments = append(segments, decoded)
	}
	resolved := strings.Join(segments, "/")
	alias, aliasErr := canonicalPartAlias(resolved)
	if aliasErr != nil {
		return "", fmt.Errorf("unsafe internal target %q", target)
	}
	return alias, nil
}
