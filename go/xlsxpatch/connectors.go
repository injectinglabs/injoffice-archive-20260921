package xlsxpatch

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"strings"
)

const (
	ConnectorExtensionPartNameV1         = "customXml/injofficeConnectors.xml"
	ConnectorExtensionRelationshipTypeV1 = "https://schemas.injoffice.dev/relationships/connectors/2026"
	ConnectorExtensionContentTypeV1      = "application/vnd.injoffice.connectors+xml"
	ConnectorExtensionNamespaceV1        = "https://schemas.injoffice.dev/xlsx/connectors/2026"
	connectorPartMaxBytes                = 1 << 20
	connectorDefinitionLimit             = 1024
)

// ConnectorDefinition is the credential-free connector subset persisted in
// InjOffice's custom OPC extension. It stores definitions only: fetched data,
// credentials, authorization state, errors, and cache contents are excluded.
type ConnectorDefinition struct {
	ID       string                `json:"id"`
	Name     string                `json:"name"`
	Source   ConnectorSource       `json:"source"`
	Target   ConnectorTarget       `json:"target"`
	Refresh  string                `json:"refresh"`
	Schedule *ConnectorSchedule    `json:"schedule,omitempty"`
	Cache    *ConnectorCachePolicy `json:"cache,omitempty"`
	Schema   *ConnectorSchema      `json:"schema,omitempty"`
}

type ConnectorSource struct {
	Kind   string `json:"kind"`
	URL    string `json:"url"`
	Format string `json:"format"`
	Path   string `json:"path,omitempty"`
}

type ConnectorTarget struct {
	SheetID     string `json:"sheetId"`
	StartRow    int    `json:"startRow"`
	StartColumn int    `json:"startColumn"`
}

type ConnectorSchedule struct {
	IntervalMS int64 `json:"intervalMs"`
}

type ConnectorCachePolicy struct {
	Mode  string `json:"mode"`
	TTLMS int64  `json:"ttlMs"`
}

type ConnectorSchema struct {
	Columns                []ConnectorColumnSchema `json:"columns"`
	AllowAdditionalColumns bool                    `json:"allowAdditionalColumns,omitempty"`
}

type ConnectorColumnSchema struct {
	Index    int    `json:"index"`
	Type     string `json:"type"`
	Nullable bool   `json:"nullable,omitempty"`
}

type connectorPayloadV1 struct {
	Version    int                   `json:"version"`
	Connectors []ConnectorDefinition `json:"connectors"`
}

type connectorPartRoute struct {
	workbook workbookPartLocation
	part     string
	relID    string
}

// ReadConnectorDefinitions reads InjOffice's explicitly versioned custom OPC
// extension. A missing extension returns an empty slice. Malformed, ambiguous,
// externally targeted, or unknown-version InjOffice state is refused.
func ReadConnectorDefinitions(data []byte) ([]ConnectorDefinition, error) {
	pkg, err := openNativeWorkbookPackage(data)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read connectors: %w", err)
	}
	route, found, err := locateConnectorPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read connectors: %w", err)
	}
	if !found {
		return []ConnectorDefinition{}, nil
	}
	definitions, err := decodeConnectorPart(pkg.files[route.part])
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: read connectors: %w", err)
	}
	return definitions, nil
}

// SetConnectorDefinitions atomically creates, replaces, or removes only the
// InjOffice connector part and its workbook relationship/content-type entry.
// Apply raw-copies and verifies every other OPC part. Existing unrepresentable
// InjOffice connector state is never overwritten.
func SetConnectorDefinitions(orig []byte, definitions []ConnectorDefinition) ([]byte, error) {
	if err := validateConnectorDefinitions(definitions); err != nil {
		return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
	}
	pkg, err := openNativeWorkbookPackage(orig)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
	}
	route, found, err := locateConnectorPart(pkg)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
	}
	if found {
		if _, err := decodeConnectorPart(pkg.files[route.part]); err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors refused existing state: %w", err)
		}
	}

	patch := Patch{}
	switch {
	case len(definitions) == 0 && !found:
		patch = Patch{}
	case len(definitions) == 0:
		rels, err := connectorRelationshipsWithout(pkg.files[route.workbook.relsPart], route)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
		}
		contentTypes, err := contentTypesWithoutParts(pkg.files[pkg.contentTypesPart], []string{route.part})
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
		}
		patch = Patch{
			Replace: map[string][]byte{route.workbook.relsPart: rels, pkg.contentTypesPart: contentTypes},
			Delete:  map[string]bool{route.part: true},
		}
	case found:
		part, err := encodeConnectorPart(definitions)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
		}
		patch = Patch{Replace: map[string][]byte{route.part: part}}
	default:
		part, err := encodeConnectorPart(definitions)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
		}
		relsXML := string(pkg.files[route.workbook.relsPart])
		relID := nextFreeRelID(relsXML)
		updatedRels, err := appendRelationship(relsXML, relID, ConnectorExtensionRelationshipTypeV1, partRelTargetFrom(route.workbook.part, ConnectorExtensionPartNameV1))
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
		}
		updatedContentTypes, err := contentTypesWith(string(pkg.files[pkg.contentTypesPart]), map[string]string{"/" + ConnectorExtensionPartNameV1: ConnectorExtensionContentTypeV1})
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: set connectors: %w", err)
		}
		patch = Patch{
			Replace: map[string][]byte{route.workbook.relsPart: []byte(updatedRels), pkg.contentTypesPart: []byte(updatedContentTypes)},
			Add:     map[string][]byte{ConnectorExtensionPartNameV1: part},
		}
	}

	out, err := Apply(orig, patch)
	if err != nil {
		return nil, err
	}
	readBack, err := ReadConnectorDefinitions(out)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: set connectors: verify round trip: %w", err)
	}
	if definitions == nil {
		definitions = []ConnectorDefinition{}
	}
	want, _ := json.Marshal(definitions)
	got, _ := json.Marshal(readBack)
	if !bytes.Equal(want, got) {
		return nil, fmt.Errorf("xlsxpatch: set connectors: verify round trip mismatch")
	}
	return out, nil
}

func locateConnectorPart(pkg *nativeWorkbookPackage) (connectorPartRoute, bool, error) {
	read := func(name string) (string, bool) {
		value, ok := pkg.files[name]
		return string(value), ok
	}
	workbook, err := locateWorkbookPart(pkg.index, read)
	if err != nil {
		return connectorPartRoute{}, false, err
	}
	route := connectorPartRoute{workbook: workbook}
	rels := pkg.files[workbook.relsPart]
	relationships, err := parseRoutingRelationships(rels)
	if err != nil {
		return route, false, fmt.Errorf("parse workbook relationships: %w", err)
	}
	for _, relationship := range relationships {
		if relationship.relType != ConnectorExtensionRelationshipTypeV1 {
			continue
		}
		if route.relID != "" {
			return route, false, fmt.Errorf("multiple InjOffice connector relationships")
		}
		if relationship.targetMode != "" && relationship.targetMode != "Internal" {
			return route, false, fmt.Errorf("connector relationship %q is not internal", relationship.id)
		}
		resolved, err := resolveRelPath(workbook.baseDir, relationship.target)
		if err != nil {
			return route, false, fmt.Errorf("connector relationship %q target: %w", relationship.id, err)
		}
		part, _, ok := pkg.index.lookupResolved(resolved)
		if !ok {
			return route, false, fmt.Errorf("connector relationship %q targets missing part %q", relationship.id, resolved)
		}
		key, _ := canonicalOPCPartKey(part)
		wanted, _ := canonicalOPCPartKey(ConnectorExtensionPartNameV1)
		if key != wanted || part != ConnectorExtensionPartNameV1 {
			return route, false, fmt.Errorf("connector relationship %q targets unsupported part %q", relationship.id, part)
		}
		route.relID, route.part = relationship.id, part
	}
	if route.relID == "" {
		if spelling, _, exists := pkg.index.lookupSpelling(ConnectorExtensionPartNameV1); exists {
			return route, false, fmt.Errorf("orphan connector part %q has no workbook relationship", spelling)
		}
		return route, false, nil
	}
	if err := requireConnectorContentType(pkg.files[pkg.contentTypesPart], route.part); err != nil {
		return route, false, err
	}
	references, err := packageRelationshipReferenceCount(pkg.files, route.part)
	if err != nil {
		return route, false, err
	}
	if references != 1 {
		return route, false, fmt.Errorf("connector part %q has %d package relationships, expected 1", route.part, references)
	}
	return route, true, nil
}

func connectorRelationshipsWithout(data []byte, route connectorPartRoute) ([]byte, error) {
	rels, _, err := directChildElements(data, "Relationships", "Relationship")
	if err != nil {
		return nil, fmt.Errorf("parse workbook relationships: %w", err)
	}
	var matches []xmlSpan
	for _, relationship := range rels {
		if attribute(relationship.start, "Id") != route.relID {
			continue
		}
		if attribute(relationship.start, "Type") != ConnectorExtensionRelationshipTypeV1 {
			return nil, fmt.Errorf("connector relationship %q changed type", route.relID)
		}
		if mode := attribute(relationship.start, "TargetMode"); mode != "" && mode != "Internal" {
			return nil, fmt.Errorf("connector relationship %q is not internal", route.relID)
		}
		resolved, err := resolveRelPath(route.workbook.baseDir, attribute(relationship.start, "Target"))
		if err != nil || resolved != route.part {
			return nil, fmt.Errorf("connector relationship %q changed target", route.relID)
		}
		matches = append(matches, relationship.span)
	}
	if len(matches) != 1 {
		return nil, fmt.Errorf("connector relationship %q has %d entries, expected 1", route.relID, len(matches))
	}
	return removeElementSpans(data, matches), nil
}

func requireConnectorContentType(data []byte, part string) error {
	overrides, _, err := directChildElements(data, "Types", "Override")
	if err != nil {
		return fmt.Errorf("parse content types: %w", err)
	}
	wanted, _ := canonicalOPCPartKey(part)
	matches := 0
	for _, override := range overrides {
		name := attribute(override.start, "PartName")
		key, err := canonicalOPCPartKey(name)
		if err != nil {
			return fmt.Errorf("invalid content-type PartName %q: %w", name, err)
		}
		if key != wanted {
			continue
		}
		matches++
		if contentType := attribute(override.start, "ContentType"); contentType != ConnectorExtensionContentTypeV1 {
			return fmt.Errorf("connector part %q has content type %q, expected %q", part, contentType, ConnectorExtensionContentTypeV1)
		}
	}
	if matches != 1 {
		return fmt.Errorf("connector part %q has %d content-type overrides, expected 1", part, matches)
	}
	return nil
}

func encodeConnectorPart(definitions []ConnectorDefinition) ([]byte, error) {
	payload, err := json.Marshal(connectorPayloadV1{Version: 1, Connectors: definitions})
	if err != nil {
		return nil, err
	}
	encoded := base64.StdEncoding.EncodeToString(payload)
	part := []byte(xml.Header + `<connectors xmlns="` + ConnectorExtensionNamespaceV1 + `" version="1" encoding="base64-json">` + encoded + `</connectors>`)
	if len(part) > connectorPartMaxBytes {
		return nil, fmt.Errorf("connector extension exceeds %d bytes", connectorPartMaxBytes)
	}
	return part, nil
}

func decodeConnectorPart(data []byte) ([]ConnectorDefinition, error) {
	if len(data) == 0 || len(data) > connectorPartMaxBytes {
		return nil, fmt.Errorf("connector extension size must be 1..%d bytes", connectorPartMaxBytes)
	}
	if _, _, err := preflightNativeCoreXML(data); err != nil {
		return nil, fmt.Errorf("connector extension XML preflight: %w", err)
	}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth, rootSeen, rootClosed := 0, false, false
	var encoded strings.Builder
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, fmt.Errorf("parse connector extension XML: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth != 1 || rootSeen || token.Name != (xml.Name{Space: ConnectorExtensionNamespaceV1, Local: "connectors"}) {
				return nil, fmt.Errorf("connector extension must contain exactly one namespaced connectors root")
			}
			if err := requireOnlySemanticXMLAttributes(token, xml.Name{Local: "version"}, xml.Name{Local: "encoding"}); err != nil {
				return nil, err
			}
			version, versionFound, err := unqualifiedXMLAttribute(token, "version")
			if err != nil || !versionFound || version != "1" {
				return nil, fmt.Errorf("unsupported connector extension version %q", version)
			}
			encodingName, encodingFound, err := unqualifiedXMLAttribute(token, "encoding")
			if err != nil || !encodingFound || encodingName != "base64-json" {
				return nil, fmt.Errorf("unsupported connector extension encoding %q", encodingName)
			}
			rootSeen = true
		case xml.EndElement:
			if depth != 1 || token.Name != (xml.Name{Space: ConnectorExtensionNamespaceV1, Local: "connectors"}) {
				return nil, fmt.Errorf("connector extension has mismatched XML elements")
			}
			rootClosed = true
			depth--
		case xml.CharData:
			if depth == 1 {
				encoded.Write(token)
			} else if len(bytes.TrimSpace(token)) != 0 {
				return nil, fmt.Errorf("connector extension has text outside its root")
			}
		case xml.ProcInst:
			if rootSeen || token.Target != "xml" || !validNativeXMLDeclaration(token) {
				return nil, fmt.Errorf("connector extension has unsupported processing instruction")
			}
		case xml.Comment, xml.Directive:
			return nil, fmt.Errorf("connector extension has unsupported comment or directive")
		}
	}
	if depth != 0 || !rootSeen || !rootClosed {
		return nil, fmt.Errorf("connector extension has no complete root")
	}
	if encoded.Len() == 0 || strings.TrimSpace(encoded.String()) != encoded.String() {
		return nil, fmt.Errorf("connector extension payload is empty or contains surrounding whitespace")
	}
	payload, err := base64.StdEncoding.Strict().DecodeString(encoded.String())
	if err != nil {
		return nil, fmt.Errorf("decode connector extension payload: %w", err)
	}
	var envelope connectorPayloadV1
	jsonDecoder := json.NewDecoder(bytes.NewReader(payload))
	jsonDecoder.DisallowUnknownFields()
	if err := jsonDecoder.Decode(&envelope); err != nil {
		return nil, fmt.Errorf("decode connector extension JSON: %w", err)
	}
	if err := requireJSONEOF(jsonDecoder); err != nil {
		return nil, fmt.Errorf("decode connector extension JSON: %w", err)
	}
	if envelope.Version != 1 {
		return nil, fmt.Errorf("unsupported connector payload version %d", envelope.Version)
	}
	if envelope.Connectors == nil {
		return nil, fmt.Errorf("connector payload connectors must be an array")
	}
	if err := validateConnectorDefinitions(envelope.Connectors); err != nil {
		return nil, err
	}
	return envelope.Connectors, nil
}

func requireJSONEOF(decoder *json.Decoder) error {
	var extra any
	if err := decoder.Decode(&extra); err != io.EOF {
		if err == nil {
			return fmt.Errorf("multiple JSON values")
		}
		return err
	}
	return nil
}

func validateConnectorDefinitions(definitions []ConnectorDefinition) error {
	if len(definitions) > connectorDefinitionLimit {
		return fmt.Errorf("connector count %d exceeds %d", len(definitions), connectorDefinitionLimit)
	}
	ids := map[string]bool{}
	for index, definition := range definitions {
		if err := validateConnectorDefinition(definition); err != nil {
			return fmt.Errorf("connector %d: %w", index+1, err)
		}
		if ids[definition.ID] {
			return fmt.Errorf("duplicate connector id %q", definition.ID)
		}
		ids[definition.ID] = true
	}
	return nil
}

func validateConnectorDefinition(definition ConnectorDefinition) error {
	if !boundedText(definition.ID, 256) || !boundedText(definition.Name, 1024) || !boundedText(definition.Target.SheetID, 256) {
		return fmt.Errorf("id, name, and target sheetId must be nonempty and bounded")
	}
	if definition.Target.StartRow < 0 || definition.Target.StartRow > 1_048_575 || definition.Target.StartColumn < 0 || definition.Target.StartColumn > 16_383 {
		return fmt.Errorf("target coordinates exceed native XLSX bounds")
	}
	if definition.Source.Kind != "http" || (definition.Source.Format != "json" && definition.Source.Format != "csv") {
		return fmt.Errorf("source must be http with json or csv format")
	}
	if err := validateCredentialFreeConnectorURL(definition.Source.URL); err != nil {
		return err
	}
	if len(definition.Source.Path) > 1024 || strings.ContainsAny(definition.Source.Path, "\x00\r\n") {
		return fmt.Errorf("source path is invalid")
	}
	if definition.Refresh != "manual" && definition.Refresh != "onOpen" && definition.Refresh != "interval" {
		return fmt.Errorf("refresh policy %q is unsupported", definition.Refresh)
	}
	if definition.Refresh == "interval" {
		if definition.Schedule == nil || definition.Schedule.IntervalMS < 1000 || definition.Schedule.IntervalMS > 31_536_000_000 {
			return fmt.Errorf("interval refresh requires a schedule from 1000ms through one year")
		}
	} else if definition.Schedule != nil {
		return fmt.Errorf("schedule is only valid for interval refresh")
	}
	if definition.Cache != nil && (definition.Cache.Mode != "none" && definition.Cache.Mode != "memory" || definition.Cache.TTLMS < 0 || definition.Cache.TTLMS > 31_536_000_000) {
		return fmt.Errorf("cache policy is invalid")
	}
	if definition.Schema != nil {
		indexes := map[int]bool{}
		if definition.Schema.Columns == nil {
			return fmt.Errorf("schema columns must be an array")
		}
		if len(definition.Schema.Columns) > 16_384 {
			return fmt.Errorf("schema has too many columns")
		}
		for _, column := range definition.Schema.Columns {
			if column.Index < 0 || column.Index > 16_383 || indexes[column.Index] ||
				(column.Type != "string" && column.Type != "number" && column.Type != "boolean" && column.Type != "any") {
				return fmt.Errorf("schema columns require unique native indexes and supported types")
			}
			indexes[column.Index] = true
		}
	}
	return nil
}

func validateCredentialFreeConnectorURL(value string) error {
	if !boundedText(value, 8192) || strings.Contains(value, "\\") {
		return fmt.Errorf("source URL is empty, oversized, or contains a backslash")
	}
	parsed, err := url.Parse(value)
	if err != nil || parsed.User != nil || parsed.RawQuery != "" || parsed.ForceQuery || parsed.Fragment != "" {
		return fmt.Errorf("source URL must not contain credentials, query parameters, or a fragment")
	}
	if parsed.Host != "" && parsed.Scheme == "" {
		return fmt.Errorf("protocol-relative source URLs are unsupported")
	}
	if parsed.Scheme != "" && (parsed.Scheme != "http" && parsed.Scheme != "https" || parsed.Host == "") {
		return fmt.Errorf("absolute source URL must use http or https with a host")
	}
	return nil
}

func boundedText(value string, limit int) bool {
	return value != "" && len(value) <= limit && strings.TrimSpace(value) == value && !strings.ContainsAny(value, "\x00\r\n")
}
