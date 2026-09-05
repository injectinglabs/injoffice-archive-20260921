package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"net/url"
	"strings"
	"unicode/utf8"
)

const (
	packageRelationshipsNamespace     = "http://schemas.openxmlformats.org/package/2006/relationships"
	relTypeOfficeDocumentTransitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	relTypeOfficeDocumentStrict       = "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument"
)

type opcPackageIndex struct {
	byExact map[string]*zip.File
	byKey   map[string]*zip.File
}

type workbookPartLocation struct {
	part     string
	relsPart string
	baseDir  string
	strict   bool
}

type routingRelationship struct {
	id, relType, target, targetMode string
}

const (
	maxRoutingRelationshipIDLength     = 256
	maxRoutingRelationshipTypeLength   = 2_048
	maxRoutingRelationshipTargetLength = 8_192
)

func newOPCPackageIndex(reader *zip.Reader) (*opcPackageIndex, error) {
	index := &opcPackageIndex{
		byExact: make(map[string]*zip.File, len(reader.File)),
		byKey:   make(map[string]*zip.File, len(reader.File)),
	}
	for _, file := range reader.File {
		if prior, duplicate := index.byExact[file.Name]; duplicate {
			return nil, fmt.Errorf("duplicate entry %q (first %q)", file.Name, prior.Name)
		}
		index.byExact[file.Name] = file
		if strings.HasSuffix(file.Name, "/") {
			continue
		}
		key, err := canonicalOPCPartKey(file.Name)
		if err != nil {
			return nil, fmt.Errorf("non-conforming OPC part name %q: %w", file.Name, err)
		}
		if prior, equivalent := index.byKey[key]; equivalent {
			return nil, fmt.Errorf("case/escape-equivalent OPC parts %q and %q", prior.Name, file.Name)
		}
		index.byKey[key] = file
	}
	return index, nil
}

func (index *opcPackageIndex) lookupSpelling(name string) (string, *zip.File, bool) {
	key, err := canonicalOPCPartKey(name)
	if err != nil {
		return "", nil, false
	}
	file, ok := index.byKey[key]
	if !ok {
		return "", nil, false
	}
	return file.Name, file, true
}

func (index *opcPackageIndex) lookupResolved(decodedName string) (string, *zip.File, bool) {
	key, err := canonicalDecodedOPCPartKey(decodedName)
	if err != nil {
		return "", nil, false
	}
	file, ok := index.byKey[key]
	if !ok {
		return "", nil, false
	}
	return file.Name, file, true
}

func canonicalOPCPartKey(name string) (string, error) {
	if name == "" {
		return "", fmt.Errorf("empty part name")
	}
	reference, err := url.Parse(name)
	if err != nil {
		return "", fmt.Errorf("invalid part URI: %w", err)
	}
	if reference.IsAbs() || reference.Host != "" || reference.Opaque != "" || reference.RawQuery != "" || reference.ForceQuery || reference.Fragment != "" {
		return "", fmt.Errorf("part name is not a package path")
	}
	lowerEscaped := strings.ToLower(reference.EscapedPath())
	if strings.Contains(lowerEscaped, "%2f") || strings.Contains(lowerEscaped, "%5c") {
		return "", fmt.Errorf("part name contains an encoded path separator")
	}
	return canonicalDecodedOPCPartKey(reference.Path)
}

func canonicalDecodedOPCPartKey(name string) (string, error) {
	if !utf8.ValidString(name) {
		return "", fmt.Errorf("part name is not valid UTF-8")
	}
	if strings.HasPrefix(name, "/") {
		name = strings.TrimPrefix(name, "/")
	}
	if name == "" || strings.HasPrefix(name, "/") || strings.HasSuffix(name, "/") || strings.Contains(name, "//") {
		return "", fmt.Errorf("part name has non-canonical path separators")
	}
	if err := validateDecodedPackagePath(name); err != nil {
		return "", err
	}
	segments := strings.Split(name, "/")
	for _, segment := range segments {
		if segment == "" || segment == "." || segment == ".." {
			return "", fmt.Errorf("part name contains a dot or empty segment")
		}
		if strings.HasSuffix(segment, ".") {
			return "", fmt.Errorf("part-name segment %q ends with a dot", segment)
		}
	}
	return asciiLower(name), nil
}

// Native wire contracts use one package-relative spelling. Routing parsers may
// accept the leading slash required by [Content_Types].xml and normalize URI
// aliases for lookup, but those aliases never cross the contract boundary.
func canonicalNativeContractPartKey(name string) (string, error) {
	if strings.HasPrefix(name, "/") {
		return "", fmt.Errorf("contract part name must be package-relative")
	}
	for index := 0; index < len(name); index++ {
		if name[index] != '%' {
			continue
		}
		if index+2 >= len(name) || !isUpperHexDigit(name[index+1]) || !isUpperHexDigit(name[index+2]) {
			return "", fmt.Errorf("contract part name contains a non-canonical percent escape")
		}
		decoded := decodeUpperHex(name[index+1])<<4 | decodeUpperHex(name[index+2])
		if decoded >= 'a' && decoded <= 'z' || decoded >= 'A' && decoded <= 'Z' || decoded >= '0' && decoded <= '9' || strings.ContainsRune("-._~", rune(decoded)) {
			return "", fmt.Errorf("contract part name percent-encodes an unreserved character")
		}
		index += 2
	}
	return canonicalOPCPartKey(name)
}

func isUpperHexDigit(value byte) bool {
	return value >= '0' && value <= '9' || value >= 'A' && value <= 'F'
}
func decodeUpperHex(value byte) byte {
	if value <= '9' {
		return value - '0'
	}
	return value - 'A' + 10
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

func locateWorkbookPart(index *opcPackageIndex, read func(string) (string, bool)) (workbookPartLocation, error) {
	return locateWorkbookPartBytes(index, func(name string) ([]byte, bool) {
		value, ok := read(name)
		return []byte(value), ok
	})
}

func locateWorkbookPartBytes(index *opcPackageIndex, read func(string) ([]byte, bool)) (workbookPartLocation, error) {
	rootRelsPart, _, ok := index.lookupSpelling("_rels/.rels")
	if !ok {
		return workbookPartLocation{}, fmt.Errorf("missing root package relationships")
	}
	rootRelsXML, ok := read(rootRelsPart)
	if !ok {
		return workbookPartLocation{}, fmt.Errorf("unreadable root package relationships")
	}
	target, strict, err := officeDocumentRelationshipTarget(rootRelsXML)
	if err != nil {
		return workbookPartLocation{}, err
	}
	resolved, err := resolveRelPath("", target)
	if err != nil {
		return workbookPartLocation{}, fmt.Errorf("officeDocument relationship target %q: %w", target, err)
	}
	workbookPart, _, ok := index.lookupResolved(resolved)
	if !ok {
		return workbookPartLocation{}, fmt.Errorf("officeDocument relationship targets missing part %q", resolved)
	}
	workbookXML, ok := read(workbookPart)
	if !ok {
		return workbookPartLocation{}, fmt.Errorf("unreadable workbook part %q", workbookPart)
	}
	workbookStrict, err := workbookRootDialect(workbookXML)
	if err != nil {
		return workbookPartLocation{}, err
	}
	if workbookStrict != strict {
		return workbookPartLocation{}, fmt.Errorf("officeDocument relationship and workbook SpreadsheetML namespaces use opposing Strict/Transitional dialects")
	}
	relsCandidate := relsPartFor(workbookPart)
	workbookRelsPart, _, ok := index.lookupSpelling(relsCandidate)
	if !ok {
		return workbookPartLocation{}, fmt.Errorf("missing workbook relationships for %q", workbookPart)
	}
	return workbookPartLocation{part: workbookPart, relsPart: workbookRelsPart, baseDir: partBaseDir(resolved), strict: strict}, nil
}

func officeDocumentRelationshipTarget(data []byte) (string, bool, error) {
	relationships, err := parseRoutingRelationships(data)
	if err != nil {
		return "", false, fmt.Errorf("parse root package relationships: %w", err)
	}
	var match *routingRelationship
	for index := range relationships {
		if relationships[index].relType != relTypeOfficeDocumentTransitional && relationships[index].relType != relTypeOfficeDocumentStrict {
			continue
		}
		if match != nil {
			return "", false, fmt.Errorf("multiple officeDocument relationships")
		}
		match = &relationships[index]
	}
	if match == nil {
		return "", false, fmt.Errorf("root package has no officeDocument relationship")
	}
	switch match.targetMode {
	case "", "Internal":
	case "External":
		return "", false, fmt.Errorf("officeDocument relationship is external")
	default:
		return "", false, fmt.Errorf("officeDocument relationship has unsupported target mode %q", match.targetMode)
	}
	return match.target, match.relType == relTypeOfficeDocumentStrict, nil
}

func workbookRootDialect(data []byte) (bool, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	declarationSeen, prefixSeen := false, false
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			return false, fmt.Errorf("workbook part has no root element")
		}
		if err != nil {
			return false, fmt.Errorf("parse workbook root: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			if token.Name.Local != "workbook" || !isSpreadsheetMLNamespace(token.Name.Space) {
				return false, fmt.Errorf("root element is not a supported SpreadsheetML workbook")
			}
			return token.Name.Space == spreadsheetMLStrict, nil
		case xml.CharData:
			if len(strings.TrimSpace(string(token))) != 0 {
				return false, fmt.Errorf("workbook has text before its root")
			}
			if len(token) != 0 {
				prefixSeen = true
			}
		case xml.ProcInst:
			if declarationSeen || prefixSeen || !validNativeXMLDeclaration(token) {
				return false, fmt.Errorf("workbook has invalid or repeated XML declaration/processing instruction before its root")
			}
			declarationSeen, prefixSeen = true, true
		case xml.Comment:
			prefixSeen = true
		case xml.Directive:
			return false, fmt.Errorf("workbook has unsupported directive before its root")
		}
	}
}

func parseRoutingRelationships(data []byte) ([]routingRelationship, error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootSeen, rootClosed := false, false
	declarationSeen, prefixSeen := false, false
	relationships := make([]routingRelationship, 0)
	seenIDs := make(map[string]bool)
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			return nil, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if rootSeen || rootClosed || token.Name.Space != packageRelationshipsNamespace || token.Name.Local != "Relationships" {
					return nil, fmt.Errorf("root is not package Relationships")
				}
				if err := requireOnlySemanticXMLAttributes(token); err != nil {
					return nil, err
				}
				rootSeen = true
				continue
			}
			if depth == 2 {
				if token.Name.Space != packageRelationshipsNamespace || token.Name.Local != "Relationship" {
					return nil, fmt.Errorf("Relationships has unsupported direct child %q", token.Name.Local)
				}
				if err := requireOnlySemanticXMLAttributes(token,
					xml.Name{Local: "Id"}, xml.Name{Local: "Type"}, xml.Name{Local: "Target"}, xml.Name{Local: "TargetMode"},
				); err != nil {
					return nil, err
				}
				id, idFound, err := unqualifiedXMLAttribute(token, "Id")
				if err != nil {
					return nil, err
				}
				relType, typeFound, err := unqualifiedXMLAttribute(token, "Type")
				if err != nil {
					return nil, err
				}
				target, targetFound, err := unqualifiedXMLAttribute(token, "Target")
				if err != nil {
					return nil, err
				}
				targetMode, _, err := unqualifiedXMLAttribute(token, "TargetMode")
				if err != nil {
					return nil, err
				}
				if !idFound || id == "" || !typeFound || relType == "" || !targetFound || target == "" {
					return nil, fmt.Errorf("Relationship requires non-empty Id, Type, and Target")
				}
				if len(id) > maxRoutingRelationshipIDLength || len(relType) > maxRoutingRelationshipTypeLength || len(target) > maxRoutingRelationshipTargetLength || len(targetMode) > len("External") {
					return nil, fmt.Errorf("Relationship Id, Type, Target, or TargetMode exceeds its resource bound")
				}
				if targetMode != "" && targetMode != "Internal" && targetMode != "External" {
					return nil, fmt.Errorf("Relationship %q has unsupported TargetMode %q (unsupported target mode %q)", id, targetMode, targetMode)
				}
				if seenIDs[id] {
					return nil, fmt.Errorf("relationship %q is duplicated", id)
				}
				seenIDs[id] = true
				relationships = append(relationships, routingRelationship{id: id, relType: relType, target: target, targetMode: targetMode})
				if len(relationships) > NativeXLSXMaxRelationships {
					return nil, fmt.Errorf("Relationships exceeds %d entries", NativeXLSXMaxRelationships)
				}
				continue
			}
			return nil, fmt.Errorf("Relationship has unsupported nested markup")
		case xml.EndElement:
			if depth == 1 && token.Name.Space == packageRelationshipsNamespace && token.Name.Local == "Relationships" {
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if depth == 0 && len(strings.TrimSpace(string(token))) != 0 {
				return nil, fmt.Errorf("Relationships has text outside its root")
			}
			if depth > 0 && len(strings.TrimSpace(string(token))) != 0 {
				return nil, fmt.Errorf("Relationships has unsupported text content")
			}
			if depth == 0 && !rootSeen && len(token) != 0 {
				prefixSeen = true
			}
		case xml.ProcInst:
			if depth == 0 && !rootSeen && !declarationSeen && !prefixSeen && validNativeXMLDeclaration(token) {
				declarationSeen, prefixSeen = true, true
				continue
			}
			return nil, fmt.Errorf("Relationships has unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return nil, fmt.Errorf("Relationships has unsupported XML directive")
		case xml.Comment:
			if depth == 0 && !rootSeen {
				prefixSeen = true
			}
		}
	}
	if depth != 0 || !rootSeen || !rootClosed {
		return nil, fmt.Errorf("missing complete package Relationships root")
	}
	return relationships, nil
}

func unqualifiedXMLAttribute(start xml.StartElement, name string) (string, bool, error) {
	value, found := "", false
	for _, attribute := range start.Attr {
		if attribute.Name.Space != "" || attribute.Name.Local != name {
			continue
		}
		if found {
			return "", false, fmt.Errorf("duplicate unqualified attribute %q", name)
		}
		value, found = attribute.Value, true
	}
	return value, found, nil
}

func partBaseDir(part string) string {
	if index := strings.LastIndex(part, "/"); index >= 0 {
		return part[:index]
	}
	return ""
}
