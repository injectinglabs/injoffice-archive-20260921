package officecompat

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"encoding/xml"
	"hash"
	"io"
	"sort"
	"strconv"
	"strings"
)

// XMLStructuralStatus describes the bounded namespace-aware comparison result.
// Unavailable always falls back to the lexical result; it never means equal.
type XMLStructuralStatus string

const (
	XMLStructuralEqual       XMLStructuralStatus = "equal"
	XMLStructuralDifferent   XMLStructuralStatus = "different"
	XMLStructuralUnavailable XMLStructuralStatus = "unavailable"
)

// XMLFailureCode is a stable reason that a structural digest was unavailable.
type XMLFailureCode string

const (
	XMLFailureInvalidXML            XMLFailureCode = "invalid-xml"
	XMLFailureDirective             XMLFailureCode = "directive-or-doctype"
	XMLFailureProcessingInstruction XMLFailureCode = "processing-instruction"
	XMLFailureResourceLimit         XMLFailureCode = "resource-limit"
	XMLFailureRead                  XMLFailureCode = "read-failure"
)

// XMLComparison records both the authoritative byte-level comparison and a
// conservative structural comparison. Structural fingerprints normalize XML
// namespace prefixes, attribute order, quote style, entity spelling, comments,
// and empty-element spelling. Character data remains exact after XML decoding.
type XMLComparison struct {
	LexicalEqual            bool                `json:"lexical_equal"`
	BeforeSHA256            string              `json:"before_sha256"`
	AfterSHA256             string              `json:"after_sha256"`
	StructuralStatus        XMLStructuralStatus `json:"structural_status"`
	BeforeStructuralSHA256  string              `json:"before_structural_sha256,omitempty"`
	AfterStructuralSHA256   string              `json:"after_structural_sha256,omitempty"`
	BeforeStructuralFailure XMLFailureCode      `json:"before_structural_failure,omitempty"`
	AfterStructuralFailure  XMLFailureCode      `json:"after_structural_failure,omitempty"`
}

// CompareXML compares two XML payloads using DefaultLimits. Lexical equality
// remains authoritative even when structural equality is also reported.
func CompareXML(before, after []byte) XMLComparison {
	comparison, _ := CompareXMLWithLimits(before, after, DefaultLimits())
	return comparison
}

// CompareXMLWithLimits compares two XML payloads under explicit resource
// limits. Malformed or unsupported XML is a deterministic unavailable result;
// only an invalid Limits value is returned as an error.
func CompareXMLWithLimits(before, after []byte, limits Limits) (XMLComparison, error) {
	if err := limits.validate(); err != nil {
		return XMLComparison{}, err
	}
	beforeDigest := sha256.Sum256(before)
	afterDigest := sha256.Sum256(after)
	comparison := XMLComparison{
		LexicalEqual: bytes.Equal(before, after),
		BeforeSHA256: hex.EncodeToString(beforeDigest[:]),
		AfterSHA256:  hex.EncodeToString(afterDigest[:]),
	}
	if uint64(len(before)) > limits.MaxXMLBytes || uint64(len(after)) > limits.MaxXMLBytes {
		comparison.StructuralStatus = XMLStructuralUnavailable
		if uint64(len(before)) > limits.MaxXMLBytes {
			comparison.BeforeStructuralFailure = XMLFailureResourceLimit
		} else {
			beforeResult := fingerprintXML(func() (io.ReadCloser, error) {
				return io.NopCloser(bytes.NewReader(before)), nil
			}, uint64(len(before)), limits)
			comparison.BeforeStructuralSHA256 = beforeResult.digest
			comparison.BeforeStructuralFailure = beforeResult.failure
		}
		if uint64(len(after)) > limits.MaxXMLBytes {
			comparison.AfterStructuralFailure = XMLFailureResourceLimit
		} else {
			afterResult := fingerprintXML(func() (io.ReadCloser, error) {
				return io.NopCloser(bytes.NewReader(after)), nil
			}, uint64(len(after)), limits)
			comparison.AfterStructuralSHA256 = afterResult.digest
			comparison.AfterStructuralFailure = afterResult.failure
		}
		return comparison, nil
	}
	beforeResult := fingerprintXML(func() (io.ReadCloser, error) {
		return io.NopCloser(bytes.NewReader(before)), nil
	}, uint64(len(before)), limits)
	afterResult := fingerprintXML(func() (io.ReadCloser, error) {
		return io.NopCloser(bytes.NewReader(after)), nil
	}, uint64(len(after)), limits)
	finishXMLComparison(&comparison, beforeResult, afterResult)
	return comparison, nil
}

type xmlFingerprint struct {
	digest     string
	failure    XMLFailureCode
	tokens     uint64
	attributes int
}

func compareXMLFiles(before, after *zip.File, beforeSHA256, afterSHA256 string, limits Limits) XMLComparison {
	comparison := XMLComparison{
		LexicalEqual: false,
		BeforeSHA256: beforeSHA256,
		AfterSHA256:  afterSHA256,
	}
	beforeResult := fingerprintXML(before.Open, before.UncompressedSize64, limits)
	afterResult := fingerprintXML(after.Open, after.UncompressedSize64, limits)
	finishXMLComparison(&comparison, beforeResult, afterResult)
	return comparison
}

func finishXMLComparison(comparison *XMLComparison, before, after xmlFingerprint) {
	comparison.BeforeStructuralSHA256 = before.digest
	comparison.AfterStructuralSHA256 = after.digest
	comparison.BeforeStructuralFailure = before.failure
	comparison.AfterStructuralFailure = after.failure
	if before.failure != "" || after.failure != "" {
		comparison.StructuralStatus = XMLStructuralUnavailable
		return
	}
	if before.digest == after.digest {
		comparison.StructuralStatus = XMLStructuralEqual
		return
	}
	comparison.StructuralStatus = XMLStructuralDifferent
}

func fingerprintXML(open func() (io.ReadCloser, error), size uint64, limits Limits) xmlFingerprint {
	if size == 0 || size > limits.MaxXMLBytes {
		return xmlFingerprint{failure: XMLFailureResourceLimit}
	}
	rc, err := open()
	if err != nil {
		return xmlFingerprint{failure: XMLFailureRead}
	}
	defer rc.Close()

	payload, err := io.ReadAll(io.LimitReader(rc, int64(size)+1))
	if err != nil || uint64(len(payload)) != size {
		return xmlFingerprint{failure: XMLFailureRead}
	}
	if bytes.HasPrefix(payload, []byte{0xef, 0xbb, 0xbf}) {
		payload = payload[3:]
	}

	digest := sha256.New()
	decoder := xml.NewDecoder(bytes.NewReader(payload))
	decoder.Strict = true
	var inputOffset int64
	var tokens uint64
	var attributes int
	namespaces := map[string]string{"xml": "http://www.w3.org/XML/1998/namespace"}
	var stack []xmlElementFrame
	seenRoot := false
	completedRoot := false
	seenDeclaration := false
	declarationAllowed := true
	var pendingCharacterData []byte
	flushCharacterData := func() {
		if len(pendingCharacterData) == 0 {
			return
		}
		writeXMLMarker(digest, 'T')
		writeXMLBytes(digest, pendingCharacterData)
		pendingCharacterData = pendingCharacterData[:0]
	}
	for {
		tokenStart := inputOffset
		token, err := decoder.RawToken()
		inputOffset = decoder.InputOffset()
		if err == io.EOF {
			if !seenRoot || len(stack) != 0 {
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			return xmlFingerprint{digest: hex.EncodeToString(digest.Sum(nil)), tokens: tokens, attributes: attributes}
		}
		if err != nil {
			return xmlFingerprint{failure: XMLFailureInvalidXML}
		}
		tokens++
		if tokens > limits.MaxXMLTokens {
			return xmlFingerprint{failure: XMLFailureResourceLimit}
		}

		switch value := token.(type) {
		case xml.StartElement:
			if len(value.Attr) > limits.MaxXMLAttributes-attributes {
				return xmlFingerprint{failure: XMLFailureResourceLimit}
			}
			attributes += len(value.Attr)
			if tokenStart < 0 || inputOffset < tokenStart || inputOffset > int64(len(payload)) {
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			normalizedAttributes, ok := normalizeXMLStartAttributes(payload[tokenStart:inputOffset], value.Attr)
			if !ok {
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			value.Attr = normalizedAttributes
			if len(stack) != 0 {
				flushCharacterData()
			}
			declarationAllowed = false
			if len(stack) == 0 {
				if completedRoot {
					return xmlFingerprint{failure: XMLFailureInvalidXML}
				}
				seenRoot = true
			}
			if len(stack)+1 > limits.MaxXMLDepth {
				return xmlFingerprint{failure: XMLFailureResourceLimit}
			}
			frame := xmlElementFrame{rawName: value.Name}
			declared := map[string]bool{}
			for _, attribute := range value.Attr {
				prefix, namespaceDeclaration := xmlNamespaceDeclaration(attribute)
				if !namespaceDeclaration {
					continue
				}
				if declared[prefix] || !validXMLNamespaceBinding(prefix, attribute.Value) {
					return xmlFingerprint{failure: XMLFailureInvalidXML}
				}
				declared[prefix] = true
				previous, existed := namespaces[prefix]
				frame.namespaceChanges = append(frame.namespaceChanges, xmlNamespaceChange{prefix: prefix, previous: previous, existed: existed})
				if attribute.Value == "" {
					delete(namespaces, prefix)
				} else {
					namespaces[prefix] = attribute.Value
				}
			}
			resolvedName, ok := resolveXMLName(value.Name, namespaces, true)
			if !ok {
				restoreXMLNamespaces(namespaces, frame.namespaceChanges)
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			frame.resolvedName = resolvedName
			filtered := make([]xml.Attr, 0, len(value.Attr))
			for _, attribute := range value.Attr {
				if _, namespaceDeclaration := xmlNamespaceDeclaration(attribute); namespaceDeclaration {
					continue
				}
				resolvedAttribute, ok := resolveXMLName(attribute.Name, namespaces, false)
				if !ok {
					restoreXMLNamespaces(namespaces, frame.namespaceChanges)
					return xmlFingerprint{failure: XMLFailureInvalidXML}
				}
				attribute.Name = resolvedAttribute
				filtered = append(filtered, attribute)
			}
			sort.Slice(filtered, func(i, j int) bool {
				left, right := filtered[i], filtered[j]
				if left.Name.Space != right.Name.Space {
					return left.Name.Space < right.Name.Space
				}
				if left.Name.Local != right.Name.Local {
					return left.Name.Local < right.Name.Local
				}
				return left.Value < right.Value
			})
			for index := 1; index < len(filtered); index++ {
				if filtered[index-1].Name == filtered[index].Name {
					restoreXMLNamespaces(namespaces, frame.namespaceChanges)
					return xmlFingerprint{failure: XMLFailureInvalidXML}
				}
			}
			writeXMLMarker(digest, 'S')
			writeXMLString(digest, resolvedName.Space)
			writeXMLString(digest, resolvedName.Local)
			writeXMLUint(digest, uint64(len(filtered)))
			for _, attribute := range filtered {
				writeXMLString(digest, attribute.Name.Space)
				writeXMLString(digest, attribute.Name.Local)
				writeXMLString(digest, attribute.Value)
			}
			stack = append(stack, frame)
		case xml.EndElement:
			if len(stack) == 0 {
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			flushCharacterData()
			frame := stack[len(stack)-1]
			resolvedName, ok := resolveXMLName(value.Name, namespaces, true)
			if !ok || value.Name != frame.rawName || resolvedName != frame.resolvedName {
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			writeXMLMarker(digest, 'E')
			writeXMLString(digest, resolvedName.Space)
			writeXMLString(digest, resolvedName.Local)
			restoreXMLNamespaces(namespaces, frame.namespaceChanges)
			stack = stack[:len(stack)-1]
			if len(stack) == 0 {
				completedRoot = true
			}
		case xml.CharData:
			if len(stack) == 0 {
				if !onlyXMLSpace(value) {
					return xmlFingerprint{failure: XMLFailureInvalidXML}
				}
				if len(value) != 0 {
					declarationAllowed = false
				}
				continue
			}
			if uint64(len(pendingCharacterData)) > limits.MaxXMLBytes-uint64(len(value)) {
				return xmlFingerprint{failure: XMLFailureResourceLimit}
			}
			pendingCharacterData = append(pendingCharacterData, value...)
		case xml.Comment:
			// Comments have no OOXML document semantics. Their lexical bytes are
			// still covered by the authoritative payload digest.
			if !seenRoot {
				declarationAllowed = false
			}
		case xml.ProcInst:
			if value.Target != "xml" || !declarationAllowed || seenRoot || seenDeclaration {
				return xmlFingerprint{failure: XMLFailureProcessingInstruction}
			}
			if !validXMLDeclaration(value.Inst) {
				return xmlFingerprint{failure: XMLFailureInvalidXML}
			}
			seenDeclaration = true
			declarationAllowed = false
		case xml.Directive:
			return xmlFingerprint{failure: XMLFailureDirective}
		}
	}
}

// normalizeXMLStartAttributes applies XML 1.0 end-of-line and attribute-value
// normalization to literal attribute whitespace. Numeric character references
// deliberately retain their referenced tab/CR/LF value; encoding/xml exposes
// both forms identically, so this bounded lexical pass preserves the distinction.
func normalizeXMLStartAttributes(raw []byte, decoded []xml.Attr) ([]xml.Attr, bool) {
	if len(raw) < 3 || raw[0] != '<' {
		return nil, false
	}
	index := 1
	for index < len(raw) && !isXMLSpace(raw[index]) && raw[index] != '/' && raw[index] != '>' {
		index++
	}
	if index == 1 {
		return nil, false
	}

	normalized := append([]xml.Attr(nil), decoded...)
	attributeIndex := 0
	for {
		index = skipXMLSpace(raw, index)
		if index >= len(raw) {
			return nil, false
		}
		if raw[index] == '>' {
			index++
			break
		}
		if raw[index] == '/' && index+1 < len(raw) && raw[index+1] == '>' {
			index += 2
			break
		}
		if attributeIndex >= len(normalized) {
			return nil, false
		}
		nameStart := index
		for index < len(raw) && !isXMLSpace(raw[index]) && raw[index] != '=' && raw[index] != '/' && raw[index] != '>' {
			index++
		}
		if index == nameStart {
			return nil, false
		}
		index = skipXMLSpace(raw, index)
		if index >= len(raw) || raw[index] != '=' {
			return nil, false
		}
		index = skipXMLSpace(raw, index+1)
		if index >= len(raw) || raw[index] != '\'' && raw[index] != '"' {
			return nil, false
		}
		quote := raw[index]
		index++
		valueStart := index
		for index < len(raw) && raw[index] != quote {
			index++
		}
		if index == len(raw) {
			return nil, false
		}
		value, ok := normalizeXMLAttributeValue(raw[valueStart:index])
		if !ok {
			return nil, false
		}
		normalized[attributeIndex].Value = value
		attributeIndex++
		index++
	}
	return normalized, attributeIndex == len(normalized) && index == len(raw)
}

func normalizeXMLAttributeValue(raw []byte) (string, bool) {
	var normalized strings.Builder
	normalized.Grow(len(raw))
	for index := 0; index < len(raw); {
		switch raw[index] {
		case '\t', '\n':
			normalized.WriteByte(' ')
			index++
		case '\r':
			normalized.WriteByte(' ')
			index++
			if index < len(raw) && raw[index] == '\n' {
				index++
			}
		case '&':
			end := bytes.IndexByte(raw[index+1:], ';')
			if end < 0 {
				return "", false
			}
			end += index + 1
			character, ok := decodeXMLCharacterReference(raw[index+1 : end])
			if !ok {
				return "", false
			}
			normalized.WriteRune(character)
			index = end + 1
		default:
			normalized.WriteByte(raw[index])
			index++
		}
	}
	return normalized.String(), true
}

func decodeXMLCharacterReference(reference []byte) (rune, bool) {
	predefined := map[string]rune{"amp": '&', "lt": '<', "gt": '>', "apos": '\'', "quot": '"'}
	if len(reference) == 0 {
		return 0, false
	}
	if reference[0] != '#' {
		character, ok := predefined[string(reference)]
		return character, ok
	}
	base := 10
	digits := reference[1:]
	if len(digits) > 0 && digits[0] == 'x' {
		base = 16
		digits = digits[1:]
	}
	if len(digits) == 0 {
		return 0, false
	}
	value, err := strconv.ParseUint(string(digits), base, 32)
	character := rune(value)
	if err != nil || !validXMLCharacter(character) {
		return 0, false
	}
	return character, true
}

func validXMLCharacter(character rune) bool {
	return character == '\t' || character == '\n' || character == '\r' ||
		character >= 0x20 && character <= 0xd7ff ||
		character >= 0xe000 && character <= 0xfffd ||
		character >= 0x10000 && character <= 0x10ffff
}

func skipXMLSpace(value []byte, index int) int {
	for index < len(value) && isXMLSpace(value[index]) {
		index++
	}
	return index
}

func isXMLSpace(value byte) bool {
	return value == ' ' || value == '\t' || value == '\r' || value == '\n'
}

func onlyXMLSpace(value []byte) bool {
	for _, character := range value {
		if !isXMLSpace(character) {
			return false
		}
	}
	return true
}

func validXMLDeclaration(instruction []byte) bool {
	index := skipXMLDeclarationSpace(instruction, 0)
	version, next, ok := parseXMLDeclarationAttribute(instruction, index, "version")
	if !ok || version != "1.0" {
		return false
	}
	index = next
	if index == len(instruction) {
		return true
	}
	spaced := skipXMLDeclarationSpace(instruction, index)
	if spaced == index {
		return false
	}
	if spaced == len(instruction) {
		return true
	}
	index = spaced

	if bytes.HasPrefix(instruction[index:], []byte("encoding")) {
		encoding, next, ok := parseXMLDeclarationAttribute(instruction, index, "encoding")
		if !ok || !strings.EqualFold(encoding, "UTF-8") {
			return false
		}
		index = next
		if index == len(instruction) {
			return true
		}
		spaced = skipXMLDeclarationSpace(instruction, index)
		if spaced == index {
			return false
		}
		if spaced == len(instruction) {
			return true
		}
		index = spaced
	}

	standalone, next, ok := parseXMLDeclarationAttribute(instruction, index, "standalone")
	if !ok || (standalone != "yes" && standalone != "no") {
		return false
	}
	return skipXMLDeclarationSpace(instruction, next) == len(instruction)
}

func parseXMLDeclarationAttribute(instruction []byte, index int, name string) (string, int, bool) {
	if !bytes.HasPrefix(instruction[index:], []byte(name)) {
		return "", index, false
	}
	index += len(name)
	index = skipXMLDeclarationSpace(instruction, index)
	if index >= len(instruction) || instruction[index] != '=' {
		return "", index, false
	}
	index++
	index = skipXMLDeclarationSpace(instruction, index)
	if index >= len(instruction) || (instruction[index] != '\'' && instruction[index] != '"') {
		return "", index, false
	}
	quote := instruction[index]
	index++
	start := index
	for index < len(instruction) && instruction[index] != quote {
		index++
	}
	if index == len(instruction) {
		return "", index, false
	}
	value := string(instruction[start:index])
	return value, index + 1, true
}

func skipXMLDeclarationSpace(instruction []byte, index int) int {
	for index < len(instruction) {
		switch instruction[index] {
		case ' ', '\t', '\r', '\n':
			index++
		default:
			return index
		}
	}
	return index
}

type xmlNamespaceChange struct {
	prefix   string
	previous string
	existed  bool
}

type xmlElementFrame struct {
	rawName          xml.Name
	resolvedName     xml.Name
	namespaceChanges []xmlNamespaceChange
}

func xmlNamespaceDeclaration(attribute xml.Attr) (string, bool) {
	if attribute.Name.Space == "" && attribute.Name.Local == "xmlns" {
		return "", true
	}
	if attribute.Name.Space == "xmlns" {
		return attribute.Name.Local, true
	}
	return "", false
}

func validXMLNamespaceBinding(prefix, namespace string) bool {
	const (
		xmlNamespace   = "http://www.w3.org/XML/1998/namespace"
		xmlnsNamespace = "http://www.w3.org/2000/xmlns/"
	)
	if prefix == "xmlns" || namespace == xmlnsNamespace {
		return false
	}
	if prefix == "xml" {
		return namespace == xmlNamespace
	}
	if namespace == xmlNamespace {
		return false
	}
	return prefix == "" || namespace != ""
}

func resolveXMLName(name xml.Name, namespaces map[string]string, element bool) (xml.Name, bool) {
	if name.Space == "" {
		if element {
			return xml.Name{Space: namespaces[""], Local: name.Local}, true
		}
		return name, true
	}
	namespace, exists := namespaces[name.Space]
	if !exists || namespace == "" {
		return xml.Name{}, false
	}
	return xml.Name{Space: namespace, Local: name.Local}, true
}

func restoreXMLNamespaces(namespaces map[string]string, changes []xmlNamespaceChange) {
	for index := len(changes) - 1; index >= 0; index-- {
		change := changes[index]
		if change.existed {
			namespaces[change.prefix] = change.previous
		} else {
			delete(namespaces, change.prefix)
		}
	}
}

func writeXMLMarker(digest hash.Hash, marker byte) {
	_, _ = digest.Write([]byte{marker})
}

func writeXMLString(digest hash.Hash, value string) {
	writeXMLBytes(digest, []byte(value))
}

func writeXMLBytes(digest hash.Hash, value []byte) {
	writeXMLUint(digest, uint64(len(value)))
	_, _ = digest.Write(value)
}

func writeXMLUint(digest hash.Hash, value uint64) {
	var encoded [8]byte
	binary.BigEndian.PutUint64(encoded[:], value)
	_, _ = digest.Write(encoded[:])
}

func isXMLPart(name string) bool {
	lower := strings.ToLower(name)
	return strings.HasSuffix(lower, ".xml") || strings.HasSuffix(lower, ".rels")
}
