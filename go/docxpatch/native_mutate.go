package docxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"reflect"
	"sort"
	"strings"
	"unicode/utf8"
)

const (
	NativeDOCXMaxMutations            = 10_000
	NativeDOCXMaxMutationPayloadBytes = 8 * 1024 * 1024
)

// NativeDOCXTextMutationV1 replaces text at an exact native paragraph or text
// run anchor. Paragraph replacement is intentionally narrow: the paragraph
// must contain exactly one modeled text node. ExpectedXMLSHA256 is mandatory
// and prevents a target from being applied after its anchored XML changed.
type NativeDOCXTextMutationV1 struct {
	TargetKind        string `json:"target_kind"`
	TargetID          string `json:"target_id"`
	ExpectedXMLSHA256 string `json:"expected_xml_sha256"`
	Text              string `json:"text"`
}

// DecodeNativeDOCXTextMutationPayloadV1 strictly decodes the format-specific
// payload carried by the shared Office mutation envelope. The one accepted
// shape is {"mutations":[...]}; duplicate or unknown fields, non-string
// scalars, trailing JSON, and oversized batches are refused before mutation.
func DecodeNativeDOCXTextMutationPayloadV1(data []byte) ([]NativeDOCXTextMutationV1, error) {
	invalid := func(message string) ([]NativeDOCXTextMutationV1, error) {
		return nil, nativeMutationError("INVALID_PAYLOAD", "", message)
	}
	if len(data) == 0 || len(data) > NativeDOCXMaxMutationPayloadBytes {
		return invalid(fmt.Sprintf("encoded payload size must be 1..%d bytes", NativeDOCXMaxMutationPayloadBytes))
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	first, err := decoder.Token()
	if err != nil {
		return invalid("payload must be one JSON object")
	}
	if delimiter, ok := first.(json.Delim); !ok || delimiter != '{' {
		return invalid("payload must be one JSON object")
	}
	seenMutations := false
	var mutations []NativeDOCXTextMutationV1
	for decoder.More() {
		fieldToken, tokenErr := decoder.Token()
		if tokenErr != nil {
			return invalid("payload contains invalid JSON")
		}
		field, ok := fieldToken.(string)
		if !ok || field != "mutations" {
			return invalid(fmt.Sprintf("unknown payload field %q", field))
		}
		if seenMutations {
			return invalid("duplicate payload field \"mutations\"")
		}
		seenMutations = true
		arrayToken, arrayErr := decoder.Token()
		if arrayErr != nil {
			return invalid("mutations must be an array")
		}
		if delimiter, ok := arrayToken.(json.Delim); !ok || delimiter != '[' {
			return invalid("mutations must be an array")
		}
		for decoder.More() {
			if len(mutations) >= NativeDOCXMaxMutations {
				return invalid(fmt.Sprintf("mutation count must be 1..%d", NativeDOCXMaxMutations))
			}
			mutation, mutationErr := decodeNativeDOCXTextMutationV1(decoder, len(mutations))
			if mutationErr != nil {
				return nil, mutationErr
			}
			if selectorErr := validateNativeDOCXTextMutationSelector(mutation); selectorErr != nil {
				return invalid(fmt.Sprintf("mutation %d: %s", len(mutations), selectorErr.Error()))
			}
			mutations = append(mutations, mutation)
		}
		if _, closeErr := decoder.Token(); closeErr != nil {
			return invalid("mutations array is unterminated")
		}
	}
	if _, closeErr := decoder.Token(); closeErr != nil {
		return invalid("payload object is unterminated")
	}
	if token, trailingErr := decoder.Token(); trailingErr != io.EOF || token != nil {
		return invalid("payload must not contain trailing JSON")
	}
	if !seenMutations || len(mutations) == 0 {
		return invalid(fmt.Sprintf("mutation count must be 1..%d", NativeDOCXMaxMutations))
	}
	return mutations, nil
}

// ApplyNativeTextMutationPayloadV1 is the renderer-neutral Office mutation
// envelope entrypoint. It proves the envelope's full exact-byte CAS before
// decoding the format payload, then applies the decoded text replacements as
// one transaction. A refusal at either boundary returns no candidate package.
func ApplyNativeTextMutationPayloadV1(packageBytes, payload []byte, outerExpectedRevision string) (*NativeDOCXMutationResultV1, error) {
	exactRevision := nativeSHA(packageBytes)
	if outerExpectedRevision != exactRevision {
		return nil, nativeMutationError("STALE_REVISION", "", fmt.Sprintf("outer expected revision %q does not match exact package revision %q", outerExpectedRevision, exactRevision))
	}
	mutations, err := DecodeNativeDOCXTextMutationPayloadV1(payload)
	if err != nil {
		return nil, err
	}
	return ApplyNativeTextMutationsV1(packageBytes, outerExpectedRevision, mutations)
}

func decodeNativeDOCXTextMutationV1(decoder *json.Decoder, index int) (NativeDOCXTextMutationV1, error) {
	invalid := func(message string) (NativeDOCXTextMutationV1, error) {
		return NativeDOCXTextMutationV1{}, nativeMutationError("INVALID_PAYLOAD", "", fmt.Sprintf("mutation %d: %s", index, message))
	}
	first, err := decoder.Token()
	if err != nil {
		return invalid("must be an object")
	}
	if delimiter, ok := first.(json.Delim); !ok || delimiter != '{' {
		return invalid("must be an object")
	}
	seen := map[string]bool{}
	values := map[string]string{}
	allowed := map[string]bool{"target_kind": true, "target_id": true, "expected_xml_sha256": true, "text": true}
	for decoder.More() {
		fieldToken, tokenErr := decoder.Token()
		if tokenErr != nil {
			return invalid("contains invalid JSON")
		}
		field, ok := fieldToken.(string)
		if !ok || !allowed[field] {
			return invalid(fmt.Sprintf("unknown field %q", field))
		}
		if seen[field] {
			return invalid(fmt.Sprintf("duplicate field %q", field))
		}
		seen[field] = true
		var rawValue json.RawMessage
		if valueErr := decoder.Decode(&rawValue); valueErr != nil {
			return invalid(fmt.Sprintf("field %q contains invalid JSON", field))
		}
		value, valueErr := decodeNativeMutationJSONString(rawValue)
		if valueErr != nil {
			return invalid(fmt.Sprintf("field %q must be a string", field))
		}
		values[field] = value
	}
	if _, closeErr := decoder.Token(); closeErr != nil {
		return invalid("object is unterminated")
	}
	for _, field := range []string{"target_kind", "target_id", "expected_xml_sha256", "text"} {
		if !seen[field] {
			return invalid(fmt.Sprintf("field %q is required", field))
		}
	}
	return NativeDOCXTextMutationV1{
		TargetKind: values["target_kind"], TargetID: values["target_id"],
		ExpectedXMLSHA256: values["expected_xml_sha256"], Text: values["text"],
	}, nil
}

// encoding/json deliberately repairs malformed UTF-8 and unpaired UTF-16
// escapes with U+FFFD. Native mutation text is exact authority, so inspect the
// original JSON string token and refuse either form before decoding it.
func decodeNativeMutationJSONString(raw []byte) (string, error) {
	if len(raw) < 2 || raw[0] != '"' || raw[len(raw)-1] != '"' || !utf8.Valid(raw) {
		return "", fmt.Errorf("must be one valid UTF-8 JSON string")
	}
	for index := 1; index < len(raw)-1; index++ {
		if raw[index] != '\\' {
			continue
		}
		index++
		if index >= len(raw)-1 || raw[index] != 'u' {
			continue
		}
		unit, ok := nativeMutationHexCodeUnit(raw[index+1:])
		if !ok {
			return "", fmt.Errorf("contains an invalid Unicode escape")
		}
		index += 4
		if unit >= 0xD800 && unit <= 0xDBFF {
			if index+6 >= len(raw) || raw[index+1] != '\\' || raw[index+2] != 'u' {
				return "", fmt.Errorf("contains an unpaired UTF-16 high surrogate")
			}
			low, lowOK := nativeMutationHexCodeUnit(raw[index+3:])
			if !lowOK || low < 0xDC00 || low > 0xDFFF {
				return "", fmt.Errorf("contains an unpaired UTF-16 high surrogate")
			}
			index += 6
		} else if unit >= 0xDC00 && unit <= 0xDFFF {
			return "", fmt.Errorf("contains an unpaired UTF-16 low surrogate")
		}
	}
	var value string
	if err := json.Unmarshal(raw, &value); err != nil || !utf8.ValidString(value) {
		return "", fmt.Errorf("must be one valid UTF-8 JSON string")
	}
	return value, nil
}

func nativeMutationHexCodeUnit(raw []byte) (uint16, bool) {
	if len(raw) < 4 {
		return 0, false
	}
	var value uint16
	for _, character := range raw[:4] {
		value <<= 4
		switch {
		case character >= '0' && character <= '9':
			value += uint16(character - '0')
		case character >= 'a' && character <= 'f':
			value += uint16(character-'a') + 10
		case character >= 'A' && character <= 'F':
			value += uint16(character-'A') + 10
		default:
			return 0, false
		}
	}
	return value, true
}

// NativeDOCXChangedPartV1 records the content hashes proved by a successful
// mutation. Untouched parts are independently compared by ApplyPatch.
type NativeDOCXChangedPartV1 struct {
	PartName     string
	BeforeSHA256 string
	AfterSHA256  string
}

// NativeDOCXMutationEvidenceV1 is preservation evidence for a completed
// mutation transaction.
type NativeDOCXMutationEvidenceV1 struct {
	SourceRevision         string
	ResultRevision         string
	ChangedParts           []NativeDOCXChangedPartV1
	UntouchedPartsVerified int
}

// NativeDOCXMutationResultV1 returns the reopened and re-extracted native
// document together with preservation evidence. Package is authoritative.
type NativeDOCXMutationResultV1 struct {
	Package  []byte
	Document *NativeDocumentV1
	Evidence NativeDOCXMutationEvidenceV1
}

// NativeDOCXMutationErrorV1 exposes stable fail-closed refusal codes.
type NativeDOCXMutationErrorV1 struct {
	Code     string
	TargetID string
	Message  string
}

func (e *NativeDOCXMutationErrorV1) Error() string {
	if e.TargetID == "" {
		return fmt.Sprintf("docxpatch: native mutation %s: %s", e.Code, e.Message)
	}
	return fmt.Sprintf("docxpatch: native mutation %s for %q: %s", e.Code, e.TargetID, e.Message)
}

type nativeTextTarget struct {
	kind      string
	partName  string
	path      string
	anchor    NativeSourceAnchorV1
	splice    NativeSourceAnchorV1
	text      string
	paragraph *NativeParagraphV1
}

type nativeTextSplice struct {
	start int64
	end   int64
	text  []byte
}

// ApplyNativeTextMutationsV1 atomically applies renderer-neutral text changes
// to native WordprocessingML bytes. expectedRevision must be the full exact
// package digest from NativeDocumentV1.Source.PackageSHA256 (sha256 plus all 64
// lowercase hex digits); the shorter opaque Document.Revision is never save
// authority. The function verifies the package and every target anchor,
// performs only exact text-node splices, reopens the OPC, re-extracts the
// native contract with identity continuity, and validates the requested text
// at the same part-qualified XML paths. Any refusal returns no output package.
func ApplyNativeTextMutationsV1(packageBytes []byte, expectedRevision string, mutations []NativeDOCXTextMutationV1) (*NativeDOCXMutationResultV1, error) {
	if len(mutations) == 0 || len(mutations) > NativeDOCXMaxMutations {
		return nil, nativeMutationError("INVALID_MUTATION_COUNT", "", fmt.Sprintf("mutation count must be 1..%d", NativeDOCXMaxMutations))
	}
	if !nativeSHA256.MatchString(expectedRevision) {
		return nil, nativeMutationError("INVALID_REVISION", "", "expected revision must be sha256 followed by the full lowercase exact-byte digest")
	}
	doc, err := ExtractNativeDocumentV1(packageBytes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native mutation source validation: %w", err)
	}
	exactRevision := nativeSHA(packageBytes)
	if doc.Source.PackageSHA256 != exactRevision {
		return nil, nativeMutationError("SOURCE_MISMATCH", "", "extracted source fingerprint does not match the authoritative package bytes")
	}
	if expectedRevision == "" || expectedRevision != exactRevision {
		return nil, nativeMutationError("STALE_REVISION", "", fmt.Sprintf("expected %q, current exact package revision is %q", expectedRevision, exactRevision))
	}
	pkg, err := openNativeDOCXPackage(packageBytes)
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native mutation reopen source: %w", err)
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return nil, nativeMutationError("SIGNED_PACKAGE", "", "a signed OPC package cannot be mutated without invalidating its digital signature")
	}
	targets := indexNativeTextTargets(doc)
	splicesByPart := map[string][]nativeTextSplice{}
	expectedByPath := map[string]string{}
	seenTarget := map[string]bool{}
	requestedBytes := int64(0)
	for _, mutation := range mutations {
		if mutation.TargetKind != "paragraph" && mutation.TargetKind != "run" {
			return nil, nativeMutationError("UNSUPPORTED_TARGET_KIND", mutation.TargetID, "target_kind must be paragraph or run")
		}
		if selectorErr := validateNativeDOCXTextMutationSelector(mutation); selectorErr != nil {
			return nil, nativeMutationError("INVALID_SELECTOR", mutation.TargetID, selectorErr.Error())
		}
		if !utf8.ValidString(mutation.Text) || nativeMutationUTF16CodeUnits(mutation.Text) > NativeDOCXMaxTextLength || !nativeMutationXMLTextValid(mutation.Text) {
			return nil, nativeMutationError("INVALID_TEXT", mutation.TargetID, fmt.Sprintf("replacement must be valid XML text of at most %d UTF-16 code units", NativeDOCXMaxTextLength))
		}
		requestedBytes += int64(len(mutation.Text))
		if requestedBytes > NativeDOCXMaxUncompressedBytes {
			return nil, nativeMutationError("RESOURCE_LIMIT", mutation.TargetID, fmt.Sprintf("combined replacement text exceeds %d bytes", NativeDOCXMaxUncompressedBytes))
		}
		key := mutation.TargetKind + "\x00" + mutation.TargetID
		if seenTarget[key] {
			return nil, nativeMutationError("DUPLICATE_TARGET", mutation.TargetID, "a target may appear only once in an atomic transaction")
		}
		seenTarget[key] = true
		target, ok := targets[key]
		if !ok {
			return nil, nativeMutationError("TARGET_NOT_FOUND", mutation.TargetID, "target is absent from the current native contract")
		}
		if target.paragraph.EditPolicy.Mode != "read-write" || !nativePolicyAllows(target.paragraph.EditPolicy, "text.replace") {
			message := "owning paragraph is read-only"
			if target.paragraph.EditPolicy.Refusal != nil {
				message += ": " + target.paragraph.EditPolicy.Refusal.Code
			}
			return nil, nativeMutationError("UNSUPPORTED_CONSTRUCT", mutation.TargetID, message)
		}
		if mutation.ExpectedXMLSHA256 == "" || mutation.ExpectedXMLSHA256 != target.anchor.XMLSHA256 {
			return nil, nativeMutationError("STALE_TARGET", mutation.TargetID, fmt.Sprintf("expected XML fingerprint %q, current fingerprint is %q", mutation.ExpectedXMLSHA256, target.anchor.XMLSHA256))
		}
		part := pkg.files[target.partName]
		if target.anchor.StartByte == nil || target.anchor.EndByte == nil || *target.anchor.StartByte < 0 || *target.anchor.EndByte > int64(len(part)) {
			return nil, nativeMutationError("INVALID_ANCHOR", mutation.TargetID, "target anchor is outside its owning part")
		}
		guarded := part[*target.anchor.StartByte:*target.anchor.EndByte]
		if nativeSHA(guarded) != target.anchor.XMLSHA256 {
			return nil, nativeMutationError("STALE_TARGET", mutation.TargetID, "source bytes no longer match the issued anchor fingerprint")
		}
		if mutation.Text == target.text {
			return nil, nativeMutationError("SEMANTIC_NO_OP", mutation.TargetID, "replacement text is identical to the source-authoritative native text")
		}
		if target.splice.StartByte == nil || target.splice.EndByte == nil || *target.splice.StartByte < *target.anchor.StartByte || *target.splice.EndByte > *target.anchor.EndByte {
			return nil, nativeMutationError("INVALID_ANCHOR", mutation.TargetID, "text splice anchor is outside its guarded target")
		}
		raw := part[*target.splice.StartByte:*target.splice.EndByte]
		if nativeSHA(raw) != target.splice.XMLSHA256 {
			return nil, nativeMutationError("STALE_TARGET", mutation.TargetID, "text splice bytes no longer match the issued run anchor")
		}
		replacement, replaceErr := replaceNativeTextElement(raw, mutation.Text)
		if replaceErr != nil {
			return nil, nativeMutationError("UNSUPPORTED_LEXICAL_FORM", mutation.TargetID, replaceErr.Error())
		}
		splicesByPart[target.partName] = append(splicesByPart[target.partName], nativeTextSplice{start: *target.splice.StartByte, end: *target.splice.EndByte, text: replacement})
		expectedByPath[nativeMutationPathKey(target.kind, target.partName, target.path)] = mutation.Text
	}

	replacements := map[string][]byte{}
	changedNames := make([]string, 0, len(splicesByPart))
	projectedTotal := int64(0)
	for _, part := range pkg.files {
		projectedTotal += int64(len(part))
	}
	for partName, splices := range splicesByPart {
		part, spliceErr := applyNativeTextSplices(pkg.files[partName], splices)
		if spliceErr != nil {
			return nil, spliceErr
		}
		if len(part) > NativeDOCXMaxXMLPartBytes {
			return nil, nativeMutationError("RESOURCE_LIMIT", "", fmt.Sprintf("mutated XML part %q exceeds %d bytes", partName, NativeDOCXMaxXMLPartBytes))
		}
		if !bytes.Equal(part, pkg.files[partName]) {
			projectedTotal += int64(len(part) - len(pkg.files[partName]))
			if projectedTotal > NativeDOCXMaxUncompressedBytes {
				return nil, nativeMutationError("RESOURCE_LIMIT", "", fmt.Sprintf("mutated package exceeds %d uncompressed bytes", NativeDOCXMaxUncompressedBytes))
			}
			replacements[partName] = part
			changedNames = append(changedNames, partName)
		}
	}
	if len(replacements) == 0 {
		return nil, nativeMutationError("SEMANTIC_NO_OP", "", "mutation batch produced no native part payload change")
	}
	produced, err := ApplyPatch(packageBytes, Patch{Replace: replacements})
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native mutation write verification: %w", err)
	}
	// Preserve the full native identity graph while reissuing every anchor from
	// the produced bytes. This is the same explicit continuity boundary used by
	// incremental native extraction; it prevents an enclosing story or section
	// identity from churning merely because a descendant text node changed.
	after, err := ExtractNativeDocumentV1WithOptions(produced, NativeExtractionOptions{Previous: doc, RetainPathIdentity: true})
	if err != nil {
		return nil, fmt.Errorf("docxpatch: native mutation post-write extraction: %w", err)
	}
	if issues := ValidateNativeDocumentV1(after); len(issues) > 0 {
		return nil, fmt.Errorf("docxpatch: native mutation post-write validation: %w", &NativeValidationError{Issues: issues})
	}
	if err := validateNativeMutationResults(doc, after, expectedByPath); err != nil {
		return nil, err
	}
	sort.Strings(changedNames)
	changed := make([]NativeDOCXChangedPartV1, 0, len(changedNames))
	for _, name := range changedNames {
		changed = append(changed, NativeDOCXChangedPartV1{PartName: name, BeforeSHA256: nativeSHA(pkg.files[name]), AfterSHA256: nativeSHA(replacements[name])})
	}
	resultRevision := nativeSHA(produced)
	if after.Source.PackageSHA256 != resultRevision {
		return nil, nativeMutationError("POST_WRITE_MISMATCH", "", "reopened source fingerprint does not match the produced package bytes")
	}
	return &NativeDOCXMutationResultV1{Package: produced, Document: after, Evidence: NativeDOCXMutationEvidenceV1{SourceRevision: exactRevision, ResultRevision: resultRevision, ChangedParts: changed, UntouchedPartsVerified: len(pkg.files) - len(changed)}}, nil
}

func validateNativeDOCXTextMutationSelector(mutation NativeDOCXTextMutationV1) error {
	if !nativeIDPattern.MatchString(mutation.TargetID) {
		return fmt.Errorf("target_id must be a bounded native identifier")
	}
	if !nativeSHA256.MatchString(mutation.ExpectedXMLSHA256) {
		return fmt.Errorf("expected_xml_sha256 must contain a full lowercase SHA-256")
	}
	return nil
}

func nativeDOCXPackageHasDigitalSignature(pkg *nativePackage) bool {
	const signatureRelationshipPrefix = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/"
	for _, relationships := range pkg.rels {
		for _, relationship := range relationships {
			if strings.HasPrefix(relationship.Type, signatureRelationshipPrefix) {
				return true
			}
		}
	}
	for name := range pkg.files {
		key, err := nativeDecodedPartKey(name)
		if err == nil && (key == "_xmlsignatures" || strings.HasPrefix(key, "_xmlsignatures/")) {
			return true
		}
	}
	return false
}

func nativeMutationError(code, targetID, message string) error {
	return &NativeDOCXMutationErrorV1{Code: code, TargetID: targetID, Message: message}
}

func nativePolicyAllows(policy NativeEditPolicyV1, operation string) bool {
	for _, allowed := range policy.AllowedOperations {
		if allowed == operation {
			return true
		}
	}
	return false
}

func indexNativeTextTargets(doc *NativeDocumentV1) map[string]nativeTextTarget {
	targets := map[string]nativeTextTarget{}
	indexParagraph := func(paragraph *NativeParagraphV1) {
		if paragraph == nil {
			return
		}
		textRuns := make([]NativeRunV1, 0, len(paragraph.Runs))
		for _, run := range paragraph.Runs {
			if run.Kind != "text" || run.Text == nil {
				continue
			}
			textRuns = append(textRuns, run)
			targets["run\x00"+run.ID] = nativeTextTarget{kind: "run", partName: run.Anchor.PartName, path: run.Anchor.Path, anchor: run.Anchor, splice: run.Anchor, text: *run.Text, paragraph: paragraph}
		}
		if len(paragraph.Runs) == 1 && len(textRuns) == 1 {
			run := textRuns[0]
			targets["paragraph\x00"+paragraph.ID] = nativeTextTarget{kind: "paragraph", partName: paragraph.Anchor.PartName, path: run.Anchor.Path, anchor: paragraph.Anchor, splice: run.Anchor, text: *run.Text, paragraph: paragraph}
		}
	}
	var indexStory func(*NativeStoryV1)
	indexStory = func(story *NativeStoryV1) {
		for index := range story.Blocks {
			block := &story.Blocks[index]
			if block.Paragraph != nil {
				indexParagraph(block.Paragraph)
			}
			if block.Table != nil {
				for rowIndex := range block.Table.Rows {
					for cellIndex := range block.Table.Rows[rowIndex].Cells {
						cell := &block.Table.Rows[rowIndex].Cells[cellIndex]
						for paragraphIndex := range cell.Paragraphs {
							indexParagraph(&cell.Paragraphs[paragraphIndex])
						}
					}
				}
			}
		}
	}
	indexStory(&doc.Body)
	for _, stories := range [][]NativeStoryV1{doc.Headers, doc.Footers, doc.Notes, doc.CommentStories} {
		for index := range stories {
			indexStory(&stories[index])
		}
	}
	return targets
}

func replaceNativeTextElement(raw []byte, text string) ([]byte, error) {
	if len(raw) < 4 || raw[0] != '<' {
		return nil, fmt.Errorf("text anchor is not an XML element")
	}
	nameEnd := 1
	for nameEnd < len(raw) && raw[nameEnd] != ' ' && raw[nameEnd] != '\t' && raw[nameEnd] != '\r' && raw[nameEnd] != '\n' && raw[nameEnd] != '/' && raw[nameEnd] != '>' {
		nameEnd++
	}
	if nameEnd == 1 {
		return nil, fmt.Errorf("text anchor has no lexical element name")
	}
	name := raw[1:nameEnd]
	openEnd, quote := -1, byte(0)
	for index := nameEnd; index < len(raw); index++ {
		character := raw[index]
		if quote != 0 {
			if character == quote {
				quote = 0
			}
			continue
		}
		if character == '\'' || character == '"' {
			quote = character
			continue
		}
		if character == '>' {
			openEnd = index + 1
			break
		}
	}
	if openEnd < 0 || quote != 0 {
		return nil, fmt.Errorf("text anchor has an unterminated start tag")
	}
	if nativeMutationNeedsPreservedSpace(text) && !nativeMutationHasSpacePreserve(raw[:openEnd], nameEnd) {
		return nil, fmt.Errorf("leading or trailing XML whitespace requires an existing xml:space=preserve attribute")
	}
	beforeClose := openEnd - 2
	for beforeClose >= 0 && (raw[beforeClose] == ' ' || raw[beforeClose] == '\t' || raw[beforeClose] == '\r' || raw[beforeClose] == '\n') {
		beforeClose--
	}
	escaped := []byte(nativeMutationEscapeText(text))
	if beforeClose >= 0 && raw[beforeClose] == '/' {
		if text == "" {
			return append([]byte(nil), raw...), nil
		}
		opening := append([]byte(nil), raw[:beforeClose]...)
		opening = append(opening, raw[beforeClose+1:openEnd]...)
		opening = append(opening, escaped...)
		opening = append(opening, []byte("</")...)
		opening = append(opening, name...)
		opening = append(opening, '>')
		return opening, nil
	}
	closeAt := bytes.LastIndex(raw[openEnd:], []byte("</"))
	if closeAt < 0 {
		return nil, fmt.Errorf("text anchor has no end tag")
	}
	closeAt += openEnd
	if nativeMutationContainsXMLComment(raw[openEnd:closeAt]) {
		return nil, fmt.Errorf("text anchor contains an XML comment that must be preserved verbatim")
	}
	closing := raw[closeAt:]
	if !bytes.HasPrefix(closing, append([]byte("</"), name...)) || len(closing) <= len(name)+2 || !nativeMutationNameBoundary(closing[len(name)+2]) {
		return nil, fmt.Errorf("text anchor end tag does not match its start tag")
	}
	result := append([]byte(nil), raw[:openEnd]...)
	result = append(result, escaped...)
	result = append(result, closing...)
	return result, nil
}

func nativeMutationContainsXMLComment(content []byte) bool {
	for len(content) > 0 {
		markup := bytes.IndexByte(content, '<')
		if markup < 0 {
			return false
		}
		content = content[markup:]
		if bytes.HasPrefix(content, []byte("<!--")) {
			return true
		}
		if bytes.HasPrefix(content, []byte("<![CDATA[")) {
			end := bytes.Index(content[len("<![CDATA["):], []byte("]]>"))
			if end < 0 {
				return true
			}
			content = content[len("<![CDATA[")+end+len("]]>"):]
			continue
		}
		content = content[1:]
	}
	return false
}

func applyNativeTextSplices(source []byte, splices []nativeTextSplice) ([]byte, error) {
	sort.Slice(splices, func(i, j int) bool { return splices[i].start < splices[j].start })
	outputSize := int64(len(source))
	cursor := int64(0)
	for _, splice := range splices {
		if splice.start < cursor || splice.end < splice.start || splice.end > int64(len(source)) {
			return nil, nativeMutationError("OVERLAPPING_TARGETS", "", "paragraph and run targets may not overlap")
		}
		outputSize += int64(len(splice.text)) - (splice.end - splice.start)
		if outputSize > NativeDOCXMaxXMLPartBytes {
			return nil, nativeMutationError("RESOURCE_LIMIT", "", fmt.Sprintf("mutated XML part exceeds %d bytes", NativeDOCXMaxXMLPartBytes))
		}
		cursor = splice.end
	}
	var output bytes.Buffer
	output.Grow(int(outputSize))
	cursor = 0
	for _, splice := range splices {
		output.Write(source[cursor:splice.start])
		output.Write(splice.text)
		cursor = splice.end
	}
	output.Write(source[cursor:])
	return output.Bytes(), nil
}

func nativeMutationHasSpacePreserve(opening []byte, position int) bool {
	for position < len(opening) {
		for position < len(opening) && nativeMutationXMLSpace(opening[position]) {
			position++
		}
		if position >= len(opening) || opening[position] == '>' || opening[position] == '/' {
			return false
		}
		nameStart := position
		for position < len(opening) && !nativeMutationXMLSpace(opening[position]) && opening[position] != '=' && opening[position] != '>' && opening[position] != '/' {
			position++
		}
		name := string(opening[nameStart:position])
		for position < len(opening) && nativeMutationXMLSpace(opening[position]) {
			position++
		}
		if position >= len(opening) || opening[position] != '=' {
			return false
		}
		position++
		for position < len(opening) && nativeMutationXMLSpace(opening[position]) {
			position++
		}
		if position >= len(opening) || opening[position] != '\'' && opening[position] != '"' {
			return false
		}
		quote := opening[position]
		position++
		valueStart := position
		for position < len(opening) && opening[position] != quote {
			position++
		}
		if position >= len(opening) {
			return false
		}
		value := string(opening[valueStart:position])
		position++
		if name == "xml:space" && value == "preserve" {
			return true
		}
	}
	return false
}

func nativeMutationXMLSpace(character byte) bool {
	return character == ' ' || character == '\t' || character == '\r' || character == '\n'
}

func nativeMutationNameBoundary(character byte) bool {
	return character == '>' || nativeMutationXMLSpace(character)
}

func nativeMutationEscapeText(value string) string {
	replacer := strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;")
	return replacer.Replace(value)
}

func nativeMutationNeedsPreservedSpace(value string) bool {
	if value == "" {
		return false
	}
	first, _ := utf8.DecodeRuneInString(value)
	last, _ := utf8.DecodeLastRuneInString(value)
	isXMLSpace := func(character rune) bool {
		return character == ' ' || character == '\t' || character == '\n' || character == '\r'
	}
	return isXMLSpace(first) || isXMLSpace(last)
}

func nativeMutationXMLTextValid(value string) bool {
	for _, character := range value {
		if character == '\r' {
			return false // XML newline normalization would violate exact validation.
		}
		if character == '\t' || character == '\n' || character >= 0x20 && character <= 0xD7FF || character >= 0xE000 && character <= 0xFFFD || character >= 0x10000 && character <= 0x10FFFF {
			continue
		}
		return false
	}
	return true
}

func nativeMutationPathKey(kind, partName, path string) string {
	return kind + "\x00" + partName + "\x00" + path
}

func validateNativeMutationResults(before, after *NativeDocumentV1, expected map[string]string) error {
	if before.DocumentID != after.DocumentID {
		return nativeMutationError("POST_WRITE_MISMATCH", "", "document identity changed across native save")
	}
	if before.Source.MainPart != after.Source.MainPart {
		return nativeMutationError("POST_WRITE_MISMATCH", "", "main document part changed across native save")
	}
	if err := verifyNativeDOCXUnsupportedInventory(before.Unsupported, after.Unsupported); err != nil {
		return err
	}
	if !reflect.DeepEqual(before.PassthroughParts, after.PassthroughParts) {
		return nativeMutationError("POST_WRITE_MISMATCH", "", "passthrough part inventory changed across native save")
	}
	if err := verifyNativeDOCXTopology(before, after); err != nil {
		return err
	}
	beforeTargets, err := nativeMutationTargetsByPath(before)
	if err != nil {
		return err
	}
	afterTargets, err := nativeMutationTargetsByPath(after)
	if err != nil {
		return err
	}
	if err := verifyNativeMutationTargetInventory(beforeTargets, afterTargets); err != nil {
		return err
	}
	for key, want := range expected {
		if _, ok := beforeTargets[key]; !ok {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("requested native target %q exists only after save", key))
		}
		got, ok := afterTargets[key]
		if !ok || got != want {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("reopened native text at %q is %q, expected %q", key, got, want))
		}
	}
	return nil
}

func verifyNativeDOCXUnsupportedInventory(before, after []NativeUnsupportedCapabilityV1) error {
	if len(before) != len(after) {
		return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("unsupported inventory changed from %d to %d items", len(before), len(after)))
	}
	beforeByID := make(map[string]NativeUnsupportedCapabilityV1, len(before))
	afterByID := make(map[string]NativeUnsupportedCapabilityV1, len(after))
	for _, item := range before {
		if _, duplicate := beforeByID[item.ID]; duplicate {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("source unsupported inventory contains duplicate id %q", item.ID))
		}
		beforeByID[item.ID] = item
	}
	for _, item := range after {
		if _, duplicate := afterByID[item.ID]; duplicate {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("result unsupported inventory contains duplicate id %q", item.ID))
		}
		afterByID[item.ID] = item
	}
	for id, item := range beforeByID {
		candidate, ok := afterByID[id]
		if !ok || !nativeDOCXUnsupportedInventoryItemEqual(item, candidate) {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("source-authoritative unsupported item %q was lost or remapped", id))
		}
	}
	for id := range afterByID {
		if _, ok := beforeByID[id]; !ok {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("unsupported item %q exists only after save", id))
		}
	}
	return nil
}

func nativeDOCXUnsupportedInventoryItemEqual(before, after NativeUnsupportedCapabilityV1) bool {
	if before.ID != after.ID || before.Code != after.Code || before.Capability != after.Capability || before.ScopeID != after.ScopeID || before.Preservation != after.Preservation || before.Message != after.Message {
		return false
	}
	if before.Anchor == nil || after.Anchor == nil {
		return before.Anchor == nil && after.Anchor == nil
	}
	// Offsets and fingerprints are expected to move when an earlier or owning
	// supported text node changes. The source location itself may not remap.
	return before.Anchor.PartName == after.Anchor.PartName && before.Anchor.Path == after.Anchor.Path
}

func nativeMutationTargetsByPath(doc *NativeDocumentV1) (map[string]string, error) {
	result := map[string]string{}
	for _, target := range indexNativeTextTargets(doc) {
		key := nativeMutationPathKey(target.kind, target.partName, target.path)
		if _, duplicate := result[key]; duplicate {
			return nil, nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("native target inventory contains duplicate path %q", key))
		}
		result[key] = target.text
	}
	return result, nil
}

func verifyNativeMutationTargetInventory(before, after map[string]string) error {
	if len(before) != len(after) {
		return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("native text target inventory changed from %d to %d items", len(before), len(after)))
	}
	for key := range before {
		if _, ok := after[key]; !ok {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("native text target %q was lost after save", key))
		}
	}
	for key := range after {
		if _, ok := before[key]; !ok {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("native text target %q exists only after save", key))
		}
	}
	return nil
}

func verifyNativeDOCXTopology(before, after *NativeDocumentV1) error {
	want := nativeDOCXTopologyInventory(before)
	got := nativeDOCXTopologyInventory(after)
	if len(want) != len(got) {
		return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("document topology changed from %d to %d records", len(want), len(got)))
	}
	for index := range want {
		if want[index] != got[index] {
			return nativeMutationError("POST_WRITE_MISMATCH", "", fmt.Sprintf("document topology changed at record %d: %q became %q", index, want[index], got[index]))
		}
	}
	return nil
}

func nativeDOCXTopologyInventory(doc *NativeDocumentV1) []string {
	records := []string{"document|" + doc.DocumentID + "|" + doc.Source.MainPart + "|" + nativeMutationProjection(doc.Capabilities)}
	storyLocations := map[string]string{}
	blockLocations := map[string]string{}
	anchorPath := func(anchor *NativeSourceAnchorV1) string {
		if anchor == nil {
			return ""
		}
		return anchor.PartName + "|" + anchor.Path
	}
	var paragraph func(string, NativeParagraphV1)
	paragraph = func(location string, value NativeParagraphV1) {
		records = append(records, "paragraph|"+location+"|"+value.ID+"|"+anchorPath(&value.Anchor)+"|"+nativeMutationProjection(value.EditPolicy)+"|"+nativeMutationProjection(value.Properties))
		for runIndex, run := range value.Runs {
			runLocation := fmt.Sprintf("%s/run:%d", location, runIndex)
			referenceKind := ""
			if run.Reference != nil {
				referenceKind = run.Reference.Kind
			}
			referenceTarget := ""
			if run.Reference != nil {
				referenceTarget = run.Reference.TargetID
			}
			records = append(records, "run|"+runLocation+"|"+run.ID+"|"+run.Kind+"|"+run.Control+"|"+referenceKind+"|"+referenceTarget+"|"+anchorPath(&run.Anchor)+"|"+nativeMutationProjection(run.Properties))
			if run.Drawing != nil {
				drawing := run.Drawing
				records = append(records, "drawing|"+runLocation+"|"+drawing.ID+"|"+drawing.Placement+"|"+anchorPath(&drawing.Anchor)+"|"+nativeMutationProjection(struct {
					RelationshipID, MediaPart, ContentType, Name, AltText *string
					WidthEMU, HeightEMU, XEMU, YEMU                       *int64
					HorizontalRelativeFrom, VerticalRelativeFrom, Wrap    *string
					EditPolicy                                            NativeEditPolicyV1
				}{drawing.RelationshipID, drawing.MediaPart, drawing.ContentType, drawing.Name, drawing.AltText, drawing.WidthEMU, drawing.HeightEMU, drawing.XEMU, drawing.YEMU, drawing.HorizontalRelativeFrom, drawing.VerticalRelativeFrom, drawing.Wrap, drawing.EditPolicy}))
			}
		}
	}
	visitStory := func(collection string, index int, story NativeStoryV1) {
		location := fmt.Sprintf("%s:%d", collection, index)
		storyLocations[story.ID] = location
		nativeID := ""
		if story.NativeStoryID != nil {
			nativeID = *story.NativeStoryID
		}
		records = append(records, "story|"+location+"|"+story.ID+"|"+story.Kind+"|"+story.PartName+"|"+nativeID+"|"+anchorPath(story.Anchor))
		for blockIndex, block := range story.Blocks {
			blockLocation := fmt.Sprintf("%s/block:%d", location, blockIndex)
			blockLocations[block.ID] = blockLocation
			records = append(records, "block|"+blockLocation+"|"+block.ID+"|"+block.Kind)
			if block.Paragraph != nil {
				paragraph(blockLocation, *block.Paragraph)
			}
			if block.Table != nil {
				table := block.Table
				records = append(records, "table|"+blockLocation+"|"+table.ID+"|"+anchorPath(&table.Anchor)+"|"+nativeMutationProjection(struct {
					EditPolicy      NativeEditPolicyV1
					TableStyleID    *string
					WidthTwips      *int64
					Layout          *string
					Alignment       *string
					IndentTwips     *int64
					GridWidthsTwips []int64
					CellMargins     *NativeTableCellMarginsV1
					Borders         *NativeTableBordersV1
				}{table.EditPolicy, table.TableStyleID, table.WidthTwips, table.Layout, table.Alignment, table.IndentTwips, table.GridWidthsTwips, table.CellMargins, table.Borders}))
				for rowIndex, row := range table.Rows {
					rowLocation := fmt.Sprintf("%s/row:%d", blockLocation, rowIndex)
					records = append(records, "row|"+rowLocation+"|"+row.ID+"|"+anchorPath(&row.Anchor)+"|"+nativeMutationProjection(struct {
						HeightTwips  *int64
						HeightRule   *string
						RepeatHeader *bool
						CantSplit    *bool
					}{row.HeightTwips, row.HeightRule, row.RepeatHeader, row.CantSplit}))
					for cellIndex, cell := range row.Cells {
						cellLocation := fmt.Sprintf("%s/cell:%d", rowLocation, cellIndex)
						records = append(records, "cell|"+cellLocation+"|"+cell.ID+"|"+anchorPath(&cell.Anchor)+"|"+nativeMutationProjection(struct {
							WidthTwips    *int64
							GridSpan      *int
							VerticalMerge string
							Borders       *NativeTableBordersV1
							ShadingRGB    *string
						}{cell.WidthTwips, cell.GridSpan, cell.VerticalMerge, cell.Borders, cell.ShadingRGB}))
						for paragraphIndex, nested := range cell.Paragraphs {
							paragraph(fmt.Sprintf("%s/paragraph:%d", cellLocation, paragraphIndex), nested)
						}
					}
				}
			}
		}
	}
	visitStory("body", 0, doc.Body)
	for _, collection := range []struct {
		name    string
		stories []NativeStoryV1
	}{{"header", doc.Headers}, {"footer", doc.Footers}, {"note", doc.Notes}, {"comment-story", doc.CommentStories}} {
		for index, story := range collection.stories {
			visitStory(collection.name, index, story)
		}
	}
	for index, section := range doc.Sections {
		location := fmt.Sprintf("section:%d", index)
		records = append(records, "section|"+location+"|"+section.ID+"|"+anchorPath(&section.Anchor)+"|"+blockLocations[section.StartsAtBlockID]+"|"+section.BreakType+"|"+nativeMutationProjection(section.TitlePage)+"|"+nativeMutationProjection(section.Page))
		for refIndex, ref := range section.HeaderRefs {
			records = append(records, fmt.Sprintf("header-ref|%s/%d|%s|%s|%s", location, refIndex, ref.Kind, ref.RelationshipID, storyLocations[ref.StoryID]))
		}
		for refIndex, ref := range section.FooterRefs {
			records = append(records, fmt.Sprintf("footer-ref|%s/%d|%s|%s|%s", location, refIndex, ref.Kind, ref.RelationshipID, storyLocations[ref.StoryID]))
		}
	}
	for index, comment := range doc.Comments {
		records = append(records, fmt.Sprintf("comment|%d|%s|%s|%s|%s|%s|%s|%s", index, comment.ID, comment.NativeCommentID, comment.Author, nativeMutationProjection(comment.Initials), nativeMutationProjection(comment.CreatedAt), anchorPath(comment.Anchor), storyLocations[comment.BodyStoryID]))
	}
	return records
}

func nativeMutationProjection(value any) string {
	encoded, err := json.Marshal(value)
	if err != nil {
		return "<invalid>"
	}
	return string(encoded)
}

func nativeMutationUTF16CodeUnits(value string) int {
	units := 0
	for _, character := range value {
		units++
		if character > 0xFFFF {
			units++
		}
	}
	return units
}
