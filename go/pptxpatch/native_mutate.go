package pptxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"unicode/utf8"
)

const (
	MaxNativePPTXMutationPayloadBytes = 3 << 20
	nativeMaxMutations                = 10_000
	nativeMaxMutationParagraphs       = 10_000
	nativeMaxMutationRuns             = 100_000
	nativeMaxMutationNodes            = nativeMaxMutationParagraphs + nativeMaxMutationRuns
	nativeMaxMutationJSONDepth        = 64
)

// NativePPTXMutationKind is the source-authoritative PPTX mutation subset.
// It deliberately excludes structural edits, pictures, charts, tables, and
// any mutation that would require reconstructing a slide from a render model.
type NativePPTXMutationKind string

const (
	NativePPTXReplaceText     NativePPTXMutationKind = "text.replace"
	NativePPTXUpdateAutoShape NativePPTXMutationKind = "autoshape.update"
)

// NativePPTXMutationRequest binds an atomic mutation batch to the exact input
// package revision returned by ExtractNativePPTX. A caller must re-extract and
// retry after any stale-revision failure.
type NativePPTXMutationRequest struct {
	ExpectedSourceRevision string               `json:"expectedSourceRevision"`
	Operations             []NativePPTXMutation `json:"operations"`
}

// NativePPTXMutation addresses an extracted element by durable ID and exact
// source-subtree fingerprint. Exactly one payload is permitted by Kind.
type NativePPTXMutation struct {
	OperationID               string                       `json:"operationId"`
	Kind                      NativePPTXMutationKind       `json:"kind"`
	ElementID                 string                       `json:"elementId"`
	ExpectedFingerprintSHA256 string                       `json:"expectedFingerprintSha256"`
	Paragraphs                *[]NativeParagraph           `json:"paragraphs,omitempty"`
	AutoShape                 *NativePPTXAutoShapeMutation `json:"autoShape,omitempty"`
}

// NativePPTXAutoShapeMutation is a complete replacement for the exact native
// AutoShape property subset. Nil Fill means explicit a:noFill; nil Stroke
// means an explicit no-fill outline. Partial or inherited shape semantics are
// not accepted.
type NativePPTXAutoShapeMutation struct {
	Transform NativeTransform   `json:"transform"`
	Preset    NativeShapePreset `json:"preset"`
	Fill      *string           `json:"fill,omitempty"`
	Stroke    *NativeStroke     `json:"stroke,omitempty"`
}

type nativePPTXResolvedMutation struct {
	operation NativePPTXMutation
	element   NativeElement
	part      string
	objectID  string
}

type nativeXMLReplacement struct {
	start int64
	end   int64
	data  []byte
}

type nativeMutationSlide struct {
	data       []byte
	dialect    nativeExtractDialect
	shapeNodes map[string]*nativeXMLNode
}

type nativeMutationBudget struct {
	paragraphs    int
	runs          int
	nodes         int
	textCodeUnits int64
}

// DecodeNativePPTXMutationRequest strictly decodes the native PPTX payload
// carried by the shared Office mutation envelope. outerExpectedRevision is
// the envelope's exact-byte sha256:<digest> CAS; the payload must bind the
// native extractor's rev-<digest> spelling of the same digest.
func DecodeNativePPTXMutationRequest(payload []byte, outerExpectedRevision string) (NativePPTXMutationRequest, error) {
	nativeRevision, err := nativePPTXRevisionFromOuterCAS(outerExpectedRevision)
	if err != nil {
		return NativePPTXMutationRequest{}, err
	}
	if len(payload) == 0 || len(payload) > MaxNativePPTXMutationPayloadBytes || !utf8.Valid(payload) {
		return NativePPTXMutationRequest{}, fmt.Errorf("pptxpatch: native mutations: encoded payload must be valid UTF-8 in 1..%d bytes", MaxNativePPTXMutationPayloadBytes)
	}
	trimmed := bytes.TrimSpace(payload)
	if len(trimmed) == 0 || trimmed[0] != '{' {
		return NativePPTXMutationRequest{}, fmt.Errorf("pptxpatch: native mutations: encoded payload must be one JSON object")
	}
	if err := validateNativePPTXMutationJSON(trimmed); err != nil {
		return NativePPTXMutationRequest{}, fmt.Errorf("pptxpatch: native mutations: invalid encoded payload: %w", err)
	}
	if err := validateNativePPTXMutationFieldNames(trimmed); err != nil {
		return NativePPTXMutationRequest{}, fmt.Errorf("pptxpatch: native mutations: invalid encoded payload: %w", err)
	}

	decoder := json.NewDecoder(bytes.NewReader(trimmed))
	decoder.DisallowUnknownFields()
	var request NativePPTXMutationRequest
	if err := decoder.Decode(&request); err != nil {
		return NativePPTXMutationRequest{}, fmt.Errorf("pptxpatch: native mutations: invalid encoded payload: %w", err)
	}
	if request.ExpectedSourceRevision != nativeRevision {
		return NativePPTXMutationRequest{}, fmt.Errorf("pptxpatch: native mutations: payload source revision %q does not match outer CAS %q", request.ExpectedSourceRevision, outerExpectedRevision)
	}
	if err := validateNativePPTXMutationRequestPayload(request); err != nil {
		return NativePPTXMutationRequest{}, err
	}
	return request, nil
}

// ApplyNativePPTXMutationPayload is the native policy bridge for a strictly
// decoded shared Office mutation payload. It independently rechecks the outer
// exact-byte CAS before invoking the typed source-authoritative applier.
func ApplyNativePPTXMutationPayload(orig []byte, outerExpectedRevision string, payload []byte) ([]byte, error) {
	if outerExpectedRevision != "sha256:"+nativeSHA256(orig) {
		return nil, fmt.Errorf("pptxpatch: native mutations: stale outer source revision")
	}
	request, err := DecodeNativePPTXMutationRequest(payload, outerExpectedRevision)
	if err != nil {
		return nil, err
	}
	return ApplyNativePPTXMutations(orig, request)
}

func nativePPTXRevisionFromOuterCAS(outerExpectedRevision string) (string, error) {
	const prefix = "sha256:"
	if !strings.HasPrefix(outerExpectedRevision, prefix) || !sha256Pattern.MatchString(strings.TrimPrefix(outerExpectedRevision, prefix)) {
		return "", fmt.Errorf("pptxpatch: native mutations: outer expected revision must be sha256 followed by the full lowercase exact-byte digest")
	}
	return "rev-" + strings.TrimPrefix(outerExpectedRevision, prefix), nil
}

func validateNativePPTXMutationJSON(payload []byte) error {
	decoder := json.NewDecoder(bytes.NewReader(payload))
	decoder.UseNumber()
	if err := scanNativePPTXMutationJSONValue(decoder, 0); err != nil {
		return err
	}
	if token, err := decoder.Token(); err != io.EOF || token != nil {
		return fmt.Errorf("trailing JSON is not permitted")
	}
	return nil
}

func scanNativePPTXMutationJSONValue(decoder *json.Decoder, depth int) error {
	if depth > nativeMaxMutationJSONDepth {
		return fmt.Errorf("JSON nesting exceeds %d", nativeMaxMutationJSONDepth)
	}
	token, err := decoder.Token()
	if err != nil {
		return fmt.Errorf("invalid JSON: %w", err)
	}
	delimiter, compound := token.(json.Delim)
	if !compound {
		return nil
	}
	switch delimiter {
	case '{':
		seen := map[string]bool{}
		for decoder.More() {
			nameToken, err := decoder.Token()
			if err != nil {
				return fmt.Errorf("invalid object field: %w", err)
			}
			name, ok := nameToken.(string)
			if !ok {
				return fmt.Errorf("object field name is not a string")
			}
			if seen[name] {
				return fmt.Errorf("duplicate field %q", name)
			}
			seen[name] = true
			if err := scanNativePPTXMutationJSONValue(decoder, depth+1); err != nil {
				return err
			}
		}
		closing, err := decoder.Token()
		if err != nil || closing != json.Delim('}') {
			return fmt.Errorf("unterminated JSON object")
		}
	case '[':
		for decoder.More() {
			if err := scanNativePPTXMutationJSONValue(decoder, depth+1); err != nil {
				return err
			}
		}
		closing, err := decoder.Token()
		if err != nil || closing != json.Delim(']') {
			return fmt.Errorf("unterminated JSON array")
		}
	default:
		return fmt.Errorf("unexpected JSON delimiter %q", delimiter)
	}
	return nil
}

func validateNativePPTXMutationFieldNames(payload []byte) error {
	request, err := nativePPTXJSONObject(payload, map[string]bool{
		"expectedSourceRevision": true,
		"operations":             true,
	})
	if err != nil {
		return err
	}
	var operations []json.RawMessage
	if raw, ok := request["operations"]; ok {
		if err := json.Unmarshal(raw, &operations); err != nil {
			return err
		}
	}
	for _, rawOperation := range operations {
		operation, err := nativePPTXJSONObject(rawOperation, map[string]bool{
			"operationId":               true,
			"kind":                      true,
			"elementId":                 true,
			"expectedFingerprintSha256": true,
			"paragraphs":                true,
			"autoShape":                 true,
		})
		if err != nil {
			return err
		}
		if raw, ok := operation["paragraphs"]; ok {
			if err := validateNativePPTXParagraphFieldNames(raw); err != nil {
				return err
			}
		}
		if raw, ok := operation["autoShape"]; ok && string(raw) != "null" {
			autoShape, err := nativePPTXJSONObject(raw, map[string]bool{
				"transform": true,
				"preset":    true,
				"fill":      true,
				"stroke":    true,
			})
			if err != nil {
				return err
			}
			if raw, ok := autoShape["transform"]; ok {
				if _, err := nativePPTXJSONObject(raw, map[string]bool{"x": true, "y": true, "cx": true, "cy": true}); err != nil {
					return err
				}
			}
			if raw, ok := autoShape["stroke"]; ok && string(raw) != "null" {
				if _, err := nativePPTXJSONObject(raw, map[string]bool{
					"color": true, "widthEmu": true, "cap": true, "join": true, "dash": true, "miterLimit": true,
				}); err != nil {
					return err
				}
			}
		}
	}
	return nil
}

func validateNativePPTXParagraphFieldNames(raw json.RawMessage) error {
	var paragraphs []json.RawMessage
	if err := json.Unmarshal(raw, &paragraphs); err != nil {
		return err
	}
	for _, rawParagraph := range paragraphs {
		paragraph, err := nativePPTXJSONObject(rawParagraph, map[string]bool{
			"runs": true, "align": true, "level": true, "bullet": true,
		})
		if err != nil {
			return err
		}
		var runs []json.RawMessage
		if rawRuns, ok := paragraph["runs"]; ok {
			if err := json.Unmarshal(rawRuns, &runs); err != nil {
				return err
			}
		}
		for _, rawRun := range runs {
			if _, err := nativePPTXJSONObject(rawRun, map[string]bool{
				"text": true, "bold": true, "italic": true, "fontSizeHundredthPt": true, "color": true, "fontFamily": true,
			}); err != nil {
				return err
			}
		}
	}
	return nil
}

func nativePPTXJSONObject(raw json.RawMessage, allowed map[string]bool) (map[string]json.RawMessage, error) {
	var object map[string]json.RawMessage
	if err := json.Unmarshal(raw, &object); err != nil {
		return nil, err
	}
	if object == nil {
		return nil, fmt.Errorf("expected JSON object")
	}
	for name := range object {
		if !allowed[name] {
			return nil, fmt.Errorf("unknown field %q", name)
		}
	}
	return object, nil
}

// ApplyNativePPTXMutations applies exact text or AutoShape property changes to
// source OOXML. It never calls the reconstructive deck reader or writer.
// Unchanged OPC entries are raw-copied and then byte-verified. Before success,
// the result is reopened through ExtractNativePPTX and each requested value,
// identity, fingerprint, and untouched package part is validated.
func ApplyNativePPTXMutations(orig []byte, request NativePPTXMutationRequest) ([]byte, error) {
	if err := validateNativePPTXMutationRequestPayload(request); err != nil {
		return nil, err
	}
	wantSourceRevision := "rev-" + nativeSHA256(orig)
	if request.ExpectedSourceRevision != wantSourceRevision {
		if !strings.HasPrefix(request.ExpectedSourceRevision, "rev-") || !sha256Pattern.MatchString(strings.TrimPrefix(request.ExpectedSourceRevision, "rev-")) {
			return nil, fmt.Errorf("pptxpatch: native mutations: invalid expected source revision")
		}
		return nil, fmt.Errorf("pptxpatch: native mutations: stale source revision: got %q want %q", request.ExpectedSourceRevision, wantSourceRevision)
	}

	extractOptions := nativeMutationExtractOptions()
	before, err := ExtractNativePPTX(orig, extractOptions)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: extract source: %w", err)
	}
	if before.SourceRevision == nil || *before.SourceRevision != request.ExpectedSourceRevision {
		return nil, fmt.Errorf("pptxpatch: native mutations: stale source revision: got %q want %q", request.ExpectedSourceRevision, nativeStringValue(before.SourceRevision))
	}

	resolved, err := resolveNativePPTXMutations(before, request.Operations)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeExtractPackage(orig)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: reopen source OPC: %w", err)
	}

	slides := map[string]*nativeMutationSlide{}
	replacements := map[string][]nativeXMLReplacement{}
	for _, mutation := range resolved {
		slide := slides[mutation.part]
		if slide == nil {
			parsed, parseErr := parseNativeMutationSlide(pkg.parts[mutation.part], mutation.part)
			if parseErr != nil {
				return nil, parseErr
			}
			slide = parsed
			slides[mutation.part] = slide
		}
		node := slide.shapeNodes[mutation.objectID]
		if node == nil {
			return nil, fmt.Errorf("pptxpatch: native mutations: operation %q: source object %q is missing from %q", mutation.operation.OperationID, mutation.objectID, mutation.part)
		}
		var replacement nativeXMLReplacement
		switch mutation.operation.Kind {
		case NativePPTXReplaceText:
			replacement, err = nativeTextParagraphReplacement(node, slide.dialect, *mutation.operation.Paragraphs)
		case NativePPTXUpdateAutoShape:
			replacement, err = nativeAutoShapePropertyReplacement(node, slide.dialect, *mutation.operation.AutoShape)
		default:
			panic("validated native mutation kind")
		}
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: native mutations: operation %q: %w", mutation.operation.OperationID, err)
		}
		replacements[mutation.part] = append(replacements[mutation.part], replacement)
	}

	partReplacements := make(map[string][]byte, len(replacements))
	for part, edits := range replacements {
		updated, patchErr := applyNativeXMLReplacements(slides[part].data, edits)
		if patchErr != nil {
			return nil, fmt.Errorf("pptxpatch: native mutations: part %q: %w", part, patchErr)
		}
		if !bytes.Equal(updated, slides[part].data) {
			partReplacements[part] = updated
		}
	}
	if len(partReplacements) == 0 {
		return nil, fmt.Errorf("pptxpatch: native mutations: batch produced no semantic OOXML change")
	}

	produced, err := applyNativePPTXPartReplacements(orig, partReplacements)
	if err != nil {
		return nil, err
	}
	afterOptions := nativeMutationExtractOptions()
	afterOptions.Previous = &before
	after, err := ExtractNativePPTX(produced, afterOptions)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: reopen/extract produced package: %w", err)
	}
	if err := verifyNativePPTXMutationResult(before, after, resolved); err != nil {
		return nil, err
	}
	return produced, nil
}

func validateNativePPTXMutationRequestPayload(request NativePPTXMutationRequest) error {
	if len(request.Operations) == 0 {
		return fmt.Errorf("pptxpatch: native mutations: empty batch")
	}
	if len(request.Operations) > nativeMaxMutations {
		return fmt.Errorf("pptxpatch: native mutations: batch exceeds %d operations", nativeMaxMutations)
	}
	seenOperations := map[string]bool{}
	seenElements := map[string]bool{}
	budget := nativeMutationBudget{}
	for index, operation := range request.Operations {
		prefix := fmt.Sprintf("pptxpatch: native mutations: operation %d", index)
		if !nativeIDPattern.MatchString(operation.OperationID) {
			return fmt.Errorf("%s: invalid operation id %q", prefix, operation.OperationID)
		}
		if seenOperations[operation.OperationID] {
			return fmt.Errorf("%s: duplicate operation id %q", prefix, operation.OperationID)
		}
		seenOperations[operation.OperationID] = true
		if !nativeIDPattern.MatchString(operation.ElementID) {
			return fmt.Errorf("%s: invalid element id", prefix)
		}
		if seenElements[operation.ElementID] {
			return fmt.Errorf("%s: multiple operations target element %q", prefix, operation.ElementID)
		}
		seenElements[operation.ElementID] = true
		if !sha256Pattern.MatchString(operation.ExpectedFingerprintSHA256) {
			return fmt.Errorf("%s: invalid expected element fingerprint", prefix)
		}
		switch operation.Kind {
		case NativePPTXReplaceText:
			if operation.Paragraphs == nil || operation.AutoShape != nil {
				return fmt.Errorf("%s: text.replace requires only paragraphs", prefix)
			}
			if err := validateNativeMutationParagraphs(*operation.Paragraphs, &budget); err != nil {
				return fmt.Errorf("%s: %w", prefix, err)
			}
		case NativePPTXUpdateAutoShape:
			if operation.AutoShape == nil || operation.Paragraphs != nil {
				return fmt.Errorf("%s: autoshape.update requires only autoShape", prefix)
			}
			if err := validateNativeAutoShapeMutation(*operation.AutoShape); err != nil {
				return fmt.Errorf("%s: %w", prefix, err)
			}
		default:
			return fmt.Errorf("%s: unsupported mutation kind %q", prefix, operation.Kind)
		}
	}
	return nil
}

func nativeMutationExtractOptions() NativePPTXExtractOptions {
	return NativePPTXExtractOptions{TokenFactory: nativeMutationTokenFactory{}}
}

// nativeMutationTokenFactory has no external publication side effects, but it
// still implements the extractor's all-or-nothing transaction contract. This
// keeps group projection atomic during both the before and after extraction;
// a late group refusal cannot leak a descendant capability into verification.
type nativeMutationTokenFactory struct{}

func (nativeMutationTokenFactory) IssueNativePassthroughToken(request NativePassthroughTokenRequest) (string, error) {
	return nativeMutationPassthroughToken(request), nil
}

func (nativeMutationTokenFactory) BeginNativePassthroughTokenTransaction() (NativePassthroughTokenTransaction, error) {
	return nativeMutationTokenTransaction{}, nil
}

type nativeMutationTokenTransaction struct{}

func (nativeMutationTokenTransaction) IssueNativePassthroughToken(request NativePassthroughTokenRequest) (string, error) {
	return nativeMutationPassthroughToken(request), nil
}

func (nativeMutationTokenTransaction) CommitNativePassthroughTokens() error { return nil }

func (nativeMutationTokenTransaction) RollbackNativePassthroughTokens() {}

func nativeMutationPassthroughToken(request NativePassthroughTokenRequest) string {
	digest := nativeSHA256([]byte(request.SourceRevision + "\x00" + request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.FingerprintSHA256 + "\x00" + request.Reason))
	return "native-mutation-" + digest[:32]
}

func nativeStringValue(value *string) string {
	if value == nil {
		return ""
	}
	return *value
}

func resolveNativePPTXMutations(deck NativePPTXDeck, operations []NativePPTXMutation) ([]nativePPTXResolvedMutation, error) {
	elements, err := indexNativePPTXElements(deck)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: source contract: %w", err)
	}
	seenOperations := map[string]bool{}
	seenElements := map[string]bool{}
	budget := nativeMutationBudget{}
	resolved := make([]nativePPTXResolvedMutation, 0, len(operations))
	for index, operation := range operations {
		prefix := fmt.Sprintf("pptxpatch: native mutations: operation %d", index)
		if !nativeIDPattern.MatchString(operation.OperationID) {
			return nil, fmt.Errorf("%s: invalid operation id %q", prefix, operation.OperationID)
		}
		if seenOperations[operation.OperationID] {
			return nil, fmt.Errorf("%s: duplicate operation id %q", prefix, operation.OperationID)
		}
		seenOperations[operation.OperationID] = true
		if !nativeIDPattern.MatchString(operation.ElementID) {
			return nil, fmt.Errorf("%s: invalid element id", prefix)
		}
		if seenElements[operation.ElementID] {
			return nil, fmt.Errorf("%s: multiple operations target element %q", prefix, operation.ElementID)
		}
		seenElements[operation.ElementID] = true
		if !sha256Pattern.MatchString(operation.ExpectedFingerprintSHA256) {
			return nil, fmt.Errorf("%s: invalid expected element fingerprint", prefix)
		}
		element, ok := elements[operation.ElementID]
		if !ok || element.Source == nil {
			return nil, fmt.Errorf("%s: parsed source element %q was not found", prefix, operation.ElementID)
		}
		if element.Source.FingerprintSHA256 != operation.ExpectedFingerprintSHA256 {
			return nil, fmt.Errorf("%s: stale element fingerprint for %q", prefix, operation.ElementID)
		}
		if element.Compatibility.Status == NativeCompatibilityStatusRefused {
			return nil, fmt.Errorf("%s: element %q is refused and cannot be mutated", prefix, operation.ElementID)
		}
		if element.Transform.QuarterTurns != nil {
			return nil, fmt.Errorf("%s: source quarter-turn transforms are preview-only", prefix)
		}
		// These flags are omitted from native paragraphs. Equal compatibility
		// summaries after rewriting cannot prove that local metadata survived.
		if operation.Kind == NativePPTXReplaceText {
			for _, diagnostic := range element.Compatibility.Diagnostics {
				if diagnostic.Code == "pptx.text-checking-metadata-preserved" {
					return nil, fmt.Errorf("%s: text checking metadata is preserve-only and cannot be replaced", prefix)
				}
			}
		}
		switch operation.Kind {
		case NativePPTXReplaceText:
			if operation.Paragraphs == nil || operation.AutoShape != nil {
				return nil, fmt.Errorf("%s: text.replace requires only paragraphs", prefix)
			}
			if element.Kind != NativeElementKindText && element.Kind != NativeElementKindShape {
				return nil, fmt.Errorf("%s: text.replace target is not text-bearing", prefix)
			}
			if element.Paragraphs == nil {
				return nil, fmt.Errorf("%s: target has no exact native paragraphs", prefix)
			}
			if err := validateNativeMutationParagraphs(*operation.Paragraphs, &budget); err != nil {
				return nil, fmt.Errorf("%s: %w", prefix, err)
			}
			if nativeParagraphsEqual(*element.Paragraphs, *operation.Paragraphs) {
				return nil, fmt.Errorf("%s: semantic no-op text replacement", prefix)
			}
		case NativePPTXUpdateAutoShape:
			if operation.AutoShape == nil || operation.Paragraphs != nil {
				return nil, fmt.Errorf("%s: autoshape.update requires only autoShape", prefix)
			}
			if element.Kind != NativeElementKindShape || element.Preset == nil {
				return nil, fmt.Errorf("%s: target is not an exact native AutoShape", prefix)
			}
			if err := validateNativeAutoShapeMutation(*operation.AutoShape); err != nil {
				return nil, fmt.Errorf("%s: %w", prefix, err)
			}
			if nativeAutoShapeEquals(element, *operation.AutoShape) {
				return nil, fmt.Errorf("%s: semantic no-op AutoShape update", prefix)
			}
		default:
			return nil, fmt.Errorf("%s: unsupported mutation kind %q", prefix, operation.Kind)
		}
		resolved = append(resolved, nativePPTXResolvedMutation{operation: operation, element: element, part: element.Source.PartName, objectID: element.Source.ObjectID})
	}
	return resolved, nil
}

func indexNativePPTXElements(deck NativePPTXDeck) (map[string]NativeElement, error) {
	elements := make(map[string]NativeElement)
	slides := make(map[string]bool, len(deck.Slides))
	count := 0
	var walk func([]NativeElement, int) error
	walk = func(values []NativeElement, depth int) error {
		if len(values) == 0 {
			return nil
		}
		if depth > nativeMaxDepth {
			return fmt.Errorf("element nesting exceeds %d", nativeMaxDepth)
		}
		for _, element := range values {
			if count >= nativeMaxTotalElements {
				return fmt.Errorf("element inventory exceeds %d", nativeMaxTotalElements)
			}
			count++
			if _, duplicate := elements[element.ID]; duplicate {
				return fmt.Errorf("duplicate element id %q", element.ID)
			}
			elements[element.ID] = element
			if err := walk(element.Children, depth+1); err != nil {
				return err
			}
		}
		return nil
	}
	for _, slide := range deck.Slides {
		if slides[slide.ID] {
			return nil, fmt.Errorf("duplicate slide id %q", slide.ID)
		}
		slides[slide.ID] = true
		if err := walk(slide.Elements, 1); err != nil {
			return nil, err
		}
	}
	return elements, nil
}

func validateNativeMutationParagraphs(paragraphs []NativeParagraph, budget *nativeMutationBudget) error {
	if len(paragraphs) == 0 {
		return fmt.Errorf("text replacement requires at least one exact paragraph")
	}
	if len(paragraphs) > nativeMaxParagraphsPerElement {
		return fmt.Errorf("paragraph count exceeds %d", nativeMaxParagraphsPerElement)
	}
	if budget == nil {
		return fmt.Errorf("missing aggregate mutation budget")
	}
	if len(paragraphs) > nativeMaxMutationParagraphs-budget.paragraphs || len(paragraphs) > nativeMaxMutationNodes-budget.nodes {
		return fmt.Errorf("aggregate paragraph/node budget exceeded")
	}
	budget.paragraphs += len(paragraphs)
	budget.nodes += len(paragraphs)
	for paragraphIndex, paragraph := range paragraphs {
		if paragraph.Runs == nil || paragraph.Align == nil || paragraph.Level == nil || paragraph.Bullet == nil {
			return fmt.Errorf("paragraph %d is not self-contained", paragraphIndex)
		}
		if *paragraph.Align != NativeTextAlignLeft && *paragraph.Align != NativeTextAlignCenter && *paragraph.Align != NativeTextAlignRight {
			return fmt.Errorf("paragraph %d has unsupported alignment", paragraphIndex)
		}
		if *paragraph.Level < 0 || *paragraph.Level > 8 {
			return fmt.Errorf("paragraph %d has invalid level", paragraphIndex)
		}
		if *paragraph.Bullet {
			return fmt.Errorf("paragraph %d uses bullet semantics outside the exact native mutation subset", paragraphIndex)
		}
		if len(paragraph.Runs) > nativeMaxRunsPerParagraph {
			return fmt.Errorf("paragraph %d run count exceeds %d", paragraphIndex, nativeMaxRunsPerParagraph)
		}
		if len(paragraph.Runs) > nativeMaxMutationRuns-budget.runs || len(paragraph.Runs) > nativeMaxMutationNodes-budget.nodes {
			return fmt.Errorf("aggregate run/node budget exceeded")
		}
		budget.runs += len(paragraph.Runs)
		budget.nodes += len(paragraph.Runs)
		for runIndex, run := range paragraph.Runs {
			if run.Text == nil || run.Bold == nil || run.Italic == nil || run.FontSizeHundredthPt == nil || run.Color == nil || run.FontFamily == nil {
				return fmt.Errorf("paragraph %d run %d is not self-contained", paragraphIndex, runIndex)
			}
			if !utf8.ValidString(*run.Text) || !utf8.ValidString(*run.FontFamily) {
				return fmt.Errorf("paragraph %d run %d is not valid UTF-8", paragraphIndex, runIndex)
			}
			if strings.ContainsAny(*run.Text, "\t\r\n") {
				return fmt.Errorf("paragraph %d run %d contains literal tab or line-break text outside the exact native mutation subset", paragraphIndex, runIndex)
			}
			units := int64(utf16CodeUnitLengthBounded(*run.Text, nativeMaxTextCodeUnits+1))
			if units > nativeMaxTextCodeUnits || budget.textCodeUnits > nativeMaxTotalTextCodeUnits-units {
				return fmt.Errorf("text resource budget exceeded")
			}
			budget.textCodeUnits += units
			if *run.FontSizeHundredthPt < 1 || *run.FontSizeHundredthPt > 400000 || !colorPattern.MatchString(*run.Color) {
				return fmt.Errorf("paragraph %d run %d has invalid exact formatting", paragraphIndex, runIndex)
			}
			familyUnits := utf16CodeUnitLengthBounded(*run.FontFamily, 257)
			if familyUnits == 0 || familyUnits > 256 {
				return fmt.Errorf("paragraph %d run %d has invalid font family", paragraphIndex, runIndex)
			}
			if _, err := nativeEscapeXML(*run.Text); err != nil {
				return fmt.Errorf("paragraph %d run %d has invalid XML text", paragraphIndex, runIndex)
			}
			if _, err := nativeEscapeXML(*run.FontFamily); err != nil {
				return fmt.Errorf("paragraph %d run %d has invalid font family", paragraphIndex, runIndex)
			}
		}
	}
	return nil
}

func validateNativeAutoShapeMutation(shape NativePPTXAutoShapeMutation) error {
	if shape.Transform.QuarterTurns != nil {
		return fmt.Errorf("quarter-turn transforms are preview-only")
	}
	if shape.Transform.X == nil || shape.Transform.Y == nil || shape.Transform.Cx == nil || shape.Transform.Cy == nil {
		return fmt.Errorf("AutoShape transform must be complete")
	}
	if *shape.Transform.X < -nativeMaxSafeInteger || *shape.Transform.X > nativeMaxSafeInteger || *shape.Transform.Y < -nativeMaxSafeInteger || *shape.Transform.Y > nativeMaxSafeInteger || *shape.Transform.Cx < 1 || *shape.Transform.Cx > nativeMaxSafeInteger || *shape.Transform.Cy < 1 || *shape.Transform.Cy > nativeMaxSafeInteger {
		return fmt.Errorf("AutoShape transform is outside the exact native range")
	}
	switch shape.Preset {
	case NativeShapePresetRect, NativeShapePresetEllipse, NativeShapePresetTriangle, NativeShapePresetDiamond:
	default:
		return fmt.Errorf("AutoShape preset %q is outside the exact native subset", shape.Preset)
	}
	if shape.Fill != nil && !colorPattern.MatchString(*shape.Fill) {
		return fmt.Errorf("AutoShape fill must be canonical uppercase sRGB")
	}
	if shape.Stroke == nil {
		return nil
	}
	stroke := shape.Stroke
	if stroke.WidthEMU == nil || *stroke.WidthEMU < 0 || *stroke.WidthEMU > nativeMaxLineWidthEmu || !colorPattern.MatchString(stroke.Color) || stroke.Cap == nil || stroke.Join == nil || stroke.Dash == nil || *stroke.Dash != NativeStrokeDashSolid {
		return fmt.Errorf("AutoShape stroke is not complete exact solid-line metadata")
	}
	if *stroke.Cap != NativeStrokeCapFlat && *stroke.Cap != NativeStrokeCapRound && *stroke.Cap != NativeStrokeCapSquare {
		return fmt.Errorf("AutoShape stroke cap is unsupported")
	}
	switch *stroke.Join {
	case NativeStrokeJoinRound, NativeStrokeJoinBevel:
		if stroke.MiterLimit != nil {
			return fmt.Errorf("AutoShape non-miter stroke must not include miterLimit")
		}
	case NativeStrokeJoinMiter:
		if stroke.MiterLimit == nil || *stroke.MiterLimit < 0 || *stroke.MiterLimit > nativeMaxDrawingPercentage {
			return fmt.Errorf("AutoShape miter stroke requires a bounded miterLimit")
		}
	default:
		return fmt.Errorf("AutoShape stroke join is unsupported")
	}
	return nil
}

func parseNativeMutationSlide(data []byte, part string) (*nativeMutationSlide, error) {
	root, err := parseNativeXML(data, part)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: parse slide %q: %w", part, err)
	}
	var dialect nativeExtractDialect
	switch root.Name {
	case xml.Name{Space: nsPresentationTransitional, Local: "sld"}:
		dialect = nativeExtractDialect{presentation: nsPresentationTransitional, drawing: nsDrawingTransitional, rels: nsOfficeRelsTransitional}
	case xml.Name{Space: nsPresentationStrict, Local: "sld"}:
		dialect = nativeExtractDialect{presentation: nsPresentationStrict, drawing: nsDrawingStrict, rels: nsOfficeRelsStrict}
	default:
		return nil, fmt.Errorf("pptxpatch: native mutations: slide %q has unsupported namespace/root", part)
	}
	cSld, err := nativeSingleton(root, dialect.presentation, "cSld", true)
	if err != nil {
		return nil, err
	}
	spTree, err := nativeSingleton(cSld, dialect.presentation, "spTree", true)
	if err != nil {
		return nil, err
	}
	objectIDs, err := collectNativeShapeTreeObjectIDs(spTree, dialect)
	if err != nil {
		return nil, err
	}
	shapeNodes := map[string]*nativeXMLNode{}
	for node, objectID := range objectIDs {
		if node.Name == (xml.Name{Space: dialect.presentation, Local: "sp"}) {
			shapeNodes[objectID] = node
		}
	}
	return &nativeMutationSlide{data: data, dialect: dialect, shapeNodes: shapeNodes}, nil
}

func nativeTextParagraphReplacement(shape *nativeXMLNode, dialect nativeExtractDialect, paragraphs []NativeParagraph) (nativeXMLReplacement, error) {
	txBody, err := nativeSingleton(shape, dialect.presentation, "txBody", true)
	if err != nil {
		return nativeXMLReplacement{}, fmt.Errorf("target lacks one exact text body: %w", err)
	}
	paragraphNodes := nativeChildren(txBody, dialect.drawing, "p")
	encoded, err := encodeNativeParagraphs(paragraphs, dialect)
	if err != nil {
		return nativeXMLReplacement{}, err
	}
	if len(paragraphNodes) != 0 {
		return nativeXMLReplacement{start: paragraphNodes[0].RawStart, end: paragraphNodes[len(paragraphNodes)-1].RawEnd, data: encoded}, nil
	}
	lstStyle, err := nativeSingleton(txBody, dialect.drawing, "lstStyle", true)
	if err != nil {
		return nativeXMLReplacement{}, err
	}
	return nativeXMLReplacement{start: lstStyle.RawEnd, end: lstStyle.RawEnd, data: encoded}, nil
}

func encodeNativeParagraphs(paragraphs []NativeParagraph, dialect nativeExtractDialect) ([]byte, error) {
	var output bytes.Buffer
	for _, paragraph := range paragraphs {
		align := map[NativeTextAlign]string{NativeTextAlignLeft: "l", NativeTextAlignCenter: "ctr", NativeTextAlignRight: "r"}[*paragraph.Align]
		fmt.Fprintf(&output, `<a:p xmlns:a=%q><a:pPr algn=%q lvl=%q><a:buNone/></a:pPr>`, dialect.drawing, align, strconv.FormatInt(*paragraph.Level, 10))
		for _, run := range paragraph.Runs {
			font, err := nativeEscapeXML(*run.FontFamily)
			if err != nil {
				return nil, err
			}
			text, err := nativeEscapeXML(*run.Text)
			if err != nil {
				return nil, err
			}
			bold, italic := "0", "0"
			if *run.Bold {
				bold = "1"
			}
			if *run.Italic {
				italic = "1"
			}
			fmt.Fprintf(&output, `<a:r><a:rPr b=%q i=%q sz=%q><a:latin typeface="%s"/><a:solidFill><a:srgbClr val=%q/></a:solidFill></a:rPr>`, bold, italic, strconv.FormatInt(*run.FontSizeHundredthPt, 10), font, *run.Color)
			if strings.TrimSpace(*run.Text) != *run.Text {
				output.WriteString(`<a:t xml:space="preserve">`)
			} else {
				output.WriteString(`<a:t>`)
			}
			output.WriteString(text)
			output.WriteString(`</a:t></a:r>`)
		}
		output.WriteString(`</a:p>`)
		if output.Len() > nativeExtractMaxXMLBytes {
			return nil, fmt.Errorf("encoded text exceeds the native slide XML resource bound")
		}
	}
	return output.Bytes(), nil
}

func nativeEscapeXML(value string) (string, error) {
	var output bytes.Buffer
	if err := xml.EscapeText(&output, []byte(value)); err != nil {
		return "", err
	}
	return output.String(), nil
}

func nativeAutoShapePropertyReplacement(shape *nativeXMLNode, dialect nativeExtractDialect, value NativePPTXAutoShapeMutation) (nativeXMLReplacement, error) {
	spPr, err := nativeSingleton(shape, dialect.presentation, "spPr", true)
	if err != nil {
		return nativeXMLReplacement{}, fmt.Errorf("target lacks one exact AutoShape property node: %w", err)
	}
	encoded := encodeNativeAutoShapeProperties(value, dialect)
	return nativeXMLReplacement{start: spPr.RawStart, end: spPr.RawEnd, data: encoded}, nil
}

func encodeNativeAutoShapeProperties(shape NativePPTXAutoShapeMutation, dialect nativeExtractDialect) []byte {
	var output strings.Builder
	fmt.Fprintf(&output, `<p:spPr xmlns:p=%q xmlns:a=%q><a:xfrm><a:off x=%q y=%q/><a:ext cx=%q cy=%q/></a:xfrm><a:prstGeom prst=%q><a:avLst/></a:prstGeom>`, dialect.presentation, dialect.drawing, strconv.FormatInt(*shape.Transform.X, 10), strconv.FormatInt(*shape.Transform.Y, 10), strconv.FormatInt(*shape.Transform.Cx, 10), strconv.FormatInt(*shape.Transform.Cy, 10), string(shape.Preset))
	if shape.Fill == nil {
		output.WriteString(`<a:noFill/>`)
	} else {
		fmt.Fprintf(&output, `<a:solidFill><a:srgbClr val=%q/></a:solidFill>`, *shape.Fill)
	}
	if shape.Stroke == nil {
		output.WriteString(`<a:ln w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/><a:prstDash val="solid"/><a:round/></a:ln>`)
	} else {
		capValue := map[NativeStrokeCap]string{NativeStrokeCapFlat: "flat", NativeStrokeCapRound: "rnd", NativeStrokeCapSquare: "sq"}[*shape.Stroke.Cap]
		fmt.Fprintf(&output, `<a:ln w=%q cap=%q cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val=%q/></a:solidFill><a:prstDash val="solid"/>`, strconv.FormatInt(*shape.Stroke.WidthEMU, 10), capValue, shape.Stroke.Color)
		switch *shape.Stroke.Join {
		case NativeStrokeJoinRound:
			output.WriteString(`<a:round/>`)
		case NativeStrokeJoinBevel:
			output.WriteString(`<a:bevel/>`)
		case NativeStrokeJoinMiter:
			fmt.Fprintf(&output, `<a:miter lim=%q/>`, strconv.FormatInt(*shape.Stroke.MiterLimit, 10))
		}
		output.WriteString(`</a:ln>`)
	}
	output.WriteString(`</p:spPr>`)
	return []byte(output.String())
}

func applyNativeXMLReplacements(source []byte, replacements []nativeXMLReplacement) ([]byte, error) {
	sort.Slice(replacements, func(left, right int) bool { return replacements[left].start < replacements[right].start })
	previousEnd := int64(0)
	resultSize := int64(len(source))
	for _, replacement := range replacements {
		if replacement.start < previousEnd || replacement.start < 0 || replacement.end < replacement.start || replacement.end > int64(len(source)) {
			return nil, fmt.Errorf("overlapping or invalid source spans")
		}
		resultSize += int64(len(replacement.data)) - (replacement.end - replacement.start)
		if resultSize <= 0 || resultSize > nativeExtractMaxXMLBytes {
			return nil, fmt.Errorf("result exceeds the native slide XML resource bound")
		}
		previousEnd = replacement.end
	}
	var output bytes.Buffer
	output.Grow(int(resultSize))
	cursor := int64(0)
	for _, replacement := range replacements {
		output.Write(source[cursor:replacement.start])
		output.Write(replacement.data)
		cursor = replacement.end
	}
	output.Write(source[cursor:])
	return output.Bytes(), nil
}

func applyNativePPTXPartReplacements(orig []byte, replacements map[string][]byte) ([]byte, error) {
	if len(replacements) == 0 {
		return bytes.Clone(orig), nil
	}
	zr, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: open original archive: %w", err)
	}
	found := map[string]bool{}
	var output bytes.Buffer
	zw := zip.NewWriter(&output)
	if err := zw.SetComment(zr.Comment); err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: preserve archive comment: %w", err)
	}
	for _, file := range zr.File {
		replacement, changed := replacements[file.Name]
		if !changed {
			if err := zw.Copy(file); err != nil {
				return nil, fmt.Errorf("pptxpatch: native mutations: raw-copy %q: %w", file.Name, err)
			}
			continue
		}
		found[file.Name] = true
		writer, err := zw.CreateHeader(&zip.FileHeader{Name: file.Name, Method: zip.Deflate})
		if err != nil {
			return nil, fmt.Errorf("pptxpatch: native mutations: replace %q: %w", file.Name, err)
		}
		if _, err := writer.Write(replacement); err != nil {
			return nil, fmt.Errorf("pptxpatch: native mutations: write %q: %w", file.Name, err)
		}
	}
	for part := range replacements {
		if !found[part] {
			return nil, fmt.Errorf("pptxpatch: native mutations: replacement part %q is missing", part)
		}
	}
	if err := zw.Close(); err != nil {
		return nil, fmt.Errorf("pptxpatch: native mutations: finalize archive: %w", err)
	}
	if err := verifyNativePPTXParts(orig, output.Bytes(), replacements); err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}

func verifyNativePPTXParts(orig, produced []byte, replacements map[string][]byte) error {
	original, err := openNativeExtractPackage(orig)
	if err != nil {
		return fmt.Errorf("pptxpatch: native mutations: verify original OPC: %w", err)
	}
	updated, err := openNativeExtractPackage(produced)
	if err != nil {
		return fmt.Errorf("pptxpatch: native mutations: verify produced OPC: %w", err)
	}
	if len(original.parts) != len(updated.parts) {
		return fmt.Errorf("pptxpatch: native mutations: produced package part count changed")
	}
	for part, originalBytes := range original.parts {
		updatedBytes, ok := updated.parts[part]
		if !ok {
			return fmt.Errorf("pptxpatch: native mutations: package part %q was lost", part)
		}
		if replacement, changed := replacements[part]; changed {
			if !bytes.Equal(updatedBytes, replacement) {
				return fmt.Errorf("pptxpatch: native mutations: replaced part %q does not match requested bytes", part)
			}
		} else if !bytes.Equal(updatedBytes, originalBytes) {
			return fmt.Errorf("pptxpatch: native mutations: untouched part %q changed", part)
		}
	}
	if err := verifyNativePPTXRawEntries(orig, produced, replacements); err != nil {
		return err
	}
	return nil
}

func verifyNativePPTXRawEntries(orig, produced []byte, replacements map[string][]byte) error {
	original, err := zip.NewReader(bytes.NewReader(orig), int64(len(orig)))
	if err != nil {
		return fmt.Errorf("pptxpatch: native mutations: raw verify original: %w", err)
	}
	updated, err := zip.NewReader(bytes.NewReader(produced), int64(len(produced)))
	if err != nil {
		return fmt.Errorf("pptxpatch: native mutations: raw verify produced: %w", err)
	}
	if original.Comment != updated.Comment {
		return fmt.Errorf("pptxpatch: native mutations: archive comment changed")
	}
	updatedByName := make(map[string]*zip.File, len(updated.File))
	for _, file := range updated.File {
		if _, duplicate := updatedByName[file.Name]; duplicate {
			return fmt.Errorf("pptxpatch: native mutations: raw verify duplicate part %q", file.Name)
		}
		updatedByName[file.Name] = file
	}
	for _, originalFile := range original.File {
		if _, changed := replacements[originalFile.Name]; changed {
			continue
		}
		updatedFile := updatedByName[originalFile.Name]
		if updatedFile == nil || originalFile.Method != updatedFile.Method || originalFile.Flags != updatedFile.Flags || originalFile.CRC32 != updatedFile.CRC32 || originalFile.CompressedSize64 != updatedFile.CompressedSize64 || originalFile.UncompressedSize64 != updatedFile.UncompressedSize64 || !originalFile.Modified.Equal(updatedFile.Modified) || originalFile.ModifiedTime != updatedFile.ModifiedTime || originalFile.ModifiedDate != updatedFile.ModifiedDate || originalFile.CreatorVersion != updatedFile.CreatorVersion || originalFile.ReaderVersion != updatedFile.ReaderVersion || originalFile.NonUTF8 != updatedFile.NonUTF8 || originalFile.Comment != updatedFile.Comment || originalFile.ExternalAttrs != updatedFile.ExternalAttrs || !bytes.Equal(originalFile.Extra, updatedFile.Extra) {
			return fmt.Errorf("pptxpatch: native mutations: untouched raw part metadata %q changed", originalFile.Name)
		}
		originalRaw, err := originalFile.OpenRaw()
		if err != nil {
			return fmt.Errorf("pptxpatch: native mutations: raw verify open original %q: %w", originalFile.Name, err)
		}
		updatedRaw, err := updatedFile.OpenRaw()
		if err != nil {
			return fmt.Errorf("pptxpatch: native mutations: raw verify open produced %q: %w", originalFile.Name, err)
		}
		originalDigest, err := nativeRawZIPDigest(originalRaw, originalFile.CompressedSize64)
		if err != nil {
			return fmt.Errorf("pptxpatch: native mutations: raw verify read original %q: %w", originalFile.Name, err)
		}
		updatedDigest, err := nativeRawZIPDigest(updatedRaw, updatedFile.CompressedSize64)
		if err != nil {
			return fmt.Errorf("pptxpatch: native mutations: raw verify read produced %q: %w", originalFile.Name, err)
		}
		if originalDigest != updatedDigest {
			return fmt.Errorf("pptxpatch: native mutations: untouched raw compressed bytes %q changed", originalFile.Name)
		}
	}
	return nil
}

func nativeRawZIPDigest(reader io.Reader, expected uint64) ([sha256.Size]byte, error) {
	if expected > NativePPTXMaxPackageBytes {
		return [sha256.Size]byte{}, fmt.Errorf("compressed entry exceeds package resource bound")
	}
	hash := sha256.New()
	written, err := io.Copy(hash, io.LimitReader(reader, int64(expected)+1))
	if err != nil {
		return [sha256.Size]byte{}, err
	}
	if written != int64(expected) {
		return [sha256.Size]byte{}, fmt.Errorf("compressed entry length mismatch")
	}
	var digest [sha256.Size]byte
	copy(digest[:], hash.Sum(nil))
	return digest, nil
}

func verifyNativePPTXMutationResult(before, after NativePPTXDeck, operations []nativePPTXResolvedMutation) error {
	if issues := ValidateNativePPTX(after); len(issues) != 0 {
		return NativeContractValidationError{Issues: issues}
	}
	beforeElements, err := indexNativePPTXElements(before)
	if err != nil {
		return fmt.Errorf("pptxpatch: native mutations: verify source topology: %w", err)
	}
	afterElements, err := indexNativePPTXElements(after)
	if err != nil {
		return fmt.Errorf("pptxpatch: native mutations: verify produced topology: %w", err)
	}
	if before.SourceRevision == nil || after.SourceRevision == nil || *before.SourceRevision == *after.SourceRevision {
		return fmt.Errorf("pptxpatch: native mutations: produced native revision did not advance")
	}
	targets := make(map[string]NativePPTXMutationKind, len(operations))
	for _, mutation := range operations {
		if _, duplicate := targets[mutation.element.ID]; duplicate {
			return fmt.Errorf("pptxpatch: native mutations: verify duplicate target id %q", mutation.element.ID)
		}
		targets[mutation.element.ID] = mutation.operation.Kind
	}
	if err := verifyNativePPTXContractPreservation(before, after, targets); err != nil {
		return err
	}
	for _, mutation := range operations {
		actual, ok := afterElements[mutation.element.ID]
		if !ok || actual.Source == nil || actual.Source.PartName != mutation.part || actual.Source.ObjectID != mutation.objectID {
			return fmt.Errorf("pptxpatch: native mutations: verify operation %q lost source identity", mutation.operation.OperationID)
		}
		if err := verifyNativePPTXTargetPreservation(mutation.element, actual, mutation.operation.Kind); err != nil {
			return fmt.Errorf("pptxpatch: native mutations: verify operation %q: %w", mutation.operation.OperationID, err)
		}
		switch mutation.operation.Kind {
		case NativePPTXReplaceText:
			if actual.Paragraphs == nil || !nativeParagraphsEqual(*actual.Paragraphs, *mutation.operation.Paragraphs) {
				return fmt.Errorf("pptxpatch: native mutations: verify operation %q text mismatch", mutation.operation.OperationID)
			}
		case NativePPTXUpdateAutoShape:
			if !nativeAutoShapeEquals(actual, *mutation.operation.AutoShape) {
				return fmt.Errorf("pptxpatch: native mutations: verify operation %q AutoShape mismatch", mutation.operation.OperationID)
			}
		}
	}
	if len(afterElements) != len(beforeElements) {
		return fmt.Errorf("pptxpatch: native mutations: verify element inventory changed")
	}
	return nil
}

func verifyNativePPTXContractPreservation(before, after NativePPTXDeck, targets map[string]NativePPTXMutationKind) error {
	if len(before.Assets) != len(after.Assets) {
		return fmt.Errorf("pptxpatch: native mutations: asset topology changed")
	}
	for index := range before.Assets {
		left, right := before.Assets[index], after.Assets[index]
		if err := verifyNativePPTXSourceAnchor(left.Source, right.Source, false); err != nil {
			return fmt.Errorf("pptxpatch: native mutations: asset %d source topology changed: %w", index, err)
		}
		if !nativePassthroughInventoryEqual(left.Passthrough, right.Passthrough, left.Source, right.Source) {
			return fmt.Errorf("pptxpatch: native mutations: asset %q unsupported inventory changed", left.ID)
		}
		expected := left
		expected.Source = right.Source
		expected.Passthrough = right.Passthrough
		if !reflect.DeepEqual(expected, right) {
			return fmt.Errorf("pptxpatch: native mutations: asset %q semantics changed", left.ID)
		}
	}

	if len(before.Slides) != len(after.Slides) {
		return fmt.Errorf("pptxpatch: native mutations: slide topology changed from %d to %d slides", len(before.Slides), len(after.Slides))
	}
	for slideIndex := range before.Slides {
		left, right := before.Slides[slideIndex], after.Slides[slideIndex]
		partTouched := nativePPTXElementListTouched(left.Elements, targets)
		if err := verifyNativePPTXSourceAnchor(left.Source, right.Source, partTouched); err != nil {
			return fmt.Errorf("pptxpatch: native mutations: slide %d source topology changed: %w", slideIndex, err)
		}
		if !nativePassthroughInventoryEqual(left.Passthrough, right.Passthrough, left.Source, right.Source) {
			return fmt.Errorf("pptxpatch: native mutations: slide %q unsupported inventory changed", left.ID)
		}
		if !reflect.DeepEqual(left.Compatibility, right.Compatibility) {
			return fmt.Errorf("pptxpatch: native mutations: slide %q compatibility inventory changed", left.ID)
		}
		if err := verifyNativePPTXElementListPreservation(left.Elements, right.Elements, targets, fmt.Sprintf("slide %q", left.ID)); err != nil {
			return fmt.Errorf("pptxpatch: native mutations: %w", err)
		}
		expectedSlide := left
		expectedSlide.Source = right.Source
		expectedSlide.Passthrough = right.Passthrough
		expectedSlide.Elements = right.Elements
		if !reflect.DeepEqual(expectedSlide, right) {
			return fmt.Errorf("pptxpatch: native mutations: slide %q non-element semantics changed", left.ID)
		}
	}

	expectedDeck := before
	expectedDeck.SourceRevision = after.SourceRevision
	expectedDeck.Assets = after.Assets
	expectedDeck.Slides = after.Slides
	if !reflect.DeepEqual(expectedDeck, after) {
		return fmt.Errorf("pptxpatch: native mutations: presentation semantics or unsupported inventory changed")
	}
	return nil
}

func nativePPTXElementListTouched(elements []NativeElement, targets map[string]NativePPTXMutationKind) bool {
	for _, element := range elements {
		if _, touched := targets[element.ID]; touched || nativePPTXElementListTouched(element.Children, targets) {
			return true
		}
	}
	return false
}

func verifyNativePPTXElementListPreservation(before, after []NativeElement, targets map[string]NativePPTXMutationKind, parent string) error {
	if len(before) != len(after) {
		return fmt.Errorf("%s element topology changed from %d to %d elements", parent, len(before), len(after))
	}
	for index := range before {
		left, right := before[index], after[index]
		kind, touched := targets[left.ID]
		subtreeTouched := touched || nativePPTXElementListTouched(left.Children, targets)
		if err := verifyNativePPTXSourceAnchor(left.Source, right.Source, subtreeTouched); err != nil {
			return fmt.Errorf("element %d under %s source topology changed: %w", index, parent, err)
		}
		if !nativePassthroughSemanticsEqual(left, right) {
			return fmt.Errorf("element %q unsupported inventory changed", left.ID)
		}
		if !reflect.DeepEqual(left.Compatibility, right.Compatibility) {
			return fmt.Errorf("element %q compatibility inventory changed", left.ID)
		}
		if err := verifyNativePPTXElementListPreservation(left.Children, right.Children, targets, fmt.Sprintf("element %q", left.ID)); err != nil {
			return err
		}
		if touched {
			if err := verifyNativePPTXTargetPreservation(left, right, kind); err != nil {
				return fmt.Errorf("element %q target preservation: %w", left.ID, err)
			}
			continue
		}
		expected := left
		expected.Source = right.Source
		expected.Passthrough = right.Passthrough
		// A targeted descendant changes the enclosing p:grpSp fingerprint and
		// child value, but never the group's own transform, child coordinate
		// space, compatibility, or unsupported inventory.
		expected.Children = right.Children
		if !reflect.DeepEqual(expected, right) {
			return fmt.Errorf("untouched element %q semantics changed", left.ID)
		}
	}
	return nil
}

func verifyNativePPTXSourceAnchor(before, after *NativeSourceAnchor, allowFingerprintChange bool) error {
	if before == nil || after == nil {
		if before == nil && after == nil {
			return nil
		}
		return fmt.Errorf("source anchor presence changed")
	}
	if before.PartName != after.PartName || before.ObjectID != after.ObjectID || !nativeStringPointerEqual(before.RelationshipID, after.RelationshipID) {
		return fmt.Errorf("part-qualified source identity changed")
	}
	if !allowFingerprintChange && before.FingerprintSHA256 != after.FingerprintSHA256 {
		return fmt.Errorf("untouched source fingerprint changed")
	}
	return nil
}

func nativePassthroughInventoryEqual(before, after []NativePassthroughRef, beforeSource, afterSource *NativeSourceAnchor) bool {
	if len(before) != len(after) {
		return false
	}
	for index := range before {
		left, right := before[index], after[index]
		if left.OwnerPart != right.OwnerPart || left.Disposition != right.Disposition {
			return false
		}
		expectedFingerprint := left.FingerprintSHA256
		if beforeSource != nil && afterSource != nil && left.FingerprintSHA256 == beforeSource.FingerprintSHA256 {
			expectedFingerprint = afterSource.FingerprintSHA256
		}
		if right.FingerprintSHA256 != expectedFingerprint {
			return false
		}
	}
	return true
}

func verifyNativePPTXTargetPreservation(before, after NativeElement, kind NativePPTXMutationKind) error {
	if before.Source == nil || after.Source == nil || before.Source.PartName != after.Source.PartName || before.Source.ObjectID != after.Source.ObjectID || !nativeStringPointerEqual(before.Source.RelationshipID, after.Source.RelationshipID) {
		return fmt.Errorf("source anchor semantics changed")
	}
	if !reflect.DeepEqual(before.Compatibility, after.Compatibility) {
		return fmt.Errorf("compatibility semantics changed")
	}
	if !nativePassthroughSemanticsEqual(before, after) {
		return fmt.Errorf("passthrough semantics changed")
	}

	expected := before
	expected.Source = after.Source
	expected.Passthrough = after.Passthrough
	expected.Compatibility = after.Compatibility
	switch kind {
	case NativePPTXReplaceText:
		expected.Paragraphs = after.Paragraphs
	case NativePPTXUpdateAutoShape:
		expected.Transform = after.Transform
		expected.Preset = after.Preset
		expected.Fill = after.Fill
		expected.Stroke = after.Stroke
	default:
		return fmt.Errorf("unsupported verification kind %q", kind)
	}
	if !reflect.DeepEqual(expected, after) {
		return fmt.Errorf("non-requested target fields changed")
	}
	return nil
}

func nativePassthroughSemanticsEqual(before, after NativeElement) bool {
	return nativePassthroughInventoryEqual(before.Passthrough, after.Passthrough, before.Source, after.Source)
}

func nativeParagraphsEqual(left, right []NativeParagraph) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		if !nativeParagraphEqual(left[index], right[index]) {
			return false
		}
	}
	return true
}

func nativeParagraphEqual(left, right NativeParagraph) bool {
	if !nativeTextAlignPointerEqual(left.Align, right.Align) || !nativeInt64PointerEqual(left.Level, right.Level) || !nativeBoolPointerEqual(left.Bullet, right.Bullet) || len(left.Runs) != len(right.Runs) {
		return false
	}
	for index := range left.Runs {
		l, r := left.Runs[index], right.Runs[index]
		if !nativeStringPointerEqual(l.Text, r.Text) || !nativeBoolPointerEqual(l.Bold, r.Bold) || !nativeBoolPointerEqual(l.Italic, r.Italic) || !nativeInt64PointerEqual(l.FontSizeHundredthPt, r.FontSizeHundredthPt) || !nativeStringPointerEqual(l.Color, r.Color) || !nativeStringPointerEqual(l.FontFamily, r.FontFamily) {
			return false
		}
	}
	return true
}

func nativeAutoShapeEquals(element NativeElement, expected NativePPTXAutoShapeMutation) bool {
	if element.Preset == nil || *element.Preset != expected.Preset || !nativeTransformEqual(element.Transform, expected.Transform) || !nativeStringPointerEqual(element.Fill, expected.Fill) {
		return false
	}
	return nativeStrokeEqual(element.Stroke, expected.Stroke)
}

func nativeTransformEqual(left, right NativeTransform) bool {
	return nativeInt64PointerEqual(left.X, right.X) && nativeInt64PointerEqual(left.Y, right.Y) && nativeInt64PointerEqual(left.Cx, right.Cx) && nativeInt64PointerEqual(left.Cy, right.Cy) && nativeInt64PointerEqual(left.QuarterTurns, right.QuarterTurns)
}

func nativeStrokeEqual(left, right *NativeStroke) bool {
	if left == nil || right == nil {
		return left == nil && right == nil
	}
	return left.Color == right.Color && nativeInt64PointerEqual(left.WidthEMU, right.WidthEMU) && nativeStrokeCapPointerEqual(left.Cap, right.Cap) && nativeStrokeJoinPointerEqual(left.Join, right.Join) && nativeStrokeDashPointerEqual(left.Dash, right.Dash) && nativeInt64PointerEqual(left.MiterLimit, right.MiterLimit)
}

func nativeStringPointerEqual(left, right *string) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}

func nativeInt64PointerEqual(left, right *int64) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}

func nativeBoolPointerEqual(left, right *bool) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}

func nativeTextAlignPointerEqual(left, right *NativeTextAlign) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}

func nativeStrokeCapPointerEqual(left, right *NativeStrokeCap) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}

func nativeStrokeJoinPointerEqual(left, right *NativeStrokeJoin) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}

func nativeStrokeDashPointerEqual(left, right *NativeStrokeDash) bool {
	return (left == nil && right == nil) || (left != nil && right != nil && *left == *right)
}
