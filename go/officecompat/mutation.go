package officecompat

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"regexp"
	"sort"
	"strings"
)

const (
	MutationProtocol         = "injoffice.office.mutations"
	MutationVersion          = 1
	MaxMutationEnvelopeBytes = 4 << 20
	MaxMutationPayloadBytes  = 3 << 20
)

type Format string

const (
	FormatDOCX Format = "docx"
	FormatPPTX Format = "pptx"
	FormatXLSX Format = "xlsx"
)

var (
	mutationIDPattern = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._:-]{0,127}$`)
	revisionPattern   = regexp.MustCompile(`^sha256:[0-9a-f]{64}$`)
)

// MutationEnvelope is renderer-neutral. Payload remains opaque here and must
// be decoded by the selected native format handler's own strict contract.
type MutationEnvelope struct {
	Protocol         string          `json:"protocol"`
	Version          int             `json:"version"`
	Format           Format          `json:"format"`
	MutationID       string          `json:"mutation_id"`
	ExpectedRevision string          `json:"expected_revision"`
	Payload          json.RawMessage `json:"payload"`
}

type MutationErrorCode string

const (
	MutationInvalidEnvelope   MutationErrorCode = "INVALID_ENVELOPE"
	MutationUnsupportedFormat MutationErrorCode = "UNSUPPORTED_FORMAT"
	MutationStaleRevision     MutationErrorCode = "STALE_REVISION"
	MutationApplyFailed       MutationErrorCode = "APPLY_FAILED"
	MutationPreservation      MutationErrorCode = "PRESERVATION_FAILED"
	MutationValidation        MutationErrorCode = "ROUND_TRIP_VALIDATION_FAILED"
)

type MutationError struct {
	Code  MutationErrorCode
	Field string
	Err   error
}

func (err *MutationError) Error() string {
	if err.Field != "" {
		return fmt.Sprintf("officecompat: %s at %s: %v", err.Code, err.Field, err.Err)
	}
	return fmt.Sprintf("officecompat: %s: %v", err.Code, err.Err)
}

func (err *MutationError) Unwrap() error { return err.Err }

// Candidate is a native handler's tentative result. MutableParts is the exact
// OPC-part allowlist for this operation; prefixes and globs are never accepted.
type Candidate struct {
	Bytes        []byte
	MutableParts []string
}

// MutationPolicy binds one envelope format to its native byte mutator and its
// independent reopen/extract/validate step. Both callbacks receive clones.
type MutationPolicy struct {
	Format         Format
	Apply          func(source []byte, payload json.RawMessage) (Candidate, error)
	ReopenValidate func(candidate []byte, payload json.RawMessage) error
}

type MutationResult struct {
	Bytes            []byte
	MutationID       string
	PreviousRevision string
	Revision         string
	MutableParts     []string
}

func ExactRevision(data []byte) string {
	digest := sha256.Sum256(data)
	return "sha256:" + hex.EncodeToString(digest[:])
}

// DecodeMutationEnvelope rejects unknown/duplicate outer fields and trailing
// JSON. The native payload is intentionally not interpreted by this layer.
func DecodeMutationEnvelope(data []byte) (MutationEnvelope, error) {
	invalid := func(field, message string) (MutationEnvelope, error) {
		return MutationEnvelope{}, &MutationError{Code: MutationInvalidEnvelope, Field: field, Err: errors.New(message)}
	}
	if len(data) == 0 || len(data) > MaxMutationEnvelopeBytes {
		return invalid("/", fmt.Sprintf("encoded envelope size must be 1..%d bytes", MaxMutationEnvelopeBytes))
	}

	decoder := json.NewDecoder(bytes.NewReader(data))
	first, err := decoder.Token()
	if err != nil {
		return invalid("/", "must be one JSON object")
	}
	if delimiter, ok := first.(json.Delim); !ok || delimiter != '{' {
		return invalid("/", "must be one JSON object")
	}
	allowed := map[string]bool{
		"protocol": true, "version": true, "format": true,
		"mutation_id": true, "expected_revision": true, "payload": true,
	}
	fields := make(map[string]json.RawMessage, len(allowed))
	for decoder.More() {
		token, tokenErr := decoder.Token()
		if tokenErr != nil {
			return invalid("/", "contains invalid JSON")
		}
		name, ok := token.(string)
		if !ok {
			return invalid("/", "contains a non-string field name")
		}
		if !allowed[name] {
			return invalid("/"+escapeJSONPointer(name), "unknown field")
		}
		if _, duplicate := fields[name]; duplicate {
			return invalid("/"+escapeJSONPointer(name), "duplicate field")
		}
		var raw json.RawMessage
		if decodeErr := decoder.Decode(&raw); decodeErr != nil {
			return invalid("/"+escapeJSONPointer(name), "contains invalid JSON")
		}
		fields[name] = bytes.Clone(raw)
	}
	if _, err := decoder.Token(); err != nil {
		return invalid("/", "contains an unterminated object")
	}
	if token, trailingErr := decoder.Token(); trailingErr != io.EOF || token != nil {
		return invalid("/", "must not contain trailing JSON")
	}

	required := []string{"protocol", "version", "format", "mutation_id", "expected_revision", "payload"}
	for _, name := range required {
		if _, ok := fields[name]; !ok {
			return invalid("/"+name, "field is required")
		}
	}
	var envelope MutationEnvelope
	if err := decodeScalar(fields["protocol"], &envelope.Protocol); err != nil || envelope.Protocol != MutationProtocol {
		return invalid("/protocol", "unsupported protocol")
	}
	if err := decodeScalar(fields["version"], &envelope.Version); err != nil || envelope.Version != MutationVersion {
		return invalid("/version", "unsupported version")
	}
	if err := decodeScalar(fields["format"], &envelope.Format); err != nil || !validFormat(envelope.Format) {
		return invalid("/format", "must be docx, pptx, or xlsx")
	}
	if err := decodeScalar(fields["mutation_id"], &envelope.MutationID); err != nil || !mutationIDPattern.MatchString(envelope.MutationID) {
		return invalid("/mutation_id", "must use 1..128 restricted ASCII identifier characters")
	}
	if err := decodeScalar(fields["expected_revision"], &envelope.ExpectedRevision); err != nil || !revisionPattern.MatchString(envelope.ExpectedRevision) {
		return invalid("/expected_revision", "must be sha256 followed by the full lowercase exact-byte digest")
	}
	rawPayload := fields["payload"]
	payload := bytes.TrimSpace(rawPayload)
	if len(payload) == 0 || len(rawPayload) > MaxMutationPayloadBytes || payload[0] != '{' {
		return invalid("/payload", fmt.Sprintf("must be an object encoded in 1..%d bytes", MaxMutationPayloadBytes))
	}
	envelope.Payload = bytes.Clone(payload)
	return envelope, nil
}

// ApplyMutation enforces the shared transaction order: strict decode, exact
// source CAS, native apply, untouched-part proof, then native reopen/validate.
// No candidate bytes escape on any failure.
func ApplyMutation(source, encodedEnvelope []byte, policy MutationPolicy) (MutationResult, error) {
	envelope, err := DecodeMutationEnvelope(encodedEnvelope)
	if err != nil {
		return MutationResult{}, err
	}
	if policy.Format != envelope.Format || !validFormat(policy.Format) || policy.Apply == nil || policy.ReopenValidate == nil {
		return MutationResult{}, &MutationError{Code: MutationUnsupportedFormat, Field: "/format", Err: errors.New("no complete native policy is registered for envelope format")}
	}
	if len(source) == 0 || len(source) > MaxPackageBytes {
		return MutationResult{}, &MutationError{Code: MutationApplyFailed, Err: fmt.Errorf("source size must be 1..%d bytes", MaxPackageBytes)}
	}
	previousRevision := ExactRevision(source)
	if envelope.ExpectedRevision != previousRevision {
		return MutationResult{}, &MutationError{Code: MutationStaleRevision, Field: "/expected_revision", Err: errors.New("source bytes do not match expected revision")}
	}

	baseline := bytes.Clone(source)
	candidate, err := policy.Apply(bytes.Clone(source), bytes.Clone(envelope.Payload))
	if err != nil {
		return MutationResult{}, &MutationError{Code: MutationApplyFailed, Err: err}
	}
	if len(candidate.Bytes) == 0 || len(candidate.Bytes) > MaxPackageBytes {
		return MutationResult{}, &MutationError{Code: MutationApplyFailed, Err: fmt.Errorf("candidate size must be 1..%d bytes", MaxPackageBytes)}
	}
	if bytes.Equal(baseline, candidate.Bytes) {
		return MutationResult{}, &MutationError{Code: MutationApplyFailed, Err: errors.New("native mutation produced no exact-byte change")}
	}
	mutableParts, err := normalizeMutableParts(candidate.MutableParts)
	if err != nil {
		return MutationResult{}, &MutationError{Code: MutationPreservation, Err: err}
	}
	report, mutableChanged, err := compareDeclaredPartMutation(baseline, candidate.Bytes, mutableParts)
	if err != nil {
		return MutationResult{}, &MutationError{Code: MutationPreservation, Err: err}
	}
	if !report.OK() {
		return MutationResult{}, &MutationError{Code: MutationPreservation, Err: &PreservationError{Report: report}}
	}
	if !mutableChanged {
		return MutationResult{}, &MutationError{Code: MutationApplyFailed, Err: errors.New("native mutation changed only ZIP container bytes, not a declared OPC part payload")}
	}
	if err := policy.ReopenValidate(bytes.Clone(candidate.Bytes), bytes.Clone(envelope.Payload)); err != nil {
		return MutationResult{}, &MutationError{Code: MutationValidation, Err: err}
	}
	return MutationResult{
		Bytes: bytes.Clone(candidate.Bytes), MutationID: envelope.MutationID,
		PreviousRevision: previousRevision, Revision: ExactRevision(candidate.Bytes),
		MutableParts: mutableParts,
	}, nil
}

func decodeScalar(raw json.RawMessage, target any) error {
	decoder := json.NewDecoder(bytes.NewReader(raw))
	if err := decoder.Decode(target); err != nil {
		return err
	}
	if token, err := decoder.Token(); err != io.EOF || token != nil {
		return errors.New("trailing JSON")
	}
	return nil
}

func validFormat(format Format) bool {
	return format == FormatDOCX || format == FormatPPTX || format == FormatXLSX
}

func normalizeMutableParts(parts []string) ([]string, error) {
	if len(parts) == 0 {
		return nil, errors.New("native mutation must declare at least one mutable OPC part")
	}
	normalized := append([]string(nil), parts...)
	sort.Strings(normalized)
	for index, name := range normalized {
		if err := validatePartName(name); err != nil {
			return nil, err
		}
		if strings.ContainsAny(name, "*?[") {
			return nil, fmt.Errorf("mutable OPC part %q must not use pattern syntax", name)
		}
		if index > 0 && normalized[index-1] == name {
			return nil, fmt.Errorf("duplicate mutable OPC part %q", name)
		}
	}
	return normalized, nil
}

func escapeJSONPointer(value string) string {
	return strings.ReplaceAll(strings.ReplaceAll(value, "~", "~0"), "/", "~1")
}
