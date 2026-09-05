package officecompat_test

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"errors"
	"io"
	"os"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func TestApplyMutationCanonicalTypeScriptDOCXAdapterEnvelopeAtomically(t *testing.T) {
	source, err := os.ReadFile("corpus/generated/packages/docx-strict-relocated.docx")
	if err != nil {
		t.Fatal(err)
	}
	vectorBytes, err := os.ReadFile("../../testdata/docx-native/transaction-adapter-envelope-v1.json")
	if err != nil {
		t.Fatal(err)
	}
	var vector struct {
		ExpectedText    string `json:"expected_text"`
		EncodedEnvelope string `json:"encoded_envelope"`
	}
	if err := json.Unmarshal(vectorBytes, &vector); err != nil {
		t.Fatal(err)
	}
	before, err := docxpatch.ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	beforeRun := before.Body.Blocks[0].Paragraph.Runs[0]
	var nativeResult *docxpatch.NativeDOCXMutationResultV1
	result, err := officecompat.ApplyMutation(source, []byte(vector.EncodedEnvelope), officecompat.MutationPolicy{
		Format: officecompat.FormatDOCX,
		Apply: func(candidateSource []byte, payload json.RawMessage) (officecompat.Candidate, error) {
			applied, applyErr := docxpatch.ApplyNativeTextMutationPayloadV1(candidateSource, payload, officecompat.ExactRevision(candidateSource))
			if applyErr != nil {
				return officecompat.Candidate{}, applyErr
			}
			nativeResult = applied
			parts := make([]string, len(applied.Evidence.ChangedParts))
			for index, part := range applied.Evidence.ChangedParts {
				parts[index] = part.PartName
			}
			return officecompat.Candidate{Bytes: applied.Package, MutableParts: parts}, nil
		},
		ReopenValidate: func(candidate []byte, _ json.RawMessage) error {
			reopened, reopenErr := docxpatch.ExtractNativeDocumentV1(candidate)
			if reopenErr != nil {
				return reopenErr
			}
			if issues := docxpatch.ValidateNativeDocumentV1(reopened); len(issues) > 0 {
				return &docxpatch.NativeValidationError{Issues: issues}
			}
			return nil
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if nativeResult == nil || result.MutationID != "pm-save-1" || result.PreviousRevision != before.Source.PackageSHA256 {
		t.Fatalf("adapter envelope did not retain exact transaction provenance: result=%+v native=%#v", result, nativeResult)
	}
	afterRun := nativeResult.Document.Body.Blocks[0].Paragraph.Runs[0]
	if afterRun.Text == nil || *afterRun.Text != vector.ExpectedText {
		t.Fatalf("adapter mutation text = %#v, want %q", afterRun.Text, vector.ExpectedText)
	}
	if nativeResult.Document.DocumentID != before.DocumentID || nativeResult.Document.Body.Blocks[0].Paragraph.ID != before.Body.Blocks[0].Paragraph.ID || afterRun.ID != beforeRun.ID {
		t.Fatal("durable native document, paragraph, or run identity changed")
	}
	if len(result.MutableParts) != 1 || result.MutableParts[0] != "Odd/Main.XML" || nativeResult.Evidence.UntouchedPartsVerified == 0 {
		t.Fatalf("unexpected preservation evidence: result=%+v evidence=%+v", result, nativeResult.Evidence)
	}
	if err := officecompat.RequireUntouchedParts(source, result.Bytes, []string{"Odd/Main.XML"}); err != nil {
		t.Fatalf("adapter mutation escaped the authoritative native part: %v", err)
	}
}

func TestApplyMutationEnforcesExactRevisionPreservationAndReopen(t *testing.T) {
	original := minimalXLSX(t)
	originalCopy := bytes.Clone(original)
	request := mutationEnvelope(t, officecompat.FormatXLSX, original, map[string]any{"operation": "cell.set_value"})
	var reopened bool
	var handlerCandidate []byte
	result, err := officecompat.ApplyMutation(original, request, officecompat.MutationPolicy{
		Format: officecompat.FormatXLSX,
		Apply: func(source []byte, payload json.RawMessage) (officecompat.Candidate, error) {
			source[0] ^= 0xff
			source = bytes.Clone(original)
			updated := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>after</t></is></c></row></sheetData></worksheet>`)
			candidate, applyErr := xlsxpatch.Apply(source, xlsxpatch.Patch{Replace: map[string][]byte{"xl/worksheets/sheet1.xml": updated}})
			handlerCandidate = candidate
			return officecompat.Candidate{Bytes: candidate, MutableParts: []string{"xl/worksheets/sheet1.xml"}}, applyErr
		},
		ReopenValidate: func(candidate []byte, payload json.RawMessage) error {
			reopened = true
			if _, extractErr := xlsxpatch.ExtractNativeWorkbookV1(candidate); extractErr != nil {
				return extractErr
			}
			candidate[0] ^= 0xff
			return nil
		},
	})
	if err != nil {
		t.Fatal(err)
	}
	if !reopened || !bytes.Equal(original, originalCopy) {
		t.Fatal("policy did not reopen candidate or mutated caller-owned source")
	}
	if result.PreviousRevision != officecompat.ExactRevision(original) || result.Revision != officecompat.ExactRevision(result.Bytes) || result.Revision == result.PreviousRevision {
		t.Fatalf("unexpected exact revisions: %+v", result)
	}
	if result.MutationID != "save-1" || len(result.MutableParts) != 1 || result.MutableParts[0] != "xl/worksheets/sheet1.xml" {
		t.Fatalf("unexpected result metadata: %+v", result)
	}
	if !bytes.Equal(result.Bytes, handlerCandidate) {
		t.Fatal("validator mutation escaped its isolated input")
	}
	handlerCandidate[0] ^= 0xff
	if bytes.Equal(result.Bytes, handlerCandidate) {
		t.Fatal("result aliases handler-owned candidate bytes")
	}
}

func TestApplyMutationRejectsStaleRevisionBeforeNativeHandler(t *testing.T) {
	original := minimalXLSX(t)
	request := mutationEnvelope(t, officecompat.FormatXLSX, []byte("different source"), map[string]any{"operation": "cell.set_value"})
	called := false
	result, err := officecompat.ApplyMutation(original, request, officecompat.MutationPolicy{
		Format: officecompat.FormatXLSX,
		Apply: func([]byte, json.RawMessage) (officecompat.Candidate, error) {
			called = true
			return officecompat.Candidate{}, nil
		},
		ReopenValidate: func([]byte, json.RawMessage) error { called = true; return nil },
	})
	assertMutationError(t, err, officecompat.MutationStaleRevision)
	if called || result.Bytes != nil {
		t.Fatal("stale mutation reached native handler or returned candidate bytes")
	}
}

func TestApplyMutationFailsClosedOnEscapedPartAndRoundTripFailure(t *testing.T) {
	original := minimalXLSX(t)
	request := mutationEnvelope(t, officecompat.FormatXLSX, original, map[string]any{"operation": "cell.set_value"})
	escaped, err := xlsxpatch.Apply(original, xlsxpatch.Patch{Replace: map[string][]byte{
		"xl/worksheets/sheet1.xml": []byte("changed sheet"),
		"xl/workbook.xml":          []byte("escaped workbook mutation"),
	}})
	if err != nil {
		t.Fatal(err)
	}
	validated := false
	result, err := officecompat.ApplyMutation(original, request, officecompat.MutationPolicy{
		Format: officecompat.FormatXLSX,
		Apply: func([]byte, json.RawMessage) (officecompat.Candidate, error) {
			return officecompat.Candidate{Bytes: escaped, MutableParts: []string{"xl/worksheets/sheet1.xml"}}, nil
		},
		ReopenValidate: func([]byte, json.RawMessage) error { validated = true; return nil },
	})
	assertMutationError(t, err, officecompat.MutationPreservation)
	if validated || result.Bytes != nil {
		t.Fatal("escaped mutation reached validation or returned candidate bytes")
	}

	candidate, err := xlsxpatch.Apply(original, xlsxpatch.Patch{Replace: map[string][]byte{"xl/worksheets/sheet1.xml": []byte("changed")}})
	if err != nil {
		t.Fatal(err)
	}
	result, err = officecompat.ApplyMutation(original, request, officecompat.MutationPolicy{
		Format: officecompat.FormatXLSX,
		Apply: func([]byte, json.RawMessage) (officecompat.Candidate, error) {
			return officecompat.Candidate{Bytes: candidate, MutableParts: []string{"xl/worksheets/sheet1.xml"}}, nil
		},
		ReopenValidate: func([]byte, json.RawMessage) error { return errors.New("native extractor refused candidate") },
	})
	assertMutationError(t, err, officecompat.MutationValidation)
	if result.Bytes != nil {
		t.Fatal("failed round-trip returned candidate bytes")
	}
}

func TestDecodeMutationEnvelopeIsStrictAndBounded(t *testing.T) {
	original := minimalXLSX(t)
	revision := officecompat.ExactRevision(original)
	valid := `{"protocol":"injoffice.office.mutations","version":1,"format":"docx","mutation_id":"m-1","expected_revision":"` + revision + `","payload":{}}`
	if _, err := officecompat.DecodeMutationEnvelope([]byte(valid)); err != nil {
		t.Fatal(err)
	}
	invalid := []string{
		`[]`,
		`{"protocol":"injoffice.office.mutations","protocol":"injoffice.office.mutations","version":1,"format":"docx","mutation_id":"m-1","expected_revision":"` + revision + `","payload":{}}`,
		strings.Replace(valid, `"payload":{}`, `"payload":[],"unknown":true`, 1),
		valid + ` {}`,
		strings.Replace(valid, `"mutation_id":"m-1"`, `"mutation_id":"spaces forbidden"`, 1),
		strings.Replace(valid, `"expected_revision":"`+revision+`"`, `"expected_revision":"rev:short"`, 1),
	}
	for _, encoded := range invalid {
		if _, err := officecompat.DecodeMutationEnvelope([]byte(encoded)); err == nil {
			t.Fatalf("accepted invalid envelope: %s", encoded)
		}
	}
	oversized := bytes.Repeat([]byte{' '}, officecompat.MaxMutationEnvelopeBytes+1)
	if _, err := officecompat.DecodeMutationEnvelope(oversized); err == nil {
		t.Fatal("accepted oversized mutation envelope")
	}
}

func TestApplyMutationRejectsIncompletePolicyNoopAndPatternAllowlist(t *testing.T) {
	original := minimalXLSX(t)
	request := mutationEnvelope(t, officecompat.FormatXLSX, original, map[string]any{"operation": "cell.set_value"})
	if _, err := officecompat.ApplyMutation(original, request, officecompat.MutationPolicy{Format: officecompat.FormatDOCX}); err == nil {
		t.Fatal("accepted mismatched incomplete native policy")
	}
	for _, candidate := range []officecompat.Candidate{
		{Bytes: original, MutableParts: []string{"xl/worksheets/sheet1.xml"}},
		{Bytes: append(bytes.Clone(original), 0), MutableParts: []string{"xl/worksheets/*.xml"}},
	} {
		_, err := officecompat.ApplyMutation(original, request, officecompat.MutationPolicy{
			Format:         officecompat.FormatXLSX,
			Apply:          func([]byte, json.RawMessage) (officecompat.Candidate, error) { return candidate, nil },
			ReopenValidate: func([]byte, json.RawMessage) error { return nil },
		})
		if err == nil {
			t.Fatalf("accepted invalid candidate: %+v", candidate)
		}
	}
}

func TestApplyMutationRejectsZIPMetadataOnlyRewriteBeforeReopen(t *testing.T) {
	original := minimalXLSX(t)
	repacked := repackWithStoredEntries(t, original)
	if bytes.Equal(original, repacked) {
		t.Fatal("test setup did not change exact ZIP bytes")
	}
	left, err := officecompat.Inspect(original)
	if err != nil {
		t.Fatal(err)
	}
	right, err := officecompat.Inspect(repacked)
	if err != nil {
		t.Fatal(err)
	}
	if left.Fingerprint != right.Fingerprint {
		t.Fatal("test setup changed an OPC part payload")
	}

	reopened := false
	result, err := officecompat.ApplyMutation(original, mutationEnvelope(t, officecompat.FormatXLSX, original, map[string]any{"operation": "cell.set_value"}), officecompat.MutationPolicy{
		Format: officecompat.FormatXLSX,
		Apply: func([]byte, json.RawMessage) (officecompat.Candidate, error) {
			return officecompat.Candidate{Bytes: repacked, MutableParts: []string{"xl/worksheets/sheet1.xml"}}, nil
		},
		ReopenValidate: func([]byte, json.RawMessage) error { reopened = true; return nil },
	})
	assertMutationError(t, err, officecompat.MutationApplyFailed)
	if reopened || result.Bytes != nil {
		t.Fatal("container-only rewrite reached native validation or returned candidate bytes")
	}
}

func TestInspectRejectsCompressionBombBeforeExpansion(t *testing.T) {
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	entries := map[string][]byte{
		"[Content_Types].xml": []byte(contentTypes),
		"_rels/.rels":         []byte(emptyRootRels),
		"custom/bomb.bin":     bytes.Repeat([]byte{'A'}, 3<<20),
	}
	for name, payload := range entries {
		part, err := writer.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := part.Write(payload); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := officecompat.Inspect(buffer.Bytes()); err == nil || !strings.Contains(err.Error(), "compression ratio") {
		t.Fatalf("expected bounded compression-ratio refusal, got %v", err)
	}
}

func mutationEnvelope(t *testing.T, format officecompat.Format, source []byte, payload map[string]any) []byte {
	t.Helper()
	encoded, err := json.Marshal(officecompat.MutationEnvelope{
		Protocol: officecompat.MutationProtocol, Version: officecompat.MutationVersion,
		Format: format, MutationID: "save-1", ExpectedRevision: officecompat.ExactRevision(source), Payload: mustJSON(t, payload),
	})
	if err != nil {
		t.Fatal(err)
	}
	return encoded
}

func mustJSON(t *testing.T, value any) json.RawMessage {
	t.Helper()
	encoded, err := json.Marshal(value)
	if err != nil {
		t.Fatal(err)
	}
	return encoded
}

func assertMutationError(t *testing.T, err error, code officecompat.MutationErrorCode) {
	t.Helper()
	var mutationErr *officecompat.MutationError
	if !errors.As(err, &mutationErr) || mutationErr.Code != code {
		t.Fatalf("expected mutation error %s, got %v", code, err)
	}
}

func repackWithStoredEntries(t *testing.T, source []byte) []byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(source), int64(len(source)))
	if err != nil {
		t.Fatal(err)
	}
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	for index := len(reader.File) - 1; index >= 0; index-- {
		file := reader.File[index]
		input, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		payload, readErr := io.ReadAll(input)
		closeErr := input.Close()
		if readErr != nil {
			t.Fatal(readErr)
		}
		if closeErr != nil {
			t.Fatal(closeErr)
		}
		entry, err := writer.CreateHeader(&zip.FileHeader{Name: file.Name, Method: zip.Store})
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write(payload); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return output.Bytes()
}
