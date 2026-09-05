package officehttp

import (
	"crypto/sha256"
	"encoding/hex"
	"errors"
	"net/http"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const (
	PPTXExtractPath   = "/v1/pptx/extract"
	PPTXMutationsPath = "/v1/pptx/mutations"
	PPTXContentType   = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
)

var pptxExtractOptions = pptxpatch.NativePPTXExtractOptions{TokenFactory: pptxHTTPTokenFactory{}}

func pptxEnvelope() xlsxhttp.Envelope {
	return xlsxhttp.Envelope{
		MaxPackageBytes: int64(pptxpatch.NativePPTXMaxPackageBytes),
		MaxPayloadBytes: int64(pptxpatch.MaxNativePPTXMutationPayloadBytes),
		FileFields:      []string{"original", "file", "pptx", "presentation"},
		FormatName:      "PPTX",
	}
}

func extractPPTXJSON(data []byte) ([]byte, error) {
	deck, err := pptxpatch.ExtractNativePPTX(data, pptxExtractOptions)
	if err != nil {
		return nil, err
	}
	return pptxpatch.MarshalNativePPTXJSON(deck)
}

func handlePPTXExtract(w http.ResponseWriter, r *http.Request, store xlsxhttp.Store) {
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	data, artifactID, err := xlsxhttp.ReadExtractRequest(w, r, store, pptxEnvelope())
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	encoded, err := extractPPTXJSON(data)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusBadRequest, err)
		return
	}
	artifactID, err = xlsxhttp.PersistExtract(store, data, artifactID)
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	if artifactID != "" {
		w.Header().Set(xlsxhttp.HeaderArtifactID, artifactID)
	}
	w.Header().Set("Content-Type", pptxpatch.NativePPTXV1MediaType)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(encoded)
}

func handlePPTXMutations(w http.ResponseWriter, r *http.Request, store xlsxhttp.Store) {
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	original, payload, outer, artifactID, err := xlsxhttp.ReadMutationRequest(w, r, store, pptxEnvelope())
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	produced, err := pptxpatch.ApplyNativePPTXMutationPayload(original, outer, payload)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusBadRequest, err)
		return
	}
	if err := xlsxhttp.PersistMutation(store, artifactID, produced, original); err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	if artifactID != "" {
		w.Header().Set(xlsxhttp.HeaderArtifactID, artifactID)
	}
	digest := sha256.Sum256(produced)
	hexDigest := hex.EncodeToString(digest[:])
	w.Header().Set("Content-Type", PPTXContentType)
	w.Header().Set("Content-Disposition", `attachment; filename="mutated.pptx"`)
	w.Header().Set(xlsxhttp.HeaderRevision, "rev-"+hexDigest)
	w.Header().Set(xlsxhttp.HeaderPackageSHA, "sha256:"+hexDigest)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(produced)
}

type pptxHTTPTokenFactory struct{}

func (pptxHTTPTokenFactory) IssueNativePassthroughToken(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
	return pptxHTTPToken(request), nil
}

func (pptxHTTPTokenFactory) BeginNativePassthroughTokenTransaction() (pptxpatch.NativePassthroughTokenTransaction, error) {
	return pptxHTTPTokenTransaction{}, nil
}

type pptxHTTPTokenTransaction struct{}

func (pptxHTTPTokenTransaction) IssueNativePassthroughToken(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
	return pptxHTTPToken(request), nil
}

func (pptxHTTPTokenTransaction) CommitNativePassthroughTokens() error { return nil }

func (pptxHTTPTokenTransaction) RollbackNativePassthroughTokens() {}

func pptxHTTPToken(request pptxpatch.NativePassthroughTokenRequest) string {
	sum := sha256.Sum256([]byte(request.SourceRevision + "\x00" + request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.FingerprintSHA256 + "\x00" + request.Reason))
	return "http-" + hex.EncodeToString(sum[:16])
}
