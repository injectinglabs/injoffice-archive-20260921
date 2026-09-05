package officehttp

import (
	"errors"
	"net/http"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const (
	DOCXExtractPath   = "/v1/docx/extract"
	DOCXMutationsPath = "/v1/docx/mutations"
	DOCXContentType   = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
)

func docxEnvelope() xlsxhttp.Envelope {
	return xlsxhttp.Envelope{
		MaxPackageBytes: int64(docxpatch.NativeDOCXMaxPackageBytes),
		MaxPayloadBytes: int64(docxpatch.NativeDOCXMaxMutationPayloadBytes),
		FileFields:      []string{"original", "file", "docx", "document"},
		FormatName:      "DOCX",
	}
}

func extractDOCXJSON(data []byte) ([]byte, error) {
	doc, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil {
		return nil, err
	}
	return docxpatch.EncodeNativeDocumentV1(doc)
}

func handleDOCXExtract(w http.ResponseWriter, r *http.Request, store xlsxhttp.Store) {
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	data, artifactID, err := xlsxhttp.ReadExtractRequest(w, r, store, docxEnvelope())
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	encoded, err := extractDOCXJSON(data)
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
	w.Header().Set("Content-Type", docxpatch.NativeDOCXV1MediaType)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(encoded)
}

func handleDOCXMutations(w http.ResponseWriter, r *http.Request, store xlsxhttp.Store) {
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	original, payload, outer, artifactID, err := xlsxhttp.ReadMutationRequest(w, r, store, docxEnvelope())
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	result, err := docxpatch.ApplyNativeTextMutationPayloadV1(original, payload, outer)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusBadRequest, err)
		return
	}
	if err := xlsxhttp.PersistMutation(store, artifactID, result.Package, original); err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	if artifactID != "" {
		w.Header().Set(xlsxhttp.HeaderArtifactID, artifactID)
	}
	w.Header().Set("Content-Type", DOCXContentType)
	w.Header().Set("Content-Disposition", `attachment; filename="mutated.docx"`)
	w.Header().Set(xlsxhttp.HeaderRevision, result.Document.Revision)
	w.Header().Set(xlsxhttp.HeaderPackageSHA, result.Document.Source.PackageSHA256)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(result.Package)
}
