// Package artifacthttp mints opaque artifact IDs from raw request bytes.
// IDs are store tokens; they are never treated as filesystem paths.
package artifacthttp

import (
	"encoding/json"
	"errors"
	"io"
	"net/http"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const Path = "/v1/artifacts"

var maxArtifactBytes = int64(max(
	xlsxpatch.NativeXLSXMaxPackageBytes,
	docxpatch.NativeDOCXMaxPackageBytes,
	pptxpatch.NativePPTXMaxPackageBytes,
))

// Handler serves POST /v1/artifacts.
type Handler struct {
	store xlsxhttp.Store
}

// New returns a handler that mints IDs through store.Create.
func New(store xlsxhttp.Store) *Handler {
	return &Handler{store: store}
}

// ServeHTTP accepts opaque bytes and returns a store-minted artifact id.
func (h *Handler) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
	w.Header().Set("Access-Control-Expose-Headers", xlsxhttp.HeaderArtifactID)
	if r.Method == http.MethodOptions {
		w.WriteHeader(http.StatusNoContent)
		return
	}
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	if h.store == nil {
		xlsxhttp.WriteError(w, http.StatusInternalServerError, errors.New("artifact store is not configured"))
		return
	}

	r.Body = http.MaxBytesReader(w, r.Body, maxArtifactBytes)
	data, err := io.ReadAll(r.Body)
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	if len(data) == 0 {
		xlsxhttp.WriteError(w, http.StatusBadRequest, errors.New("empty artifact body"))
		return
	}
	id, err := h.store.Create(data)
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	w.Header().Set(xlsxhttp.HeaderArtifactID, id)
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(http.StatusCreated)
	_ = json.NewEncoder(w).Encode(map[string]string{"artifact_id": id})
}
