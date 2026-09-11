// Package xlsxhttp is the shared HTTP surface for native XLSX extract/apply.
// cmd/xlsxnative and injoffice-server both serve these handlers so there is one
// OPC writer and one request shape.
package xlsxhttp

import (
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"mime"
	"net/http"
	"strings"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const (
	ExtractPath        = "/v1/xlsx/extract"
	PreviewObjectsPath = "/v1/xlsx/preview-objects"
	MutationsPath      = "/v1/xlsx/mutations"
	XLSXContentType    = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
	HeaderRevision     = "X-InjOffice-Revision"
	HeaderPackageSHA   = "X-InjOffice-Package-SHA256"
	HeaderArtifactID   = "X-InjOffice-Artifact-Id"
	multipartMemory    = 32 << 20
)

// Store is an optional artifact backend. IDs are opaque tokens minted by
// Create; implementations must not treat an ID as a filesystem path.
// Put is compare-and-swap: it must replace the object only when the current
// bytes still digest to the same value as expected.
type Store interface {
	Get(id string) ([]byte, error)
	Put(id string, data, expected []byte) error
	Create(data []byte) (id string, err error)
}

var (
	ErrArtifactNotFound  = errors.New("artifact not found")
	ErrInvalidArtifactID = errors.New("invalid artifact id")
	ErrArtifactExists    = errors.New("artifact already exists")
	ErrArtifactTooLarge  = errors.New("artifact too large")
	ErrArtifactStale     = errors.New("stale artifact revision")
	ErrArtifactIO        = errors.New("artifact store io")
)

// Envelope is the shared native extract/mutation HTTP request shape.
type Envelope struct {
	MaxPackageBytes int64
	MaxPayloadBytes int64
	FileFields      []string
	FormatName      string
}

// XLSXEnvelope is the helper/server request bounds for /v1/xlsx/*.
func XLSXEnvelope() Envelope {
	return Envelope{
		MaxPackageBytes: int64(xlsxpatch.NativeXLSXMaxPackageBytes),
		MaxPayloadBytes: int64(xlsxpatch.MaxNativeWorkbookMutationPayloadBytes),
		FileFields:      []string{"original", "file", "xlsx", "workbook"},
		FormatName:      "XLSX",
	}
}

// ExtractNativeJSON is the v2 extract used by both the CLI helper and HTTP.
func ExtractNativeJSON(data []byte) ([]byte, error) {
	workbook, err := xlsxpatch.ExtractNativeWorkbookV2(data)
	if err != nil {
		return nil, err
	}
	return xlsxpatch.EncodeNativeWorkbookV2(workbook)
}

// ApplyNativeMutation applies a native mutation payload. An empty outer CAS is
// filled from the original package digest (CLI convenience only).
func ApplyNativeMutation(original, payload []byte, outerExpectedRevision string) ([]byte, error) {
	if outerExpectedRevision == "" {
		workbook, err := xlsxpatch.ExtractNativeWorkbookV2(original)
		if err != nil {
			return nil, err
		}
		outerExpectedRevision = workbook.Source.PackageSHA256
	}
	result, err := xlsxpatch.ApplyNativeWorkbookMutationPayloadV1(original, payload, outerExpectedRevision)
	if err != nil {
		return nil, err
	}
	return result.Package, nil
}

// NewHandler serves native extract, read-only preview objects, and mutations.
// store may be nil; artifact_id is then refused.
func NewHandler(store Store) http.Handler {
	mux := http.NewServeMux()
	Register(mux, store)
	return WithLocalHelperHeaders(mux)
}

// Register mounts the XLSX extract/inspection/mutation routes without CORS.
func Register(mux *http.ServeMux, store Store) {
	mux.HandleFunc(PreviewObjectsPath, func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
			return
		}
		data, _, err := ReadExtractRequest(w, r, store, XLSXEnvelope())
		if err != nil {
			WriteBodyError(w, err)
			return
		}
		projection, err := xlsxpatch.InspectNativeWorkbookObjectsV1(data)
		if err != nil {
			WriteError(w, http.StatusBadRequest, err)
			return
		}
		encoded, err := json.Marshal(projection)
		if err != nil {
			WriteError(w, http.StatusInternalServerError, err)
			return
		}
		w.Header().Set("Content-Type", "application/json")
		w.Header().Set("Cache-Control", "no-store")
		w.WriteHeader(http.StatusOK)
		_, _ = w.Write(encoded)
	})
	mux.HandleFunc(ExtractPath, func(w http.ResponseWriter, r *http.Request) {
		handleExtract(w, r, store)
	})
	mux.HandleFunc(MutationsPath, func(w http.ResponseWriter, r *http.Request) {
		handleMutations(w, r, store)
	})
}

// WithLocalHelperHeaders adds CORS for the local helper and injoffice-server.
func WithLocalHelperHeaders(next http.Handler) http.Handler {
	return withLocalHelperHeaders(next)
}

func withLocalHelperHeaders(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("Access-Control-Allow-Origin", "*")
		w.Header().Set("Access-Control-Allow-Methods", "POST, OPTIONS")
		w.Header().Set("Access-Control-Allow-Headers", "Content-Type")
		w.Header().Set("Access-Control-Expose-Headers", HeaderRevision+", "+HeaderPackageSHA+", "+HeaderArtifactID)
		if r.Method == http.MethodOptions {
			w.WriteHeader(http.StatusNoContent)
			return
		}
		next.ServeHTTP(w, r)
	})
}

func handleExtract(w http.ResponseWriter, r *http.Request, store Store) {
	if r.Method != http.MethodPost {
		WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	data, artifactID, err := ReadExtractRequest(w, r, store, XLSXEnvelope())
	if err != nil {
		WriteBodyError(w, err)
		return
	}
	encoded, err := ExtractNativeJSON(data)
	if err != nil {
		WriteError(w, http.StatusBadRequest, err)
		return
	}
	artifactID, err = PersistExtract(store, data, artifactID)
	if err != nil {
		WriteBodyError(w, err)
		return
	}
	if artifactID != "" {
		w.Header().Set(HeaderArtifactID, artifactID)
	}
	w.Header().Set("Content-Type", xlsxpatch.NativeXLSXV2MediaType)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(encoded)
}

func handleMutations(w http.ResponseWriter, r *http.Request, store Store) {
	if r.Method != http.MethodPost {
		WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	original, payload, outer, artifactID, err := ReadMutationRequest(w, r, store, XLSXEnvelope())
	if err != nil {
		WriteBodyError(w, err)
		return
	}
	result, err := xlsxpatch.ApplyNativeWorkbookMutationPayloadV1(original, payload, outer)
	if err != nil {
		WriteError(w, http.StatusBadRequest, err)
		return
	}
	if err := PersistMutation(store, artifactID, result.Package, original); err != nil {
		WriteBodyError(w, err)
		return
	}
	if artifactID != "" {
		w.Header().Set(HeaderArtifactID, artifactID)
	}
	w.Header().Set("Content-Type", XLSXContentType)
	w.Header().Set("Content-Disposition", `attachment; filename="mutated.xlsx"`)
	w.Header().Set(HeaderRevision, result.Workbook.Revision)
	w.Header().Set(HeaderPackageSHA, result.Workbook.Source.PackageSHA256)
	w.WriteHeader(http.StatusOK)
	_, _ = w.Write(result.Package)
}

// ReadExtractRequest parses a native extract body: JSON artifact_id, multipart
// file, or raw package bytes. env supplies format-specific size and field names.
func ReadExtractRequest(w http.ResponseWriter, r *http.Request, store Store, env Envelope) (data []byte, artifactID string, err error) {
	r.Body = http.MaxBytesReader(w, r.Body, env.MaxPackageBytes)
	media, _, _ := mime.ParseMediaType(r.Header.Get("Content-Type"))
	if media == "application/json" {
		id, err := decodeArtifactIDJSON(r.Body)
		if err != nil {
			return nil, "", err
		}
		data, err := loadArtifact(store, id)
		return data, id, err
	}
	if strings.HasPrefix(media, "multipart/") {
		if err := r.ParseMultipartForm(multipartMemory); err != nil {
			return nil, "", err
		}
		id := strings.TrimSpace(r.FormValue("artifact_id"))
		if id != "" {
			for _, name := range env.FileFields {
				if hasFormFile(r, name) {
					return nil, "", fmt.Errorf("artifact_id cannot be combined with an uploaded %s", env.FormatName)
				}
			}
			data, err := loadArtifact(store, id)
			return data, id, err
		}
		data, err := readMultipartPackage(r, env.FileFields, env.FormatName)
		if err != nil {
			return nil, "", err
		}
		return data, "", nil
	}
	data, err = io.ReadAll(r.Body)
	if err != nil {
		return nil, "", err
	}
	return data, "", nil
}

// ReadMutationRequest parses multipart original/payload/expected_revision, or
// artifact_id in place of original bytes.
func ReadMutationRequest(w http.ResponseWriter, r *http.Request, store Store, env Envelope) (original, payload []byte, outer, artifactID string, err error) {
	r.Body = http.MaxBytesReader(w, r.Body, env.MaxPackageBytes+env.MaxPayloadBytes)
	media, _, _ := mime.ParseMediaType(r.Header.Get("Content-Type"))
	if media != "multipart/form-data" {
		return nil, nil, "", "", fmt.Errorf("multipart/form-data required (original, payload, expected_revision)")
	}
	if err := r.ParseMultipartForm(multipartMemory); err != nil {
		return nil, nil, "", "", err
	}
	artifactID = strings.TrimSpace(r.FormValue("artifact_id"))
	if artifactID != "" {
		if hasFormFile(r, "original") {
			return nil, nil, "", "", fmt.Errorf("artifact_id cannot be combined with original bytes")
		}
		original, err = loadArtifact(store, artifactID)
		if err != nil {
			return nil, nil, "", "", err
		}
	} else {
		original, err = readFormFile(r, "original")
		if err != nil {
			return nil, nil, "", "", fmt.Errorf("original: %w", err)
		}
	}
	payload, err = readFormValueOrFile(r, "payload")
	if err != nil {
		return nil, nil, "", "", fmt.Errorf("payload: %w", err)
	}
	outer = strings.TrimSpace(r.FormValue("expected_revision"))
	if outer == "" {
		if fileBytes, fileErr := readFormFile(r, "expected_revision"); fileErr == nil {
			outer = strings.TrimSpace(string(fileBytes))
		}
	}
	if outer == "" {
		return nil, nil, "", "", fmt.Errorf("expected_revision is required (sha256:<digest>)")
	}
	return original, payload, outer, artifactID, nil
}

// PersistExtract mints an opaque ID after a successful extract of uploaded bytes.
func PersistExtract(store Store, data []byte, artifactID string) (string, error) {
	if artifactID != "" || store == nil {
		return artifactID, nil
	}
	return store.Create(data)
}

// PersistMutation compare-and-swaps stored bytes after a successful apply.
func PersistMutation(store Store, artifactID string, produced, original []byte) error {
	if artifactID == "" {
		return nil
	}
	return store.Put(artifactID, produced, original)
}

func loadArtifact(store Store, id string) ([]byte, error) {
	if store == nil {
		return nil, fmt.Errorf("artifact_id requires a configured artifact store")
	}
	data, err := store.Get(id)
	if err != nil {
		return nil, err
	}
	return data, nil
}

func decodeArtifactIDJSON(body io.Reader) (string, error) {
	var req struct {
		ArtifactID string `json:"artifact_id"`
	}
	decoder := json.NewDecoder(body)
	decoder.DisallowUnknownFields()
	if err := decoder.Decode(&req); err != nil {
		return "", fmt.Errorf("artifact_id: %w", err)
	}
	id := strings.TrimSpace(req.ArtifactID)
	if id == "" {
		return "", fmt.Errorf("artifact_id is required")
	}
	return id, nil
}

func readMultipartPackage(r *http.Request, names []string, format string) ([]byte, error) {
	for _, name := range names {
		data, err := readFormFile(r, name)
		if err == nil {
			return data, nil
		}
	}
	if r.MultipartForm != nil {
		for _, files := range r.MultipartForm.File {
			if len(files) == 0 {
				continue
			}
			file, err := files[0].Open()
			if err != nil {
				return nil, err
			}
			defer file.Close()
			return io.ReadAll(file)
		}
	}
	return nil, fmt.Errorf("multipart body is missing an %s file field", format)
}

func hasFormFile(r *http.Request, name string) bool {
	if r.MultipartForm == nil {
		return false
	}
	return len(r.MultipartForm.File[name]) > 0
}

func readFormFile(r *http.Request, name string) ([]byte, error) {
	file, _, err := r.FormFile(name)
	if err != nil {
		return nil, err
	}
	defer file.Close()
	return io.ReadAll(file)
}

func readFormValueOrFile(r *http.Request, name string) ([]byte, error) {
	if value := r.FormValue(name); value != "" {
		return []byte(value), nil
	}
	return readFormFile(r, name)
}

// WriteBodyError maps artifact-store and body-limit errors to HTTP status.
func WriteBodyError(w http.ResponseWriter, err error) {
	status := http.StatusBadRequest
	var maxBytes *http.MaxBytesError
	switch {
	case errors.As(err, &maxBytes), errors.Is(err, ErrArtifactTooLarge):
		status = http.StatusRequestEntityTooLarge
	case errors.Is(err, ErrArtifactNotFound):
		status = http.StatusNotFound
	case errors.Is(err, ErrInvalidArtifactID), errors.Is(err, ErrArtifactStale):
		status = http.StatusBadRequest
	case errors.Is(err, ErrArtifactIO), errors.Is(err, ErrArtifactExists):
		status = http.StatusInternalServerError
	}
	WriteError(w, status, err)
}

// WriteError writes a JSON {"error":"..."} body.
func WriteError(w http.ResponseWriter, status int, err error) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(map[string]string{"error": err.Error()})
}
