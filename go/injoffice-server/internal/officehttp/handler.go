// Package officehttp is the injoffice-server HTTP surface for native Office
// extract/apply. XLSX routes are the shared xlsxhttp handlers; DOCX and PPTX
// reuse that envelope, CAS, and opaque artifact store.
package officehttp

import (
	"net/http"

	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

// NewHandler serves POST /v1/{xlsx,docx,pptx}/{extract,mutations}.
// store may be nil; artifact_id is then refused.
func NewHandler(store xlsxhttp.Store) http.Handler {
	return NewHandlerWithDOCXPreview(store, DOCXPreviewOptions{})
}

func NewHandlerWithDOCXPreview(store xlsxhttp.Store, preview DOCXPreviewOptions) http.Handler {
	return NewHandlerWithPreviews(store, preview, PPTXPreviewOptions{})
}

func NewHandlerWithPreviews(store xlsxhttp.Store, preview DOCXPreviewOptions, pptxPreview PPTXPreviewOptions) http.Handler {
	mux := http.NewServeMux()
	gate := make(chan struct{}, 1)
	mux.HandleFunc(DOCXPreviewPath, func(w http.ResponseWriter, r *http.Request) { handleDOCXPreview(w, r, preview, gate) })
	mux.HandleFunc(DOCXApproximatePreviewPath, func(w http.ResponseWriter, r *http.Request) { handleDOCXApproximatePreview(w, r, preview, gate) })
	mux.HandleFunc(PPTXPreviewPath, func(w http.ResponseWriter, r *http.Request) { handlePPTXPreview(w, r, pptxPreview, gate) })
	xlsxhttp.Register(mux, store)
	mux.HandleFunc(DOCXExtractPath, func(w http.ResponseWriter, r *http.Request) {
		handleDOCXExtract(w, r, store)
	})
	mux.HandleFunc(DOCXMutationsPath, func(w http.ResponseWriter, r *http.Request) {
		handleDOCXMutations(w, r, store)
	})
	mux.HandleFunc(PPTXExtractPath, func(w http.ResponseWriter, r *http.Request) {
		handlePPTXExtract(w, r, store)
	})
	mux.HandleFunc(PPTXMutationsPath, func(w http.ResponseWriter, r *http.Request) {
		handlePPTXMutations(w, r, store)
	})
	return xlsxhttp.WithLocalHelperHeaders(mux)
}
