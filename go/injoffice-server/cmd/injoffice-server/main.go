// Command injoffice-server is an optional local HTTP server for native Office
// extract, mutations, and collab HTTP+SSE. It has no agents, SSO, billing,
// tenant RBAC, or Injecting auth.
package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/injectinglabs/injoffice/go/collab"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/artifacthttp"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/collabhttp"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/fsstore"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/officehttp"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const (
	defaultAddr      = "127.0.0.1:18765"
	defaultArtifacts = "artifacts"
	healthPath       = "/healthz"
	capabilitiesPath = "/v1/capabilities"
)

type capabilityRoute struct {
	Method     string   `json:"method"`
	Path       string   `json:"path"`
	Feature    string   `json:"feature"`
	Operations []string `json:"operations,omitempty"`
}

type capabilitiesResponse struct {
	Protocol       string            `json:"protocol"`
	Version        int               `json:"version"`
	Authentication string            `json:"authentication"`
	ArtifactStore  string            `json:"artifact_store"`
	Routes         []capabilityRoute `json:"routes"`
	Limitations    []string          `json:"limitations"`
}

var serverCapabilities = capabilitiesResponse{
	Protocol:       "injoffice.server.capabilities",
	Version:        1,
	Authentication: "none",
	ArtifactStore:  "filesystem",
	Routes: []capabilityRoute{
		{Method: http.MethodGet, Path: healthPath, Feature: "health"},
		{Method: http.MethodGet, Path: capabilitiesPath, Feature: "capability-discovery"},
		{Method: http.MethodPost, Path: xlsxhttp.ExtractPath, Feature: "native-xlsx-extract-v2"},
		{Method: http.MethodPost, Path: xlsxhttp.MutationsPath, Feature: "native-xlsx-mutations", Operations: []string{"cell.set_value", "cell.clear_value", "cell.set_formula", "cell.clear_formula", "style.patch", "row.set_height", "column.set_width"}},
		{Method: http.MethodPost, Path: officehttp.DOCXExtractPath, Feature: "native-docx-extract-v1"},
		{Method: http.MethodPost, Path: officehttp.DOCXMutationsPath, Feature: "native-docx-mutations", Operations: []string{"text.replace"}},
		{Method: http.MethodPost, Path: officehttp.PPTXExtractPath, Feature: "native-pptx-extract-v1"},
		{Method: http.MethodPost, Path: officehttp.PPTXMutationsPath, Feature: "native-pptx-mutations", Operations: []string{"text.replace", "autoshape.update"}},
		{Method: http.MethodPost, Path: artifacthttp.Path, Feature: "artifact-mint"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/session", Feature: "collaboration-session"},
		{Method: http.MethodGet, Path: collabhttp.Prefix + "/session", Feature: "collaboration-session"},
		{Method: http.MethodDelete, Path: collabhttp.Prefix + "/session", Feature: "collaboration-session"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/session/close", Feature: "collaboration-session"},
		{Method: http.MethodGet, Path: collabhttp.Prefix + "/events", Feature: "collaboration-sse"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/join", Feature: "collaboration-membership"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/leave", Feature: "collaboration-membership"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/presence", Feature: "collaboration-presence"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/op/submit", Feature: "collaboration-operation-log"},
		{Method: http.MethodPost, Path: collabhttp.Prefix + "/op/since", Feature: "collaboration-operation-log"},
	},
	Limitations: []string{
		"authentication and artifact authorization are not provided",
		"collaboration operation logs are in-memory and are not written back to Office artifacts",
	},
}

func main() {
	os.Exit(run(os.Args[1:]))
}

func run(args []string) int {
	fs := flag.NewFlagSet("injoffice-server", flag.ContinueOnError)
	fs.SetOutput(os.Stderr)
	addr := fs.String("addr", envOr("INJOFFICE_ADDR", defaultAddr), "listen address (host:port)")
	dir := fs.String("artifacts", envOr("INJOFFICE_ARTIFACTS", defaultArtifacts), "directory for opaque artifact objects")
	previewWorker := fs.String("docx-preview-worker", "", "opt-in absolute path to the compiled local DOCX page-paint worker")
	docxFonts := fs.String("docx-font-manifest", "", "optional absolute operator-owned exact-font manifest for DOCX preview")
	pptxWorker := fs.String("pptx-preview-worker", "", "opt-in absolute path to the local PPTX preview worker")
	pptxFonts := fs.String("pptx-font-manifest", "", "absolute operator-owned exact-font manifest for PPTX preview")
	if err := fs.Parse(args); err != nil {
		return 2
	}
	if *previewWorker != "" && !filepath.IsAbs(*previewWorker) {
		fmt.Fprintln(os.Stderr, "injoffice-server: docx-preview-worker must be an absolute local path")
		return 2
	}
	if *docxFonts != "" && (*previewWorker == "" || !filepath.IsAbs(*docxFonts)) {
		fmt.Fprintln(os.Stderr, "injoffice-server: docx-font-manifest requires an enabled worker and an absolute local path")
		return 2
	}
	if (*pptxWorker == "") != (*pptxFonts == "") || *pptxWorker != "" && (!filepath.IsAbs(*pptxWorker) || !filepath.IsAbs(*pptxFonts)) {
		fmt.Fprintln(os.Stderr, "injoffice-server: pptx-preview-worker and pptx-font-manifest must both be absolute local paths")
		return 2
	}
	store, err := fsstore.Open(*dir)
	if err != nil {
		fmt.Fprintln(os.Stderr, "injoffice-server:", err)
		return 1
	}
	fmt.Fprintf(os.Stderr, "injoffice-server: listening on http://%s artifacts=%s (no auth)\n", *addr, *dir)
	server := &http.Server{
		Addr:              *addr,
		Handler:           newHandlerWithPreviews(store, officehttp.DOCXPreviewOptions{WorkerPath: *previewWorker, FontManifestPath: *docxFonts}, officehttp.PPTXPreviewOptions{WorkerPath: *pptxWorker, FontManifestPath: *pptxFonts}),
		ReadHeaderTimeout: 5 * time.Second,
		ReadTimeout:       60 * time.Second,
		// WriteTimeout stays 0 so the collab SSE stream can idle with keepalives.
		MaxHeaderBytes: 1 << 16,
	}
	if err := server.ListenAndServe(); err != nil {
		fmt.Fprintln(os.Stderr, "injoffice-server:", err)
		return 1
	}
	return 0
}

func newHandler(store xlsxhttp.Store) http.Handler {
	return newHandlerWithPreview(store, officehttp.DOCXPreviewOptions{})
}

func newHandlerWithPreview(store xlsxhttp.Store, preview officehttp.DOCXPreviewOptions) http.Handler {
	return newHandlerWithPreviews(store, preview, officehttp.PPTXPreviewOptions{})
}

func newHandlerWithPreviews(store xlsxhttp.Store, preview officehttp.DOCXPreviewOptions, pptxPreview officehttp.PPTXPreviewOptions) http.Handler {
	office := officehttp.NewHandlerWithPreviews(store, preview, pptxPreview)
	rooms := collabhttp.New(collab.NewHub(), store)
	artifacts := artifacthttp.New(store)
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		if r.URL.Path == healthPath {
			serveHealth(w, r)
			return
		}
		if r.URL.Path == capabilitiesPath {
			serveCapabilities(w, r)
			return
		}
		if r.URL.Path == artifacthttp.Path {
			artifacts.ServeHTTP(w, r)
			return
		}
		if strings.HasPrefix(r.URL.Path, collabhttp.Prefix) {
			rooms.ServeHTTP(w, r)
			return
		}
		office.ServeHTTP(w, r)
	})
}

func serveHealth(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		writeServerError(w, http.StatusMethodNotAllowed, "GET required")
		return
	}
	w.Header().Set("Cache-Control", "no-store")
	writeServerJSON(w, http.StatusOK, map[string]string{"status": "ok"})
}

func serveCapabilities(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodGet {
		writeServerError(w, http.StatusMethodNotAllowed, "GET required")
		return
	}
	writeServerJSON(w, http.StatusOK, serverCapabilities)
}

func writeServerError(w http.ResponseWriter, status int, message string) {
	writeServerJSON(w, status, map[string]string{"error": message})
}

func writeServerJSON(w http.ResponseWriter, status int, body any) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(body)
}

func envOr(key, fallback string) string {
	if value := strings.TrimSpace(os.Getenv(key)); value != "" {
		return value
	}
	return fallback
}
