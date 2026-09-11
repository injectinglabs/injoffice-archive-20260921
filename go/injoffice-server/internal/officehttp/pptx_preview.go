package officehttp

import (
	"context"
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"net/http"
	"net/url"
	"path/filepath"
	"strconv"
	"time"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const PPTXPreviewPath = "/v1/pptx/slide-preview"

// Paths are operator configuration, never request parameters. The manifest
// names locally licensed fonts by exact family/style and content digest.
type PPTXPreviewOptions struct{ WorkerPath, FontManifestPath string }

func pptxPreviewInput(ctx context.Context, data []byte, slide int, options PPTXPreviewOptions) (map[string]any, error) {
	if len(data) > 8*1024*1024 {
		return nil, errors.New("native preview packages are limited to 8 MiB")
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if slide < 0 {
		return nil, errors.New("slide must be a zero-based nonnegative index")
	}
	if err := preflightPPTXPreviewZIP(ctx, data); err != nil {
		return nil, err
	}
	deck, err := pptxpatch.ExtractNativePPTX(data, pptxExtractOptions)
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	if slide >= len(deck.Slides) {
		return nil, errors.New("slide index is outside the source presentation")
	}
	if err := attachPPTXPreviewImages(ctx, data, &deck, slide); err != nil {
		return nil, err
	}
	return map[string]any{"deck": deck, "package_sha256": fmt.Sprintf("%x", sha256.Sum256(data)), "slide_index": slide, "font_manifest_path": options.FontManifestPath}, nil
}

func handlePPTXPreview(w http.ResponseWriter, r *http.Request, options PPTXPreviewOptions, gate chan struct{}) {
	w.Header().Set("Cache-Control", "no-store")
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	if !filepath.IsAbs(options.WorkerPath) || !filepath.IsAbs(options.FontManifestPath) {
		xlsxhttp.WriteError(w, http.StatusServiceUnavailable, errors.New("native slide preview and exact font manifest are not configured by the server operator"))
		return
	}
	query, err := parsePPTXPreviewSlide(r)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusBadRequest, err)
		return
	}
	select {
	case gate <- struct{}{}:
		defer func() { <-gate }()
	default:
		xlsxhttp.WriteError(w, http.StatusServiceUnavailable, errors.New("native preview is busy; retry after the current document completes"))
		return
	}
	envelope := pptxEnvelope()
	envelope.MaxPackageBytes = 8 * 1024 * 1024
	data, _, err := xlsxhttp.ReadExtractRequest(w, r, nil, envelope)
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	ctx, cancel := context.WithTimeout(r.Context(), 45*time.Second)
	defer cancel()
	input, err := pptxPreviewInput(ctx, data, query, options)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, err)
		return
	}
	result, err := compilePreviewWorker(ctx, options.WorkerPath, "injoffice.pptx.preview-worker", input, 16*1024*1024, 16*1024*1024)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, err)
		return
	}
	var identity struct {
		Version       int    `json:"version"`
		PackageSHA256 string `json:"package_sha256"`
		SlideIndex    *int   `json:"slide_index"`
		SlideCount    int    `json:"slide_count"`
	}
	if json.Unmarshal(result, &identity) != nil || identity.Version != 1 || identity.PackageSHA256 != input["package_sha256"] || identity.SlideIndex == nil || *identity.SlideIndex != query || identity.SlideCount != len(input["deck"].(pptxpatch.NativePPTXDeck).Slides) {
		xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, errors.New("native preview worker result does not match the source slide"))
		return
	}
	w.Header().Set("Content-Type", "application/json")
	_, _ = w.Write(result)
}

func parsePPTXPreviewSlide(r *http.Request) (int, error) {
	// ParseQuery errors must not silently discard malformed or duplicate input.
	values, err := url.ParseQuery(r.URL.RawQuery)
	if err != nil {
		return 0, err
	}
	for key, entries := range values {
		if key != "slide" || len(entries) != 1 {
			return 0, errors.New("only one slide query parameter is supported")
		}
	}
	raw := values.Get("slide")
	if raw == "" && !values.Has("slide") {
		return 0, nil
	}
	index, err := strconv.Atoi(raw)
	if err != nil || index < 0 || strconv.Itoa(index) != raw {
		return 0, errors.New("slide must be a canonical zero-based integer")
	}
	return index, nil
}
