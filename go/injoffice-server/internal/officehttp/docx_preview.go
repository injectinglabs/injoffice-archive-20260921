package officehttp

import (
	"archive/zip"
	"bytes"
	"context"
	"crypto/sha256"
	"encoding/base64"
	"encoding/binary"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"os/exec"
	"time"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const DOCXPreviewPath = "/v1/docx/page-preview"

// DOCXPreviewOptions explicitly enables a local Node compiler. The worker path
// is operator configuration, never a URL or a caller-supplied executable.
type DOCXPreviewOptions struct{ WorkerPath string }

func docxPreviewInput(ctx context.Context, data []byte) (map[string]any, error) {
	if len(data) > 8*1024*1024 {
		return nil, errors.New("native preview packages are limited to 8 MiB")
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	doc, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	layout, err := docxpatch.ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	settings, err := docxpatch.ExtractNativePaginationSettingsV1(data)
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	inventory, err := docxpatch.ExtractNativeDOCXFontInventoryV1(data)
	if err != nil {
		return nil, err
	}
	if err := ctx.Err(); err != nil {
		return nil, err
	}
	fonts, err := docxpatch.ResolveNativeDOCXPagePaintFontAssetsV1(data, inventory)
	if err != nil {
		return nil, err
	}
	inventoryJSON, err := docxpatch.EncodeNativeDOCXFontInventoryV1(inventory)
	if err != nil {
		return nil, err
	}
	if fonts == nil {
		fonts = []docxpatch.NativeDOCXPagePaintFontAssetV1{}
	}
	// Only package-referenced PNG/JPEG bytes enter the compiler; its authoritative
	// media join independently checks part names, content digests and geometry.
	encoded, err := json.Marshal(doc)
	if err != nil {
		return nil, err
	}
	var value any
	if err = json.Unmarshal(encoded, &value); err != nil {
		return nil, err
	}
	names := map[string]string{}
	var walk func(any)
	walk = func(value any) {
		switch v := value.(type) {
		case map[string]any:
			if name, ok := v["media_part"].(string); ok && (v["content_type"] == "image/png" || v["content_type"] == "image/jpeg") {
				names[name] = v["content_type"].(string)
			}
			for _, child := range v {
				walk(child)
			}
		case []any:
			for _, child := range v {
				walk(child)
			}
		}
	}
	walk(value)
	media := []map[string]any{}
	archive, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		return nil, err
	}
	total := 0
	for _, part := range archive.File {
		if err := ctx.Err(); err != nil {
			return nil, err
		}
		if names[part.Name] == "" {
			continue
		}
		reader, err := part.Open()
		if err != nil {
			return nil, err
		}
		content, readErr := io.ReadAll(io.LimitReader(reader, 8*1024*1024+1))
		reader.Close()
		if readErr != nil {
			return nil, readErr
		}
		total += len(content)
		if len(content) > 8*1024*1024 || total > 32*1024*1024 || len(media) >= 256 {
			return nil, errors.New("preview media exceeds its byte budget")
		}
		media = append(media, map[string]any{"part_name": part.Name, "content_type": names[part.Name], "content_digest": fmt.Sprintf("sha256:%x", sha256.Sum256(content)), "bytes_base64": base64.StdEncoding.EncodeToString(content)})
	}
	return map[string]any{"protocol": "injoffice.docx.page-paint-compiler", "version": 1, "source_revision": "injoffice-docx-preview-v1", "outline_provider": map[string]string{"provider_id": "injoffice.harfbuzz-outline", "provider_revision": "v1"}, "document": doc, "resolved_layout": layout, "pagination_settings": settings, "font_inventory_json": string(inventoryJSON), "font_assets": fonts, "media_assets": media}, nil
}

func compileDOCXPreview(ctx context.Context, options DOCXPreviewOptions, input map[string]any) (json.RawMessage, error) {
	payload, err := json.Marshal(map[string]any{"protocol": "injoffice.docx.page-paint-worker", "version": 1, "id": "preview", "op": "render", "input": input})
	if err != nil {
		return nil, err
	}
	if len(payload) > 192*1024*1024 {
		return nil, errors.New("preview request exceeds its frame budget")
	}
	frame := make([]byte, 4+len(payload))
	binary.BigEndian.PutUint32(frame, uint32(len(payload)))
	copy(frame[4:], payload)
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	command := exec.CommandContext(ctx, "node", "--max-old-space-size=512", options.WorkerPath)
	command.Stdin = bytes.NewReader(frame)
	stdout, err := command.StdoutPipe()
	if err != nil {
		return nil, err
	}
	if err = command.Start(); err != nil {
		return nil, errors.New("native preview worker could not start")
	}
	output, readErr := io.ReadAll(io.LimitReader(stdout, 64*1024*1024+5))
	if readErr != nil || len(output) > 64*1024*1024+4 {
		_ = command.Process.Kill()
	}
	waitErr := command.Wait()
	if ctx.Err() != nil {
		return nil, errors.New("native preview exceeded its time budget")
	}
	if readErr != nil || waitErr != nil || len(output) < 4 || int(binary.BigEndian.Uint32(output[:4])) != len(output)-4 {
		return nil, errors.New("native preview worker returned an invalid frame")
	}
	var response struct {
		Protocol string          `json:"protocol"`
		Version  int             `json:"version"`
		ID       string          `json:"id"`
		OK       bool            `json:"ok"`
		Result   json.RawMessage `json:"result"`
		Error    struct {
			Message string `json:"message"`
		} `json:"error"`
	}
	if err = json.Unmarshal(output[4:], &response); err != nil || response.Protocol != "injoffice.docx.page-paint-worker" || response.Version != 1 || response.ID != "preview" {
		return nil, errors.New("native preview worker returned an invalid envelope")
	}
	if !response.OK {
		return nil, errors.New(response.Error.Message)
	}
	return response.Result, nil
}

func handleDOCXPreview(w http.ResponseWriter, r *http.Request, options DOCXPreviewOptions, gate chan struct{}) {
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	if options.WorkerPath == "" {
		xlsxhttp.WriteError(w, http.StatusServiceUnavailable, errors.New("native page preview is not enabled by the server operator"))
		return
	}
	select {
	case gate <- struct{}{}:
		defer func() { <-gate }()
	default:
		xlsxhttp.WriteError(w, http.StatusServiceUnavailable, errors.New("native preview is busy; retry after the current document completes"))
		return
	}
	envelope := docxEnvelope()
	envelope.MaxPackageBytes = 8 * 1024 * 1024
	data, _, err := xlsxhttp.ReadExtractRequest(w, r, nil, envelope)
	if err != nil {
		xlsxhttp.WriteBodyError(w, err)
		return
	}
	ctx, cancel := context.WithTimeout(r.Context(), 45*time.Second)
	defer cancel()
	input, err := docxPreviewInput(ctx, data)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, err)
		return
	}
	result, err := compileDOCXPreview(ctx, options, input)
	if err != nil {
		xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, err)
		return
	}
	w.Header().Set("Content-Type", "application/json")
	w.Header().Set("Cache-Control", "no-store")
	_, _ = w.Write(result)
}
