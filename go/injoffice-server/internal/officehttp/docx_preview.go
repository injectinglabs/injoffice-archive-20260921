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
const DOCXApproximatePreviewPath = "/v1/docx/page-preview-approximate"
const DOCXTextboxPreviewPath = "/v1/docx/page-preview-textboxes"
const DOCXFontSubstitutionPreviewPath = "/v1/docx/page-preview-font-substitution"

// DOCXPreviewOptions explicitly enables a local Node compiler. The worker path
// is operator configuration, never a URL or a caller-supplied executable.
type DOCXPreviewOptions struct{ WorkerPath, FontManifestPath string }

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
	args := []string{}
	if options.FontManifestPath != "" {
		args = append(args, "--font-manifest", options.FontManifestPath)
	}
	return compilePreviewWorker(ctx, options.WorkerPath, "injoffice.docx.page-paint-worker", input, 192*1024*1024, 64*1024*1024, args...)
}

func compilePreviewWorker(ctx context.Context, workerPath, protocol string, input map[string]any, maxInput, maxOutput int, workerArgs ...string) (json.RawMessage, error) {
	return compilePreviewWorkerOperation(ctx, workerPath, protocol, "render", input, maxInput, maxOutput, workerArgs...)
}

// Operations are selected by server routes, never by request bodies.
func compilePreviewWorkerOperation(ctx context.Context, workerPath, protocol, operation string, input map[string]any, maxInput, maxOutput int, workerArgs ...string) (json.RawMessage, error) {
	payload, err := json.Marshal(map[string]any{"protocol": protocol, "version": 1, "id": "preview", "op": operation, "input": input})
	if err != nil {
		return nil, err
	}
	if len(payload) > maxInput {
		return nil, errors.New("preview request exceeds its frame budget")
	}
	frame := make([]byte, 4+len(payload))
	binary.BigEndian.PutUint32(frame, uint32(len(payload)))
	copy(frame[4:], payload)
	ctx, cancel := context.WithTimeout(ctx, 30*time.Second)
	defer cancel()
	command := exec.CommandContext(ctx, "node", append([]string{"--max-old-space-size=512", workerPath}, workerArgs...)...)
	command.Stdin = bytes.NewReader(frame)
	stdout, err := command.StdoutPipe()
	if err != nil {
		return nil, err
	}
	if err = command.Start(); err != nil {
		return nil, errors.New("native preview worker could not start")
	}
	output, readErr := io.ReadAll(io.LimitReader(stdout, int64(maxOutput+5)))
	if readErr != nil || len(output) > maxOutput+4 {
		_ = command.Process.Kill()
	}
	waitErr := command.Wait()
	if ctx.Err() != nil {
		return nil, errors.New("native preview exceeded its time budget")
	}
	if readErr != nil || waitErr != nil || len(output) < 4 || len(output) > maxOutput+4 || int(binary.BigEndian.Uint32(output[:4])) != len(output)-4 {
		return nil, errors.New("native preview worker returned an invalid frame")
	}
	var response struct {
		Protocol string          `json:"protocol"`
		Version  int             `json:"version"`
		ID       string          `json:"id"`
		OK       bool            `json:"ok"`
		Result   json.RawMessage `json:"result"`
		Error    struct {
			Code    string `json:"code"`
			ScopeID string `json:"scope_id"`
			Message string `json:"message"`
		} `json:"error"`
	}
	if err = json.Unmarshal(output[4:], &response); err != nil || response.Protocol != protocol || response.Version != 1 || response.ID != "preview" {
		return nil, errors.New("native preview worker returned an invalid envelope")
	}
	if !response.OK {
		if code := response.Error.Code; code != "" && code != "COMPILATION_REFUSED" {
			return nil, &docxPreviewRefusal{Code: code, ScopeID: response.Error.ScopeID, Message: response.Error.Message}
		}
		return nil, errors.New(response.Error.Message)
	}
	return response.Result, nil
}

// docxPreviewRefusal is a preview refusal the compiler named. Its code and
// scope come from the worker's typed refusal record, so the 422 body can state
// which source fact the preview does not implement instead of carrying only an
// English sentence a caller cannot branch on.
type docxPreviewRefusal struct {
	Code    string
	ScopeID string
	Message string
}

func (refusal *docxPreviewRefusal) Error() string { return refusal.Message }

// writeDOCXPreviewError keeps the {"error": "..."} body every client already
// reads and adds the typed record when the compiler produced one.
func writeDOCXPreviewError(w http.ResponseWriter, err error) {
	var refusal *docxPreviewRefusal
	if !errors.As(err, &refusal) {
		xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, err)
		return
	}
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(http.StatusUnprocessableEntity)
	_ = json.NewEncoder(w).Encode(map[string]any{
		"error": refusal.Message,
		"refusal": map[string]any{
			"protocol": "injoffice.docx.preview-refusal",
			"version":  1,
			"code":     refusal.Code,
			"scope_id": refusal.ScopeID,
			"message":  refusal.Message,
		},
	})
}

func handleDOCXPreview(w http.ResponseWriter, r *http.Request, options DOCXPreviewOptions, gate chan struct{}) {
	handleDOCXPreviewMode(w, r, options, gate, false)
}

func handleDOCXApproximatePreview(w http.ResponseWriter, r *http.Request, options DOCXPreviewOptions, gate chan struct{}) {
	handleDOCXPreviewMode(w, r, options, gate, true)
}

func handleDOCXPreviewMode(w http.ResponseWriter, r *http.Request, options DOCXPreviewOptions, gate chan struct{}, approximate bool) {
	fontSubstitution := r.URL.Path == DOCXFontSubstitutionPreviewPath
	textbox := r.URL.Path == DOCXTextboxPreviewPath
	if r.Method != http.MethodPost {
		xlsxhttp.WriteError(w, http.StatusMethodNotAllowed, errors.New("POST required"))
		return
	}
	if options.WorkerPath == "" {
		xlsxhttp.WriteError(w, http.StatusServiceUnavailable, errors.New("native page preview is not enabled by the server operator"))
		return
	}
	if textbox && r.URL.RawQuery != "" {
		xlsxhttp.WriteError(w, http.StatusBadRequest, errors.New("textbox preview accepts no query options"))
		return
	}
	if fontSubstitution && (r.URL.RawQuery != "" || options.FontManifestPath == "") {
		xlsxhttp.WriteError(w, http.StatusBadRequest, errors.New("font substitution preview requires operator font configuration and accepts no query options"))
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
	var result json.RawMessage
	if textbox {
		var inspected []byte
		inspected, err = docxpatch.InspectNativePartialSourceV1(data)
		if err == nil {
			var source struct {
				Geometry json.RawMessage `json:"textbox_geometry"`
			}
			err = json.Unmarshal(inspected, &source)
			if err == nil {
				args := []string{}
				if options.FontManifestPath != "" {
					args = append(args, "--font-manifest", options.FontManifestPath)
				}
				// A document with no text box has no geometry part at all. Like every
				// other read-only sidecar here, omit the key rather than sending a JSON
				// null the decoder would read as a malformed evidence record.
				workerInput := map[string]any{"prepare": input}
				if nativeTextboxGeometryPresent(source.Geometry) {
					workerInput["evidence"] = source.Geometry
				}
				result, err = compilePreviewWorkerOperation(ctx, options.WorkerPath, "injoffice.docx.page-paint-worker", "render-textbox-pages", workerInput, 192*1024*1024, 64*1024*1024, args...)
			}
		}
	} else if fontSubstitution {
		var composition map[string]any
		composition, err = docxFontPreviewComposition(input, data)
		if err == nil {
			result, err = compilePreviewWorkerOperation(ctx, options.WorkerPath, "injoffice.docx.page-paint-worker", "render-font-substitution", map[string]any{"prepare": input, "composition": composition}, 192*1024*1024, 64*1024*1024, "--font-manifest", options.FontManifestPath)
		}
		if err == nil {
			err = validateDOCXFontSubstitutionPreview(result, input, options.FontManifestPath, composition)
		}
	} else if approximate {
		eligibility, eligibilityErr := docxpatch.ExtractNativeDocxApproximationEligibilityV1(data)
		if eligibilityErr != nil {
			xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, eligibilityErr)
			return
		}
		args := []string{}
		if options.FontManifestPath != "" {
			args = append(args, "--font-manifest", options.FontManifestPath)
		}
		operation, workerInput := docxApproximateWorkerInput(input, eligibility)
		if operation == "render-approximate" {
			// Same-bytes read-only sidecar; the compiler re-validates its joins.
			shapes, shapesErr := docxpatch.InspectNativeApproximateDrawingShapesV1(data)
			if shapesErr != nil {
				xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, shapesErr)
				return
			}
			if shapes != nil {
				workerInput["drawing_shapes"] = shapes
			}
			// Same-bytes read-only OMML sidecar; the compiler re-validates its joins.
			equations, equationsErr := docxpatch.InspectNativeApproximateEquationsV1(data)
			if equationsErr != nil {
				xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, equationsErr)
				return
			}
			if equations != nil {
				workerInput["equations"] = equations
			}
			charts, chartsErr := docxpatch.InspectNativeApproximateDrawingChartsV1(data)
			if chartsErr != nil {
				xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, chartsErr)
				return
			}
			if charts != nil {
				workerInput["drawing_charts"] = charts
			}
			// Same-bytes read-only nested-table sidecar; the compiler re-validates its joins.
			nested, nestedErr := docxpatch.InspectNativeApproximateNestedTablesV1(data)
			if nestedErr != nil {
				xlsxhttp.WriteError(w, http.StatusUnprocessableEntity, nestedErr)
				return
			}
			if nested != nil {
				workerInput["nested_tables"] = nested
			}
		}
		result, err = compilePreviewWorkerOperation(ctx, options.WorkerPath, "injoffice.docx.page-paint-worker", operation, workerInput, 192*1024*1024, 64*1024*1024, args...)
	} else {
		result, err = compileDOCXPreview(ctx, options, input)
	}
	if err != nil {
		writeDOCXPreviewError(w, err)
		return
	}
	w.Header().Set("Content-Type", "application/json")
	w.Header().Set("Cache-Control", "no-store")
	_, _ = w.Write(result)
}

// nativeTextboxGeometryPresent reports whether the source inspector actually
// produced a text box geometry record. The part is absent from documents that
// contain no text box, and an absent json.RawMessage marshals as "null".
func nativeTextboxGeometryPresent(geometry json.RawMessage) bool {
	trimmed := bytes.TrimSpace(geometry)
	return len(trimmed) > 0 && !bytes.Equal(trimmed, []byte("null"))
}

func docxFontPreviewComposition(input map[string]any, data []byte) (map[string]any, error) {
	descriptors, err := docxpatch.ExtractNativeDOCXFontSubstitutionEligibilityV1(data)
	if err != nil {
		return nil, err
	}
	composition := map[string]any{"source_document": input["document"], "source_resolved_layout": input["resolved_layout"], "source_pagination_settings": input["pagination_settings"], "source_font_inventory_json": input["font_inventory_json"], "font_descriptor_eligibility": descriptors}
	settings, _ := input["pagination_settings"].(*docxpatch.NativePaginationSettingsV1)
	if settings == nil {
		return nil, errors.New("font composition settings missing")
	}
	if settings.Profile != "word-modern-default" {
		legacy, err := docxpatch.ExtractNativeDocxApproximationEligibilityV1(data)
		if err != nil {
			return nil, err
		}
		if legacy.Status != "eligible" {
			return nil, errors.New("font composition legacy settings ineligible")
		}
		composition["legacy_eligibility"] = legacy
		if halfPoints, ok := docxpatch.NativeDocxHostDefaultSizeHalfPointsV1(legacy.AbsentFontSizeShape); ok && len(legacy.AbsentFontSizes) > 0 {
			composition["font_size_policy"] = map[string]any{"kind": "host-default-size-v1", "half_points": halfPoints}
		}
		if len(legacy.AbsentFontFamilies) > 0 {
			composition["font_family_policy"] = map[string]any{"kind": "host-default-family-v1", "family": "Aptos"}
		}
	}
	layout, _ := input["resolved_layout"].(*docxpatch.NativeResolvedLayoutInputV1)
	if layout != nil {
		for _, table := range layout.Tables {
			if table.AutomaticBorderPreview != nil {
				composition["automatic_borders"] = true
			}
		}
	}
	return composition, nil
}

func docxApproximateWorkerInput(input map[string]any, eligibility *docxpatch.NativeDocxApproximationEligibilityV1) (string, map[string]any) {
	addFontPolicy := func(request map[string]any) map[string]any {
		if eligibility != nil && eligibility.Status == "eligible" && len(eligibility.AbsentFontSizes) > 0 {
			// A declared preview-host choice. Its value is not invented: it is
			// the size Microsoft Word 16.112 itself writes into the Tf operator
			// for this source shape, so the two shapes select different sizes.
			if halfPoints, ok := docxpatch.NativeDocxHostDefaultSizeHalfPointsV1(eligibility.AbsentFontSizeShape); ok {
				request["font_size_policy"] = map[string]any{"kind": "host-default-size-v1", "half_points": halfPoints}
			}
		}
		if eligibility != nil && eligibility.Status == "eligible" && len(eligibility.AbsentFontFamilies) > 0 {
			// Likewise declared, never an authored face: the value was measured
			// against Microsoft Word 16.112.4 references for packages that
			// select no font anywhere.
			request["font_family_policy"] = map[string]any{"kind": "host-default-family-v1", "family": "Aptos"}
		}
		return request
	}
	layout, _ := input["resolved_layout"].(*docxpatch.NativeResolvedLayoutInputV1)
	if layout != nil {
		for _, table := range layout.Tables {
			if table.AutomaticBorderPreview == nil {
				continue
			}
			request := map[string]any{"prepare": input}
			settings, _ := input["pagination_settings"].(*docxpatch.NativePaginationSettingsV1)
			if settings == nil || settings.Profile != "word-modern-default" {
				request["legacy_eligibility"] = eligibility
			}
			return "render-auto-borders", addFontPolicy(request)
		}
	}
	return "render-approximate", addFontPolicy(map[string]any{"prepare": input, "eligibility": eligibility})
}
