package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/collab"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/artifacthttp"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/fsstore"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const happyTreeSHA = "c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513"

func TestPPTXPreviewFlagsRequireAbsolutePair(t *testing.T) {
	for _, args := range [][]string{
		{"-pptx-preview-worker", "/operator/worker.js"},
		{"-pptx-font-manifest", "/operator/fonts.json"},
		{"-pptx-preview-worker", "worker.js", "-pptx-font-manifest", "/operator/fonts.json"},
		{"-pptx-preview-worker", "/operator/worker.js", "-pptx-font-manifest", "fonts.json"},
	} {
		if got := run(args); got != 2 {
			t.Fatalf("invalid configuration returned %d", got)
		}
	}
}

func TestDOCXFontFlagsRequireEnabledWorkerAndAbsolutePath(t *testing.T) {
	for _, args := range [][]string{
		{"-docx-font-manifest", "/operator/fonts.json"},
		{"-docx-preview-worker", "/operator/worker.js", "-docx-font-manifest", "fonts.json"},
	} {
		if got := run(args); got != 2 {
			t.Fatalf("invalid configuration returned %d", got)
		}
	}
}

func readHappyTree(t *testing.T) []byte {
	t.Helper()
	path := filepath.Join("..", "..", "..", "xlsxpatch", "testdata", "excel-authored", "happy-tree.xlsx")
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(data)
	if hex.EncodeToString(sum[:]) != happyTreeSHA {
		t.Fatalf("happy-tree.xlsx sha256=%s, want %s", hex.EncodeToString(sum[:]), happyTreeSHA)
	}
	return data
}

func TestHealthz(t *testing.T) {
	request := httptest.NewRequest(http.MethodGet, healthPath, nil)
	response := httptest.NewRecorder()
	newHandler(nil).ServeHTTP(response, request)
	res := response.Result()
	defer res.Body.Close()
	if res.StatusCode != http.StatusOK {
		t.Fatalf("health HTTP %d, want 200", res.StatusCode)
	}
	if got := res.Header.Get("Content-Type"); got != "application/json" {
		t.Fatalf("health content-type %q", got)
	}
	if got := res.Header.Get("Cache-Control"); got != "no-store" {
		t.Fatalf("health cache-control %q", got)
	}
	var body map[string]string
	if err := json.NewDecoder(res.Body).Decode(&body); err != nil {
		t.Fatal(err)
	}
	if body["status"] != "ok" {
		t.Fatalf("health status %q", body["status"])
	}
}

func TestCapabilitiesDescribeImplementedRoutes(t *testing.T) {
	request := httptest.NewRequest(http.MethodGet, capabilitiesPath, nil)
	response := httptest.NewRecorder()
	newHandler(nil).ServeHTTP(response, request)
	res := response.Result()
	defer res.Body.Close()
	if res.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(res.Body)
		t.Fatalf("capabilities HTTP %d: %s", res.StatusCode, body)
	}
	var got capabilitiesResponse
	if err := json.NewDecoder(res.Body).Decode(&got); err != nil {
		t.Fatal(err)
	}
	if got.Protocol != "injoffice.server.capabilities" || got.Version != 1 {
		t.Fatalf("unexpected capability contract: %+v", got)
	}
	if got.Authentication != "none" || got.ArtifactStore != "filesystem" {
		t.Fatalf("capabilities hide runtime boundaries: %+v", got)
	}
	want := map[string]bool{
		http.MethodGet + " " + healthPath:              true,
		http.MethodGet + " " + capabilitiesPath:        true,
		http.MethodPost + " " + xlsxhttp.ExtractPath:   true,
		http.MethodPost + " " + xlsxhttp.MutationsPath: true,
		http.MethodPost + " /v1/docx/extract":          true,
		http.MethodPost + " /v1/docx/mutations":        true,
		http.MethodPost + " /v1/pptx/extract":          true,
		http.MethodPost + " /v1/pptx/mutations":        true,
		http.MethodPost + " /v1/artifacts":             true,
		http.MethodGet + " /v1/collab/events":          true,
		http.MethodPost + " /v1/collab/op/submit":      true,
		http.MethodPost + " /v1/collab/op/since":       true,
	}
	operations := map[string][]string{}
	for _, route := range got.Routes {
		delete(want, route.Method+" "+route.Path)
		operations[route.Path] = route.Operations
	}
	if len(want) != 0 {
		t.Fatalf("capabilities omit implemented routes: %v", want)
	}
	if !contains(operations[xlsxhttp.MutationsPath], "style.patch") || !contains(operations["/v1/docx/mutations"], "text.replace") || !contains(operations["/v1/pptx/mutations"], "autoshape.update") {
		t.Fatalf("capabilities omit native mutation operations: %v", operations)
	}
	if len(got.Limitations) == 0 {
		t.Fatal("capabilities must disclose server limitations")
	}
}

func TestDiscoveryRoutesRequireGet(t *testing.T) {
	for _, path := range []string{healthPath, capabilitiesPath} {
		request := httptest.NewRequest(http.MethodPost, path, strings.NewReader("{}"))
		response := httptest.NewRecorder()
		newHandler(nil).ServeHTTP(response, request)
		res := response.Result()
		body, _ := io.ReadAll(res.Body)
		res.Body.Close()
		if res.StatusCode != http.StatusMethodNotAllowed || !strings.Contains(string(body), "GET required") {
			t.Fatalf("POST %s returned HTTP %d: %s", path, res.StatusCode, body)
		}
	}
}

func contains(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}

func TestMintArtifactThenJoinCollab(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(newHandler(store))
	t.Cleanup(server.Close)

	payload := []byte("%PDF-1.4 collab-room-proof")
	req, err := http.NewRequest(http.MethodPost, server.URL+artifacthttp.Path, bytes.NewReader(payload))
	if err != nil {
		t.Fatal(err)
	}
	req.Header.Set("Content-Type", "application/octet-stream")
	mintRes, err := http.DefaultClient.Do(req)
	if err != nil {
		t.Fatal(err)
	}
	defer mintRes.Body.Close()
	mintBody, err := io.ReadAll(mintRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	if mintRes.StatusCode != http.StatusOK && mintRes.StatusCode != http.StatusCreated {
		t.Fatalf("mint HTTP %d: %s", mintRes.StatusCode, mintBody)
	}
	if got := mintRes.Header.Get("Access-Control-Allow-Origin"); got != "*" {
		t.Fatalf("mint CORS origin %q", got)
	}
	artifactID := mintRes.Header.Get(xlsxhttp.HeaderArtifactID)
	if artifactID == "" || !strings.HasPrefix(artifactID, "art_") || strings.ContainsAny(artifactID, `/\`) || strings.Contains(artifactID, "..") {
		t.Fatalf("artifact id is not opaque: %q", artifactID)
	}
	var minted struct {
		ArtifactID string `json:"artifact_id"`
	}
	if err := json.Unmarshal(mintBody, &minted); err != nil {
		t.Fatal(err)
	}
	if minted.ArtifactID != artifactID {
		t.Fatalf("body artifact_id %q, header %q", minted.ArtifactID, artifactID)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(stored, payload) {
		t.Fatalf("store Get(%q) = %q, want minted bytes", artifactID, stored)
	}

	sessionRes, err := http.Post(server.URL+"/v1/collab/session", "application/json", bytes.NewReader(nil))
	if err != nil {
		t.Fatal(err)
	}
	defer sessionRes.Body.Close()
	sessionBody, err := io.ReadAll(sessionRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	if sessionRes.StatusCode != http.StatusOK {
		t.Fatalf("session HTTP %d: %s", sessionRes.StatusCode, sessionBody)
	}
	var sess struct {
		SessionID string `json:"session_id"`
	}
	if err := json.Unmarshal(sessionBody, &sess); err != nil {
		t.Fatal(err)
	}
	joinPayload, err := json.Marshal(map[string]any{
		"session_id": sess.SessionID,
		"path":       artifactID,
		"name":       "Mint",
	})
	if err != nil {
		t.Fatal(err)
	}
	joinRes, err := http.Post(server.URL+"/v1/collab/join", "application/json", bytes.NewReader(joinPayload))
	if err != nil {
		t.Fatal(err)
	}
	defer joinRes.Body.Close()
	joinBody, err := io.ReadAll(joinRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	if joinRes.StatusCode != http.StatusOK {
		t.Fatalf("join HTTP %d: %s", joinRes.StatusCode, joinBody)
	}
	var joined collab.JoinResult
	if err := json.Unmarshal(joinBody, &joined); err != nil {
		t.Fatal(err)
	}
	if joined.Room == "" || joined.Room == artifactID || strings.ContainsAny(joined.Room, `/\`) {
		t.Fatalf("room must be opaque, got %q", joined.Room)
	}
	if joined.File.Size != int64(len(payload)) || joined.File.Version == "" {
		t.Fatalf("join file meta missing: %+v", joined.File)
	}
}

func TestMintArtifactRejectsEmptyBody(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	request := httptest.NewRequest(http.MethodPost, artifacthttp.Path, http.NoBody)
	response := httptest.NewRecorder()
	newHandler(store).ServeHTTP(response, request)
	res := response.Result()
	defer res.Body.Close()
	body, _ := io.ReadAll(res.Body)
	if res.StatusCode != http.StatusBadRequest {
		t.Fatalf("empty mint HTTP %d, want 400: %s", res.StatusCode, body)
	}
	if !strings.Contains(string(body), "empty artifact body") {
		t.Fatalf("empty mint error %s", body)
	}
	if res.Header.Get(xlsxhttp.HeaderArtifactID) != "" {
		t.Fatalf("empty mint minted %q", res.Header.Get(xlsxhttp.HeaderArtifactID))
	}
}

func TestMintArtifactOptionsCORS(t *testing.T) {
	request := httptest.NewRequest(http.MethodOptions, artifacthttp.Path, nil)
	response := httptest.NewRecorder()
	newHandler(nil).ServeHTTP(response, request)
	res := response.Result()
	defer res.Body.Close()
	if res.StatusCode != http.StatusNoContent {
		t.Fatalf("OPTIONS HTTP %d, want 204", res.StatusCode)
	}
	if got := res.Header.Get("Access-Control-Allow-Origin"); got != "*" {
		t.Fatalf("CORS origin %q", got)
	}
}

func TestExtractMutateRoundTripHTTP(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(newHandler(store))
	t.Cleanup(server.Close)

	original := readHappyTree(t)
	extractRes, err := http.Post(server.URL+xlsxhttp.ExtractPath, xlsxhttp.XLSXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	defer extractRes.Body.Close()
	if extractRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(extractRes.Body)
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, body)
	}
	if media := extractRes.Header.Get("Content-Type"); media != xlsxpatch.NativeXLSXV2MediaType {
		t.Fatalf("extract content-type %q", media)
	}
	artifactID := extractRes.Header.Get(xlsxhttp.HeaderArtifactID)
	if artifactID == "" || strings.ContainsAny(artifactID, `/\`) || strings.Contains(artifactID, "..") || strings.Contains(artifactID, "xlsxpatch") {
		t.Fatalf("artifact id is not opaque: %q", artifactID)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	before, err := xlsxpatch.DecodeNativeWorkbookV2(encoded)
	if err != nil {
		t.Fatal(err)
	}

	reloadBody, _ := json.Marshal(map[string]string{"artifact_id": artifactID})
	reload, err := http.Post(server.URL+xlsxhttp.ExtractPath, "application/json", bytes.NewReader(reloadBody))
	if err != nil {
		t.Fatal(err)
	}
	defer reload.Body.Close()
	if reload.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(reload.Body)
		t.Fatalf("extract-by-id HTTP %d: %s", reload.StatusCode, body)
	}
	if reload.Header.Get(xlsxhttp.HeaderArtifactID) != artifactID {
		t.Fatalf("reload artifact id %q, want %q", reload.Header.Get(xlsxhttp.HeaderArtifactID), artifactID)
	}

	payload, err := json.Marshal(xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: "server-edit",
			SheetID:     before.Sheets[0].ID,
			Kind:        xlsxpatch.CellSetValue,
			Cell:        xlsxpatch.CellRef{Row: 0, Column: 1},
			Value:       "from-server",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	mutateRes := postMutation(t, server.URL, payload, before.Source.PackageSHA256, artifactID)
	defer mutateRes.Body.Close()
	if mutateRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(mutateRes.Body)
		t.Fatalf("mutations HTTP %d: %s", mutateRes.StatusCode, body)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderArtifactID) != artifactID {
		t.Fatalf("mutated artifact id %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderArtifactID), artifactID)
	}
	produced, err := io.ReadAll(mutateRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(stored, produced) {
		t.Fatal("store did not persist mutated package")
	}

	afterJSON, err := xlsxhttp.ExtractNativeJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := xlsxpatch.DecodeNativeWorkbookV2(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	cell := findCell(t, after, before.Sheets[0].ID, 0, 1)
	if cell.Value == nil || cell.Value.Text == nil || *cell.Value.Text != "from-server" {
		t.Fatalf("HTTP mutation did not land: %#v", cell.Value)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderRevision) != after.Revision {
		t.Fatalf("mutation revision header %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderRevision), after.Revision)
	}
}

func TestExtractDoesNotPersistInvalidPackage(t *testing.T) {
	dir := t.TempDir()
	store, err := fsstore.Open(dir)
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(newHandler(store))
	t.Cleanup(server.Close)
	res, err := http.Post(server.URL+xlsxhttp.ExtractPath, xlsxhttp.XLSXContentType, bytes.NewReader([]byte("not-a-zip")))
	if err != nil {
		t.Fatal(err)
	}
	body, _ := io.ReadAll(res.Body)
	res.Body.Close()
	if res.StatusCode != http.StatusBadRequest {
		t.Fatalf("status %d, want 400: %s", res.StatusCode, body)
	}
	if res.Header.Get(xlsxhttp.HeaderArtifactID) != "" {
		t.Fatalf("failed extract minted artifact id %q", res.Header.Get(xlsxhttp.HeaderArtifactID))
	}
	entries, err := os.ReadDir(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 0 {
		t.Fatalf("failed extract left %d store objects", len(entries))
	}
}

func TestExtractRejectsWorkspacePaths(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(newHandler(store))
	t.Cleanup(server.Close)
	for _, id := range []string{
		"../xlsxpatch/testdata/excel-authored/happy-tree.xlsx",
		"/workspace/injecting/happy-tree.xlsx",
		filepath.Join("excel-authored", "happy-tree.xlsx"),
	} {
		body, _ := json.Marshal(map[string]string{"artifact_id": id})
		res, err := http.Post(server.URL+xlsxhttp.ExtractPath, "application/json", bytes.NewReader(body))
		if err != nil {
			t.Fatal(err)
		}
		resp, _ := io.ReadAll(res.Body)
		res.Body.Close()
		if res.StatusCode != http.StatusBadRequest {
			t.Fatalf("path %q status %d, want 400: %s", id, res.StatusCode, resp)
		}
		if !strings.Contains(string(resp), "invalid artifact id") {
			t.Fatalf("path %q error %s", id, resp)
		}
	}
}

func postMutation(t *testing.T, url string, payload []byte, expectedRevision, artifactID string) *http.Response {
	t.Helper()
	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)
	if err := writer.WriteField("artifact_id", artifactID); err != nil {
		t.Fatal(err)
	}
	if err := writer.WriteField("payload", string(payload)); err != nil {
		t.Fatal(err)
	}
	if err := writer.WriteField("expected_revision", expectedRevision); err != nil {
		t.Fatal(err)
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	res, err := http.Post(url+xlsxhttp.MutationsPath, writer.FormDataContentType(), body)
	if err != nil {
		t.Fatal(err)
	}
	return res
}

func findCell(t *testing.T, workbook *xlsxpatch.NativeWorkbookV2, sheetID string, row, column int) xlsxpatch.NativeWorkbookCellV2 {
	t.Helper()
	for _, sheet := range workbook.Sheets {
		if sheet.ID != sheetID {
			continue
		}
		for _, cell := range sheet.Cells {
			if cell.Row == row && cell.Column == column {
				return cell
			}
		}
		t.Fatalf("cell r=%d c=%d missing on sheet %s", row, column, sheetID)
	}
	t.Fatalf("sheet %s missing", sheetID)
	return xlsxpatch.NativeWorkbookCellV2{}
}
