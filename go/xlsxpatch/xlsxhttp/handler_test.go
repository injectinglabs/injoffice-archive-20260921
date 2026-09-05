package xlsxhttp

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
	"sync"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const happyTreeSHA = "c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513"

func testdata(t *testing.T, parts ...string) string {
	t.Helper()
	path := filepath.Join(append([]string{"..", "testdata"}, parts...)...)
	if _, err := os.Stat(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func readHappyTree(t *testing.T) []byte {
	t.Helper()
	data, err := os.ReadFile(testdata(t, "excel-authored", "happy-tree.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(data)
	if hex.EncodeToString(sum[:]) != happyTreeSHA {
		t.Fatalf("happy-tree.xlsx sha256=%s, want %s", hex.EncodeToString(sum[:]), happyTreeSHA)
	}
	return data
}

func TestExtractAndMutationsHTTP(t *testing.T) {
	original := readHappyTree(t)
	server := httptest.NewServer(NewHandler(nil))
	t.Cleanup(server.Close)

	extractRes, err := http.Post(server.URL+ExtractPath, XLSXContentType, bytes.NewReader(original))
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
	if extractRes.Header.Get(HeaderArtifactID) != "" {
		t.Fatal("stateless helper must not mint artifact ids")
	}
	encoded, err := io.ReadAll(extractRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	before, err := xlsxpatch.DecodeNativeWorkbookV2(encoded)
	if err != nil {
		t.Fatal(err)
	}

	payload := oneCellPayload(t, before, "http-edit", 0, 1, "from-http")
	mutateRes := postMutation(t, server.URL, original, payload, ptr(before.Source.PackageSHA256), "")
	defer mutateRes.Body.Close()
	if mutateRes.StatusCode != http.StatusOK {
		resp, _ := io.ReadAll(mutateRes.Body)
		t.Fatalf("mutations HTTP %d: %s", mutateRes.StatusCode, resp)
	}
	produced, err := io.ReadAll(mutateRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	afterJSON, err := ExtractNativeJSON(produced)
	if err != nil {
		t.Fatal(err)
	}
	after, err := xlsxpatch.DecodeNativeWorkbookV2(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	cell := findCell(t, after, before.Sheets[0].ID, 0, 1)
	if cell.Value == nil || cell.Value.Text == nil || *cell.Value.Text != "from-http" {
		t.Fatalf("HTTP mutation did not land: %#v", cell.Value)
	}
	if mutateRes.Header.Get(HeaderRevision) != after.Revision {
		t.Fatalf("mutation revision header %q, want %q", mutateRes.Header.Get(HeaderRevision), after.Revision)
	}
}

func TestMutationsHTTPRejectsCASMismatches(t *testing.T) {
	original := readHappyTree(t)
	extracted, err := ExtractNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	before, err := xlsxpatch.DecodeNativeWorkbookV2(extracted)
	if err != nil {
		t.Fatal(err)
	}
	payload := oneCellPayload(t, before, "cas-edit", 0, 0, "must-not-commit")
	server := httptest.NewServer(NewHandler(nil))
	t.Cleanup(server.Close)

	t.Run("missing expected_revision", func(t *testing.T) {
		res := postMutation(t, server.URL, original, payload, nil, "")
		assertMutationRejected(t, res, "expected_revision is required")
	})
	t.Run("rev token as outer CAS", func(t *testing.T) {
		res := postMutation(t, server.URL, original, payload, ptr(before.Revision), "")
		assertMutationRejected(t, res, "sha256")
	})
	t.Run("stale outer sha256", func(t *testing.T) {
		stale := "sha256:" + strings.Repeat("0", 64)
		res := postMutation(t, server.URL, original, payload, &stale, "")
		assertMutationRejected(t, res, "stale outer revision")
	})
}

func TestArtifactStoreRoundTripHTTP(t *testing.T) {
	original := readHappyTree(t)
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)

	extractRes, err := http.Post(server.URL+ExtractPath, XLSXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	defer extractRes.Body.Close()
	if extractRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(extractRes.Body)
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, body)
	}
	artifactID := extractRes.Header.Get(HeaderArtifactID)
	if artifactID == "" || strings.ContainsAny(artifactID, `/\`) || strings.Contains(artifactID, "..") {
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

	reload := postJSON(t, server.URL+ExtractPath, map[string]string{"artifact_id": artifactID})
	defer reload.Body.Close()
	if reload.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(reload.Body)
		t.Fatalf("extract-by-id HTTP %d: %s", reload.StatusCode, body)
	}
	if reload.Header.Get(HeaderArtifactID) != artifactID {
		t.Fatalf("reload artifact id %q, want %q", reload.Header.Get(HeaderArtifactID), artifactID)
	}

	payload := oneCellPayload(t, before, "stored-edit", 0, 1, "from-store")
	mutateRes := postMutation(t, server.URL, nil, payload, ptr(before.Source.PackageSHA256), artifactID)
	defer mutateRes.Body.Close()
	if mutateRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(mutateRes.Body)
		t.Fatalf("mutations HTTP %d: %s", mutateRes.StatusCode, body)
	}
	if mutateRes.Header.Get(HeaderArtifactID) != artifactID {
		t.Fatalf("mutated artifact id %q, want %q", mutateRes.Header.Get(HeaderArtifactID), artifactID)
	}

	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	afterJSON, err := ExtractNativeJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := xlsxpatch.DecodeNativeWorkbookV2(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	cell := findCell(t, after, before.Sheets[0].ID, 0, 1)
	if cell.Value == nil || cell.Value.Text == nil || *cell.Value.Text != "from-store" {
		t.Fatalf("stored mutation did not land: %#v", cell.Value)
	}
}

func TestExtractDoesNotPersistInvalidPackage(t *testing.T) {
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	res, err := http.Post(server.URL+ExtractPath, XLSXContentType, bytes.NewReader([]byte("not-a-zip")))
	if err != nil {
		t.Fatal(err)
	}
	body, _ := io.ReadAll(res.Body)
	res.Body.Close()
	if res.StatusCode != http.StatusBadRequest {
		t.Fatalf("status %d, want 400: %s", res.StatusCode, body)
	}
	if res.Header.Get(HeaderArtifactID) != "" {
		t.Fatalf("failed extract minted artifact id %q", res.Header.Get(HeaderArtifactID))
	}
	if store.len() != 0 {
		t.Fatal("failed extract persisted an artifact")
	}
}

func TestStoreErrorStatusCodes(t *testing.T) {
	id := "art_" + strings.Repeat("a", 32)
	original := readHappyTree(t)
	cases := []struct {
		name   string
		store  Store
		do     func(t *testing.T, url string)
		status int
		want   string
	}{
		{
			name:   "invalid id",
			store:  stubStore{getErr: ErrInvalidArtifactID},
			status: http.StatusBadRequest,
			want:   "invalid artifact id",
			do: func(t *testing.T, url string) {
				res := postJSON(t, url+ExtractPath, map[string]string{"artifact_id": id})
				assertStatusContains(t, res, http.StatusBadRequest, "invalid artifact id")
			},
		},
		{
			name:   "not found",
			store:  stubStore{getErr: ErrArtifactNotFound},
			status: http.StatusNotFound,
			want:   "artifact not found",
			do: func(t *testing.T, url string) {
				res := postJSON(t, url+ExtractPath, map[string]string{"artifact_id": id})
				assertStatusContains(t, res, http.StatusNotFound, "artifact not found")
			},
		},
		{
			name:   "too large",
			store:  stubStore{createErr: ErrArtifactTooLarge},
			status: http.StatusRequestEntityTooLarge,
			want:   "artifact too large",
			do: func(t *testing.T, url string) {
				res, err := http.Post(url+ExtractPath, XLSXContentType, bytes.NewReader(original))
				if err != nil {
					t.Fatal(err)
				}
				assertStatusContains(t, res, http.StatusRequestEntityTooLarge, "artifact too large")
			},
		},
		{
			name:   "io",
			store:  stubStore{createErr: ErrArtifactIO},
			status: http.StatusInternalServerError,
			want:   "artifact store io",
			do: func(t *testing.T, url string) {
				res, err := http.Post(url+ExtractPath, XLSXContentType, bytes.NewReader(original))
				if err != nil {
					t.Fatal(err)
				}
				assertStatusContains(t, res, http.StatusInternalServerError, "artifact store io")
			},
		},
	}
	for _, tc := range cases {
		t.Run(tc.name, func(t *testing.T) {
			server := httptest.NewServer(NewHandler(tc.store))
			t.Cleanup(server.Close)
			tc.do(t, server.URL)
		})
	}
}

func TestConcurrentMutationsCAS(t *testing.T) {
	original := readHappyTree(t)
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	extractRes, err := http.Post(server.URL+ExtractPath, XLSXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	extractRes.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	if extractRes.StatusCode != http.StatusOK {
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, encoded)
	}
	artifactID := extractRes.Header.Get(HeaderArtifactID)
	before, err := xlsxpatch.DecodeNativeWorkbookV2(encoded)
	if err != nil {
		t.Fatal(err)
	}
	payloadA := oneCellPayload(t, before, "race-a", 0, 1, "race-a")
	payloadB := oneCellPayload(t, before, "race-b", 0, 2, "race-b")
	rev := before.Source.PackageSHA256
	results := make(chan int, 2)
	go func() {
		res := postMutation(t, server.URL, nil, payloadA, ptr(rev), artifactID)
		results <- res.StatusCode
		res.Body.Close()
	}()
	go func() {
		res := postMutation(t, server.URL, nil, payloadB, ptr(rev), artifactID)
		results <- res.StatusCode
		res.Body.Close()
	}()
	first, second := <-results, <-results
	ok, stale := 0, 0
	for _, code := range []int{first, second} {
		switch code {
		case http.StatusOK:
			ok++
		case http.StatusBadRequest:
			stale++
		default:
			t.Fatalf("unexpected mutation status %d", code)
		}
	}
	if ok != 1 || stale != 1 {
		t.Fatalf("CAS race statuses ok=%d stale=%d (got %d and %d)", ok, stale, first, second)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	afterJSON, err := ExtractNativeJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := xlsxpatch.DecodeNativeWorkbookV2(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	a := findCellOptional(after, before.Sheets[0].ID, 0, 1)
	b := findCellOptional(after, before.Sheets[0].ID, 0, 2)
	aHit := a != nil && a.Value != nil && a.Value.Text != nil && *a.Value.Text == "race-a"
	bHit := b != nil && b.Value != nil && b.Value.Text != nil && *b.Value.Text == "race-b"
	if aHit == bHit {
		t.Fatalf("store should keep exactly one racing mutation: a=%v b=%v", aHit, bHit)
	}
}

func TestArtifactIDRefusesPathsAndMissingStore(t *testing.T) {
	server := httptest.NewServer(NewHandler(nil))
	t.Cleanup(server.Close)
	for _, id := range []string{
		"../xlsxpatch/testdata/excel-authored/happy-tree.xlsx",
		"/workspace/injecting/happy-tree.xlsx",
		`C:\Users\docs\file.xlsx`,
	} {
		res := postJSON(t, server.URL+ExtractPath, map[string]string{"artifact_id": id})
		body, _ := io.ReadAll(res.Body)
		res.Body.Close()
		if res.StatusCode != http.StatusBadRequest {
			t.Fatalf("path %q status %d, want 400: %s", id, res.StatusCode, body)
		}
		if !strings.Contains(string(body), "artifact store") {
			t.Fatalf("path %q error %s", id, body)
		}
	}
}

func oneCellPayload(t *testing.T, before *xlsxpatch.NativeWorkbookV2, op string, row, column int, value string) []byte {
	t.Helper()
	payload, err := json.Marshal(xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: op,
			SheetID:     before.Sheets[0].ID,
			Kind:        xlsxpatch.CellSetValue,
			Cell:        xlsxpatch.CellRef{Row: row, Column: column},
			Value:       value,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	return payload
}

func postJSON(t *testing.T, url string, body any) *http.Response {
	t.Helper()
	encoded, err := json.Marshal(body)
	if err != nil {
		t.Fatal(err)
	}
	res, err := http.Post(url, "application/json", bytes.NewReader(encoded))
	if err != nil {
		t.Fatal(err)
	}
	return res
}

func postMutation(t *testing.T, url string, original, payload []byte, expectedRevision *string, artifactID string) *http.Response {
	t.Helper()
	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)
	if artifactID != "" {
		if err := writer.WriteField("artifact_id", artifactID); err != nil {
			t.Fatal(err)
		}
	} else {
		originalPart, err := writer.CreateFormFile("original", "happy-tree.xlsx")
		if err != nil {
			t.Fatal(err)
		}
		if _, err := originalPart.Write(original); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.WriteField("payload", string(payload)); err != nil {
		t.Fatal(err)
	}
	if expectedRevision != nil {
		if err := writer.WriteField("expected_revision", *expectedRevision); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	res, err := http.Post(url+MutationsPath, writer.FormDataContentType(), body)
	if err != nil {
		t.Fatal(err)
	}
	return res
}

func assertMutationRejected(t *testing.T, res *http.Response, want string) {
	t.Helper()
	defer res.Body.Close()
	body, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatal(err)
	}
	if res.StatusCode != http.StatusBadRequest {
		t.Fatalf("status %d, want 400: %s", res.StatusCode, body)
	}
	if res.Header.Get("Content-Type") == XLSXContentType || bytes.HasPrefix(body, []byte("PK")) {
		t.Fatalf("rejected mutation returned XLSX bytes (%d)", len(body))
	}
	if !strings.Contains(string(body), want) {
		t.Fatalf("error %q does not contain %q", body, want)
	}
}

func findCell(t *testing.T, workbook *xlsxpatch.NativeWorkbookV2, sheetID string, row, column int) xlsxpatch.NativeWorkbookCellV2 {
	t.Helper()
	if cell := findCellOptional(workbook, sheetID, row, column); cell != nil {
		return *cell
	}
	t.Fatalf("cell r=%d c=%d missing on sheet %s", row, column, sheetID)
	return xlsxpatch.NativeWorkbookCellV2{}
}

func findCellOptional(workbook *xlsxpatch.NativeWorkbookV2, sheetID string, row, column int) *xlsxpatch.NativeWorkbookCellV2 {
	for _, sheet := range workbook.Sheets {
		if sheet.ID != sheetID {
			continue
		}
		for i, cell := range sheet.Cells {
			if cell.Row == row && cell.Column == column {
				return &sheet.Cells[i]
			}
		}
	}
	return nil
}

func assertStatusContains(t *testing.T, res *http.Response, status int, want string) {
	t.Helper()
	defer res.Body.Close()
	body, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatal(err)
	}
	if res.StatusCode != status {
		t.Fatalf("status %d, want %d: %s", res.StatusCode, status, body)
	}
	if !strings.Contains(string(body), want) {
		t.Fatalf("error %q does not contain %q", body, want)
	}
}

func ptr(s string) *string { return &s }

type memStore struct {
	mu      sync.Mutex
	next    int
	objects map[string][]byte
}

func newMemStore() *memStore {
	return &memStore{objects: map[string][]byte{}}
}

func (s *memStore) len() int {
	s.mu.Lock()
	defer s.mu.Unlock()
	return len(s.objects)
}

func (s *memStore) Create(data []byte) (string, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.next++
	id := "art_" + strings.Repeat("0", 31) + "0123456789abcdef"[s.next%16:s.next%16+1]
	s.objects[id] = append([]byte(nil), data...)
	return id, nil
}

func (s *memStore) Get(id string) ([]byte, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	data, ok := s.objects[id]
	if !ok {
		return nil, ErrArtifactNotFound
	}
	return append([]byte(nil), data...), nil
}

func (s *memStore) Put(id string, data, expected []byte) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	current, ok := s.objects[id]
	if !ok {
		return ErrArtifactNotFound
	}
	if sha256.Sum256(current) != sha256.Sum256(expected) {
		return ErrArtifactStale
	}
	s.objects[id] = append([]byte(nil), data...)
	return nil
}

type stubStore struct {
	getErr    error
	createErr error
	putErr    error
}

func (s stubStore) Get(string) ([]byte, error) { return nil, s.getErr }

func (s stubStore) Create([]byte) (string, error) { return "", s.createErr }

func (s stubStore) Put(string, []byte, []byte) error { return s.putErr }
