package artifacthttp

import (
	"bytes"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

func TestMintOpaqueBytes(t *testing.T) {
	store := &recordingStore{}
	server := httptest.NewServer(New(store))
	t.Cleanup(server.Close)

	payload := []byte("%PDF-1.4 opaque proof")
	res, err := http.Post(server.URL+Path, "application/octet-stream", bytes.NewReader(payload))
	if err != nil {
		t.Fatal(err)
	}
	defer res.Body.Close()
	body, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatal(err)
	}
	if res.StatusCode != http.StatusCreated && res.StatusCode != http.StatusOK {
		t.Fatalf("mint HTTP %d: %s", res.StatusCode, body)
	}
	if got := res.Header.Get("Access-Control-Allow-Origin"); got != "*" {
		t.Fatalf("CORS origin %q", got)
	}
	id := res.Header.Get(xlsxhttp.HeaderArtifactID)
	if id == "" || !strings.HasPrefix(id, "art_") || strings.ContainsAny(id, `/\`) || strings.Contains(id, "..") {
		t.Fatalf("artifact id is not an opaque token: %q", id)
	}
	var got struct {
		ArtifactID string `json:"artifact_id"`
	}
	if err := json.Unmarshal(body, &got); err != nil {
		t.Fatal(err)
	}
	if got.ArtifactID != id {
		t.Fatalf("body artifact_id %q, header %q", got.ArtifactID, id)
	}
	if !bytes.Equal(store.created, payload) {
		t.Fatalf("store received %q, want %q", store.created, payload)
	}
}

func TestMintAcceptsUnlabeledAndJSONBytesWithoutSniffing(t *testing.T) {
	payload := []byte(`{"artifact_id":"not-a-request"}`)
	store := &recordingStore{id: "art_aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa"}
	handler := New(store)

	unlabeled := httptest.NewRequest(http.MethodPost, Path, bytes.NewReader(payload))
	unlabeledRes := httptest.NewRecorder()
	handler.ServeHTTP(unlabeledRes, unlabeled)
	if unlabeledRes.Code != http.StatusCreated && unlabeledRes.Code != http.StatusOK {
		t.Fatalf("unlabeled HTTP %d: %s", unlabeledRes.Code, unlabeledRes.Body.Bytes())
	}
	if !bytes.Equal(store.created, payload) {
		t.Fatalf("unlabeled store received %q", store.created)
	}

	store.created = nil
	jsonReq := httptest.NewRequest(http.MethodPost, Path, bytes.NewReader(payload))
	jsonReq.Header.Set("Content-Type", "application/json")
	jsonRes := httptest.NewRecorder()
	handler.ServeHTTP(jsonRes, jsonReq)
	if jsonRes.Code != http.StatusCreated && jsonRes.Code != http.StatusOK {
		t.Fatalf("json HTTP %d: %s", jsonRes.Code, jsonRes.Body.Bytes())
	}
	if !bytes.Equal(store.created, payload) {
		t.Fatalf("json content-type was sniffed; store received %q", store.created)
	}
}

func TestMintRejectsEmptyBody(t *testing.T) {
	store := &recordingStore{}
	response := httptest.NewRecorder()
	New(store).ServeHTTP(response, httptest.NewRequest(http.MethodPost, Path, http.NoBody))
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
	if store.created != nil {
		t.Fatal("empty mint called Create")
	}
}

func TestMintRejectsOversized(t *testing.T) {
	old := maxArtifactBytes
	maxArtifactBytes = 8
	t.Cleanup(func() { maxArtifactBytes = old })

	t.Run("http limit", func(t *testing.T) {
		store := &recordingStore{}
		response := httptest.NewRecorder()
		New(store).ServeHTTP(response, httptest.NewRequest(http.MethodPost, Path, bytes.NewReader([]byte("0123456789"))))
		res := response.Result()
		defer res.Body.Close()
		body, _ := io.ReadAll(res.Body)
		if res.StatusCode != http.StatusRequestEntityTooLarge {
			t.Fatalf("oversize HTTP %d, want 413: %s", res.StatusCode, body)
		}
		if store.created != nil {
			t.Fatal("oversize mint called Create")
		}
	})

	t.Run("store too large", func(t *testing.T) {
		store := &recordingStore{createErr: xlsxhttp.ErrArtifactTooLarge}
		response := httptest.NewRecorder()
		New(store).ServeHTTP(response, httptest.NewRequest(http.MethodPost, Path, bytes.NewReader([]byte("abc"))))
		res := response.Result()
		defer res.Body.Close()
		body, _ := io.ReadAll(res.Body)
		if res.StatusCode != http.StatusRequestEntityTooLarge || !strings.Contains(string(body), "artifact too large") {
			t.Fatalf("store oversize HTTP %d: %s", res.StatusCode, body)
		}
		if res.Header.Get(xlsxhttp.HeaderArtifactID) != "" {
			t.Fatalf("oversize mint minted %q", res.Header.Get(xlsxhttp.HeaderArtifactID))
		}
	})
}

func TestMintOptionsCORS(t *testing.T) {
	response := httptest.NewRecorder()
	New(nil).ServeHTTP(response, httptest.NewRequest(http.MethodOptions, Path, nil))
	res := response.Result()
	defer res.Body.Close()
	if res.StatusCode != http.StatusNoContent {
		t.Fatalf("OPTIONS HTTP %d, want 204", res.StatusCode)
	}
	if got := res.Header.Get("Access-Control-Allow-Origin"); got != "*" {
		t.Fatalf("CORS origin %q", got)
	}
}

func TestMintRequiresPOST(t *testing.T) {
	response := httptest.NewRecorder()
	New(&recordingStore{}).ServeHTTP(response, httptest.NewRequest(http.MethodGet, Path, nil))
	res := response.Result()
	body, _ := io.ReadAll(res.Body)
	res.Body.Close()
	if res.StatusCode != http.StatusMethodNotAllowed || !strings.Contains(string(body), "POST required") {
		t.Fatalf("GET mint HTTP %d: %s", res.StatusCode, body)
	}
}

type recordingStore struct {
	created   []byte
	id        string
	createErr error
}

func (s *recordingStore) Get(string) ([]byte, error) { return nil, xlsxhttp.ErrArtifactNotFound }

func (s *recordingStore) Put(string, []byte, []byte) error { return xlsxhttp.ErrArtifactNotFound }

func (s *recordingStore) Create(data []byte) (string, error) {
	if s.createErr != nil {
		return "", s.createErr
	}
	s.created = append([]byte(nil), data...)
	if s.id == "" {
		s.id = "art_0123456789abcdef0123456789abcdef"
	}
	return s.id, nil
}
