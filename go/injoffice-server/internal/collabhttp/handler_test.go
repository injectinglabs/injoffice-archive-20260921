package collabhttp

import (
	"bufio"
	"bytes"
	"encoding/json"
	"io"
	"net/http"
	"net/http/httptest"
	"strings"
	"testing"
	"time"

	"github.com/injectinglabs/injoffice/go/collab"
	"github.com/injectinglabs/injoffice/go/injoffice-server/internal/fsstore"
)

func TestJoinAndStaleBaseHTTP(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	artifact, err := store.Create([]byte("PK-collab-demo-bytes"))
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(New(collab.NewHub(), store))
	t.Cleanup(server.Close)

	alice := openSession(t, server.URL)
	bob := openSession(t, server.URL)

	aliceJoin := postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": alice.SessionID,
		"path":       artifact,
		"name":       "Alice",
		"client_id":  "forged-client",
		"user_id":    "forged-user",
	})
	if aliceJoin.Status != http.StatusOK {
		t.Fatalf("alice join HTTP %d: %s", aliceJoin.Status, aliceJoin.Body)
	}
	var aliceGot collab.JoinResult
	if err := json.Unmarshal(aliceJoin.Body, &aliceGot); err != nil {
		t.Fatal(err)
	}
	if aliceGot.Room == "" || aliceGot.Room == artifact || strings.ContainsAny(aliceGot.Room, `/\`) || strings.Contains(aliceGot.Room, "..") {
		t.Fatalf("room must be opaque, got %q", aliceGot.Room)
	}
	if aliceGot.Self.ClientID != alice.ClientID || aliceGot.Self.UserID != alice.UserID {
		t.Fatalf("join must ignore client-forged identity, got %+v want client=%s user=%s", aliceGot.Self, alice.ClientID, alice.UserID)
	}
	if aliceGot.Self.Name != "Alice" || aliceGot.Self.Color == "" {
		t.Fatalf("alice self %+v", aliceGot.Self)
	}
	if len(aliceGot.Peers) != 0 {
		t.Fatalf("first joiner should see no peers, got %+v", aliceGot.Peers)
	}
	if aliceGot.Log.Seq != 0 {
		t.Fatalf("fresh log seq %d", aliceGot.Log.Seq)
	}
	if aliceGot.File.Size == 0 || aliceGot.File.Version == "" {
		t.Fatalf("join file meta missing: %+v", aliceGot.File)
	}

	bobJoin := postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": bob.SessionID,
		"path":       artifact,
		"name":       "Bob",
	})
	if bobJoin.Status != http.StatusOK {
		t.Fatalf("bob join HTTP %d: %s", bobJoin.Status, bobJoin.Body)
	}
	var bobGot collab.JoinResult
	if err := json.Unmarshal(bobJoin.Body, &bobGot); err != nil {
		t.Fatal(err)
	}
	if bobGot.Room != aliceGot.Room {
		t.Fatalf("peers must share a room, alice=%q bob=%q", aliceGot.Room, bobGot.Room)
	}
	if len(bobGot.Peers) != 1 || bobGot.Peers[0].ClientID != alice.ClientID || bobGot.Peers[0].Name != "Alice" {
		t.Fatalf("bob should see alice, got %+v", bobGot.Peers)
	}

	ops := json.RawMessage(`[{"id":"sheet.mutation.set-range-values","params":{"plain":"json"}}]`)
	aliceSubmit := postJSON(t, server.URL+"/v1/collab/op/submit", map[string]any{
		"session_id": alice.SessionID,
		"path":       artifact,
		"base_seq":   0,
		"ops":        []map[string]any{{"id": "sheet.mutation.set-range-values", "params": map[string]any{"plain": "json"}}},
	})
	if aliceSubmit.Status != http.StatusOK {
		t.Fatalf("alice submit HTTP %d: %s", aliceSubmit.Status, aliceSubmit.Body)
	}
	var accepted struct {
		Seq int64 `json:"seq"`
	}
	if err := json.Unmarshal(aliceSubmit.Body, &accepted); err != nil {
		t.Fatal(err)
	}
	if accepted.Seq != 1 {
		t.Fatalf("first seq %d, want 1", accepted.Seq)
	}

	carol := openSession(t, server.URL)
	carolJoin := postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": carol.SessionID,
		"path":       artifact,
		"name":       "Carol",
	})
	if carolJoin.Status != http.StatusOK {
		t.Fatalf("carol join HTTP %d: %s", carolJoin.Status, carolJoin.Body)
	}
	var carolGot collab.JoinResult
	if err := json.Unmarshal(carolJoin.Body, &carolGot); err != nil {
		t.Fatal(err)
	}
	if carolGot.Log.Seq != 1 {
		t.Fatalf("late joiner log seq %d, want 1", carolGot.Log.Seq)
	}
	carolGap := postJSON(t, server.URL+"/v1/collab/op/since", map[string]any{
		"session_id": carol.SessionID,
		"path":       artifact,
		"since_seq":  0,
	})
	if carolGap.Status != http.StatusOK {
		t.Fatalf("carol since HTTP %d: %s", carolGap.Status, carolGap.Body)
	}
	var carolSince collab.SinceResult
	if err := json.Unmarshal(carolGap.Body, &carolSince); err != nil {
		t.Fatal(err)
	}
	if carolSince.Reset || carolSince.Head != 1 || len(carolSince.Ops) != 1 || carolSince.Ops[0].Seq != 1 {
		t.Fatalf("late joiner since %+v", carolSince)
	}

	stale := postJSON(t, server.URL+"/v1/collab/op/submit", map[string]any{
		"session_id": bob.SessionID,
		"path":       artifact,
		"base_seq":   0,
		"ops":        []map[string]any{{"id": "edit", "params": map[string]any{}}},
	})
	if stale.Status != http.StatusConflict {
		t.Fatalf("stale submit HTTP %d, want 409: %s", stale.Status, stale.Body)
	}
	if !strings.Contains(string(stale.Body), "STALE_BASE") {
		t.Fatalf("stale error must contain STALE_BASE, got %s", stale.Body)
	}
	var staleBody errorBody
	if err := json.Unmarshal(stale.Body, &staleBody); err != nil {
		t.Fatal(err)
	}
	if staleBody.Error != "STALE_BASE" || staleBody.Head != 1 {
		t.Fatalf("stale body %+v", staleBody)
	}

	gap := postJSON(t, server.URL+"/v1/collab/op/since", map[string]any{
		"session_id": bob.SessionID,
		"path":       artifact,
		"since_seq":  0,
	})
	if gap.Status != http.StatusOK {
		t.Fatalf("since HTTP %d: %s", gap.Status, gap.Body)
	}
	var since collab.SinceResult
	if err := json.Unmarshal(gap.Body, &since); err != nil {
		t.Fatal(err)
	}
	if since.Reset || since.Head != 1 || len(since.Ops) != 1 || since.Ops[0].Seq != 1 {
		t.Fatalf("since %+v", since)
	}
	if !bytes.Equal(since.Ops[0].Ops, ops) && !bytes.Contains(since.Ops[0].Ops, []byte("sheet.mutation.set-range-values")) {
		t.Fatalf("since ops %s", since.Ops[0].Ops)
	}

	retry := postJSON(t, server.URL+"/v1/collab/op/submit", map[string]any{
		"session_id": bob.SessionID,
		"path":       artifact,
		"base_seq":   since.Head,
		"ops":        []map[string]any{{"id": "edit", "params": map[string]any{}}},
	})
	if retry.Status != http.StatusOK {
		t.Fatalf("retry submit HTTP %d: %s", retry.Status, retry.Body)
	}
}

func TestSSEFanoutOp(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	artifact, err := store.Create([]byte("PK-collab-demo-bytes"))
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(New(collab.NewHub(), store))
	t.Cleanup(server.Close)

	alice := openSession(t, server.URL)
	bob := openSession(t, server.URL)
	if postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": alice.SessionID, "path": artifact, "name": "Alice",
	}).Status != http.StatusOK {
		t.Fatal("alice join")
	}
	if postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": bob.SessionID, "path": artifact, "name": "Bob",
	}).Status != http.StatusOK {
		t.Fatal("bob join")
	}

	req, err := http.NewRequest(http.MethodGet, server.URL+"/v1/collab/events?session_id="+bob.SessionID, nil)
	if err != nil {
		t.Fatal(err)
	}
	res, err := http.DefaultClient.Do(req)
	if err != nil {
		t.Fatal(err)
	}
	t.Cleanup(func() { _ = res.Body.Close() })
	if res.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(res.Body)
		t.Fatalf("events HTTP %d: %s", res.StatusCode, body)
	}
	if media := res.Header.Get("Content-Type"); !strings.HasPrefix(media, "text/event-stream") {
		t.Fatalf("events content-type %q", media)
	}

	frames := make(chan frame, 8)
	go func() {
		reader := bufio.NewReader(res.Body)
		for {
			line, err := reader.ReadString('\n')
			if err != nil {
				return
			}
			line = strings.TrimSpace(line)
			if !strings.HasPrefix(line, "data:") {
				continue
			}
			raw := strings.TrimSpace(strings.TrimPrefix(line, "data:"))
			var fr frame
			if err := json.Unmarshal([]byte(raw), &fr); err != nil {
				return
			}
			frames <- fr
		}
	}()

	submit := postJSON(t, server.URL+"/v1/collab/op/submit", map[string]any{
		"session_id": alice.SessionID,
		"path":       artifact,
		"base_seq":   0,
		"ops":        []map[string]any{{"id": "sheet.mutation.set-range-values", "params": map[string]any{"v": "hi"}}},
	})
	if submit.Status != http.StatusOK {
		t.Fatalf("alice submit HTTP %d: %s", submit.Status, submit.Body)
	}

	deadline := time.After(2 * time.Second)
	for {
		select {
		case fr := <-frames:
			if fr.Event != collab.EventOp {
				continue
			}
			payload, _ := json.Marshal(fr.Payload)
			if !bytes.Contains(payload, []byte(`"seq":1`)) && !bytes.Contains(payload, []byte(`"seq": 1`)) {
				t.Fatalf("collab.op payload %s", payload)
			}
			if !bytes.Contains(payload, []byte(alice.ClientID)) {
				t.Fatalf("collab.op missing submitter client_id: %s", payload)
			}
			return
		case <-deadline:
			t.Fatal("timed out waiting for collab.op SSE frame")
		}
	}
}

func TestCloseSessionPost(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	artifact, err := store.Create([]byte("PK-collab-demo-bytes"))
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(New(collab.NewHub(), store))
	t.Cleanup(server.Close)

	sess := openSession(t, server.URL)
	if postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": sess.SessionID, "path": artifact, "name": "Ada",
	}).Status != http.StatusOK {
		t.Fatal("join")
	}

	res, err := http.Post(server.URL+"/v1/collab/session/close?session_id="+sess.SessionID, "text/plain", bytes.NewReader(nil))
	if err != nil {
		t.Fatal(err)
	}
	_, _ = io.Copy(io.Discard, res.Body)
	res.Body.Close()
	if res.StatusCode != http.StatusNoContent {
		t.Fatalf("close HTTP %d", res.StatusCode)
	}

	again := postJSON(t, server.URL+"/v1/collab/join", map[string]any{
		"session_id": sess.SessionID, "path": artifact, "name": "Ada",
	})
	if again.Status != http.StatusUnauthorized {
		t.Fatalf("closed session join HTTP %d, want 401: %s", again.Status, again.Body)
	}
}

func TestJoinRejectsFilesystemPaths(t *testing.T) {
	store, err := fsstore.Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	server := httptest.NewServer(New(collab.NewHub(), store))
	t.Cleanup(server.Close)
	sess := openSession(t, server.URL)

	for _, path := range []string{
		"../xlsxpatch/testdata/excel-authored/happy-tree.xlsx",
		"/workspace/injecting/happy-tree.xlsx",
		"excel-authored/happy-tree.xlsx",
		"",
	} {
		res := postJSON(t, server.URL+"/v1/collab/join", map[string]any{
			"session_id": sess.SessionID,
			"path":       path,
			"name":       "Ada",
		})
		if res.Status != http.StatusBadRequest {
			t.Fatalf("path %q status %d, want 400: %s", path, res.Status, res.Body)
		}
		if strings.Contains(string(res.Body), "xlsxpatch") || strings.Contains(string(res.Body), "workspace") {
			t.Fatalf("error leaked a path for %q: %s", path, res.Body)
		}
	}
}

type sessionInfo struct {
	SessionID string `json:"session_id"`
	UserID    string `json:"user_id"`
	ClientID  string `json:"client_id"`
}

type httpResult struct {
	Status int
	Body   []byte
}

func openSession(t *testing.T, base string) sessionInfo {
	t.Helper()
	res, err := http.Post(base+"/v1/collab/session", "application/json", bytes.NewReader(nil))
	if err != nil {
		t.Fatal(err)
	}
	defer res.Body.Close()
	body, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatal(err)
	}
	if res.StatusCode != http.StatusOK {
		t.Fatalf("session HTTP %d: %s", res.StatusCode, body)
	}
	var info sessionInfo
	if err := json.Unmarshal(body, &info); err != nil {
		t.Fatal(err)
	}
	if info.SessionID == "" || info.ClientID == "" || info.UserID == "" {
		t.Fatalf("session response %+v", info)
	}
	if strings.ContainsAny(info.SessionID, `/\`) {
		t.Fatalf("session id looks like a path: %q", info.SessionID)
	}
	return info
}

func postJSON(t *testing.T, url string, payload map[string]any) httpResult {
	t.Helper()
	encoded, err := json.Marshal(payload)
	if err != nil {
		t.Fatal(err)
	}
	res, err := http.Post(url, "application/json", bytes.NewReader(encoded))
	if err != nil {
		t.Fatal(err)
	}
	defer res.Body.Close()
	body, err := io.ReadAll(res.Body)
	if err != nil {
		t.Fatal(err)
	}
	return httpResult{Status: res.StatusCode, Body: body}
}
