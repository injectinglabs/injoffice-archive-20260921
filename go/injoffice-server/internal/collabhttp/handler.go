// Package collabhttp is the HTTP + SSE transport for go/collab.
// Artifact IDs are opaque tokens; they are never treated as filesystem paths.
package collabhttp

import (
	"crypto/rand"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"net/http"
	"strings"
	"sync"
	"sync/atomic"
	"time"

	"github.com/injectinglabs/injoffice/go/collab"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const (
	// Prefix is the URL tree for every collab RPC and the event stream.
	Prefix = "/v1/collab"

	maxBodyBytes    = collab.MaxOpBytes + 4096
	maxSessions     = 256
	eventBuffer     = 32
	keepAliveEvery  = 15 * time.Second
	idleTTL         = 30 * time.Second
	sseRetryMS      = 1000
	sessionIDBytes  = 16
	userIDBytes     = 8
	headerSessionID = "X-InjOffice-Collab-Session"
)

// Handler serves docs/COLLABORATION-PROTOCOL.md over HTTP RPCs plus SSE.
type Handler struct {
	hub   *collab.Hub
	store xlsxhttp.Store

	mu       sync.Mutex
	sessions map[string]*session
}

type session struct {
	id        string
	conn      *collab.Conn
	events    chan frame
	done      chan struct{}
	once      sync.Once
	streaming atomic.Bool
	lastUsed  atomic.Int64
}

type frame struct {
	Event   string `json:"event"`
	Payload any    `json:"payload"`
}

type errorBody struct {
	Error string `json:"error"`
	Head  int64  `json:"head,omitempty"`
}

// New returns a handler bound to hub and the artifact store used for FileMeta.
func New(hub *collab.Hub, store xlsxhttp.Store) *Handler {
	if hub == nil {
		hub = collab.NewHub()
	}
	return &Handler{
		hub:      hub,
		store:    store,
		sessions: map[string]*session{},
	}
}

// ServeHTTP routes /v1/collab/* RPCs and the SSE event stream.
func (h *Handler) ServeHTTP(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Access-Control-Allow-Origin", "*")
	w.Header().Set("Access-Control-Allow-Methods", "GET, POST, DELETE, OPTIONS")
	w.Header().Set("Access-Control-Allow-Headers", "Content-Type, "+headerSessionID)
	if r.Method == http.MethodOptions {
		w.WriteHeader(http.StatusNoContent)
		return
	}

	switch {
	case r.URL.Path == Prefix+"/session" && r.Method == http.MethodPost:
		h.createSession(w, r)
	case r.URL.Path == Prefix+"/session" && r.Method == http.MethodGet:
		h.getSession(w, r)
	case r.URL.Path == Prefix+"/session" && r.Method == http.MethodDelete:
		h.closeSession(w, r)
	case r.URL.Path == Prefix+"/session/close" && r.Method == http.MethodPost:
		h.closeSession(w, r)
	case r.URL.Path == Prefix+"/events" && r.Method == http.MethodGet:
		h.events(w, r)
	case r.URL.Path == Prefix+"/join" && r.Method == http.MethodPost:
		h.join(w, r)
	case r.URL.Path == Prefix+"/leave" && r.Method == http.MethodPost:
		h.leave(w, r)
	case r.URL.Path == Prefix+"/presence" && r.Method == http.MethodPost:
		h.presence(w, r)
	case r.URL.Path == Prefix+"/op/submit" && r.Method == http.MethodPost:
		h.submit(w, r)
	case r.URL.Path == Prefix+"/op/since" && r.Method == http.MethodPost:
		h.since(w, r)
	default:
		writeError(w, http.StatusNotFound, errors.New("not found"))
	}
}

func (h *Handler) createSession(w http.ResponseWriter, r *http.Request) {
	h.expireIdle()
	h.mu.Lock()
	if len(h.sessions) >= maxSessions {
		h.mu.Unlock()
		writeError(w, http.StatusServiceUnavailable, errors.New("too many sessions"))
		return
	}
	h.mu.Unlock()

	id, err := newToken("ses_", sessionIDBytes)
	if err != nil {
		writeError(w, http.StatusInternalServerError, errors.New("session mint failed"))
		return
	}
	userID, err := newToken("guest_", userIDBytes)
	if err != nil {
		writeError(w, http.StatusInternalServerError, errors.New("session mint failed"))
		return
	}

	sess := &session{
		id:     id,
		events: make(chan frame, eventBuffer),
		done:   make(chan struct{}),
	}
	sess.conn = collab.NewConn(userID, func(event string, payload any) {
		select {
		case sess.events <- frame{Event: event, Payload: payload}:
		default:
			sess.kill()
		}
	})

	h.mu.Lock()
	if len(h.sessions) >= maxSessions {
		h.mu.Unlock()
		writeError(w, http.StatusServiceUnavailable, errors.New("too many sessions"))
		return
	}
	h.sessions[id] = sess
	h.mu.Unlock()
	sess.touch()

	writeJSON(w, http.StatusOK, map[string]string{
		"session_id": id,
		"user_id":    sess.conn.UserID(),
		"client_id":  sess.conn.ClientID(),
	})
}

func (h *Handler) getSession(w http.ResponseWriter, r *http.Request) {
	sess, err := h.lookup(sessionIDFrom(r, ""))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	writeJSON(w, http.StatusOK, map[string]string{
		"session_id": sess.id,
		"user_id":    sess.conn.UserID(),
		"client_id":  sess.conn.ClientID(),
	})
}

func (h *Handler) closeSession(w http.ResponseWriter, r *http.Request) {
	sess, err := h.lookup(sessionIDFrom(r, ""))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	h.drop(sess)
	w.WriteHeader(http.StatusNoContent)
}

func (h *Handler) events(w http.ResponseWriter, r *http.Request) {
	sess, err := h.lookup(sessionIDFrom(r, ""))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	flusher, ok := w.(http.Flusher)
	if !ok {
		writeError(w, http.StatusInternalServerError, errors.New("streaming unsupported"))
		return
	}

	w.Header().Set("Content-Type", "text/event-stream")
	w.Header().Set("Cache-Control", "no-cache")
	w.Header().Set("Connection", "keep-alive")
	w.Header().Set("X-Accel-Buffering", "no")
	w.WriteHeader(http.StatusOK)
	if _, err := fmt.Fprintf(w, "retry: %d\n\n", sseRetryMS); err != nil {
		return
	}
	flusher.Flush()

	sess.streaming.Store(true)
	ticker := time.NewTicker(keepAliveEvery)
	defer ticker.Stop()
	defer func() {
		sess.streaming.Store(false)
		sess.touch()
	}()

	for {
		select {
		case <-r.Context().Done():
			return
		case <-sess.done:
			return
		case <-ticker.C:
			sess.touch()
			if _, err := io.WriteString(w, ": keepalive\n\n"); err != nil {
				return
			}
			flusher.Flush()
		case fr, ok := <-sess.events:
			if !ok {
				return
			}
			payload, err := json.Marshal(fr)
			if err != nil {
				continue
			}
			if _, err := fmt.Fprintf(w, "data: %s\n\n", payload); err != nil {
				return
			}
			flusher.Flush()
		}
	}
}

func (h *Handler) join(w http.ResponseWriter, r *http.Request) {
	var req struct {
		SessionID string `json:"session_id"`
		Path      string `json:"path"`
		Name      string `json:"name"`
	}
	if err := decodeJSON(w, r, &req); err != nil {
		writeError(w, http.StatusBadRequest, err)
		return
	}
	sess, err := h.lookup(sessionIDFrom(r, req.SessionID))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	file, err := h.fileMeta(req.Path)
	if err != nil {
		writeStoreError(w, err)
		return
	}
	got, err := h.hub.Join(sess.conn, req.Path, req.Name, file)
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	writeJSON(w, http.StatusOK, got)
}

func (h *Handler) leave(w http.ResponseWriter, r *http.Request) {
	var req struct {
		SessionID string `json:"session_id"`
		Path      string `json:"path"`
	}
	if err := decodeJSON(w, r, &req); err != nil {
		writeError(w, http.StatusBadRequest, err)
		return
	}
	sess, err := h.lookup(sessionIDFrom(r, req.SessionID))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	h.hub.Leave(sess.conn, req.Path)
	w.WriteHeader(http.StatusNoContent)
}

func (h *Handler) presence(w http.ResponseWriter, r *http.Request) {
	var req struct {
		SessionID string          `json:"session_id"`
		Path      string          `json:"path"`
		Selection json.RawMessage `json:"selection"`
	}
	if err := decodeJSON(w, r, &req); err != nil {
		writeError(w, http.StatusBadRequest, err)
		return
	}
	sess, err := h.lookup(sessionIDFrom(r, req.SessionID))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	if err := h.hub.Presence(sess.conn, req.Path, req.Selection); err != nil {
		writeHubError(w, err, 0)
		return
	}
	w.WriteHeader(http.StatusNoContent)
}

func (h *Handler) submit(w http.ResponseWriter, r *http.Request) {
	var req struct {
		SessionID string          `json:"session_id"`
		Path      string          `json:"path"`
		BaseSeq   int64           `json:"base_seq"`
		Ops       json.RawMessage `json:"ops"`
	}
	if err := decodeJSON(w, r, &req); err != nil {
		writeError(w, http.StatusBadRequest, err)
		return
	}
	sess, err := h.lookup(sessionIDFrom(r, req.SessionID))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	seq, err := h.hub.Submit(r.Context(), sess.conn, req.Path, req.Ops, req.BaseSeq)
	if err != nil {
		writeHubError(w, err, seq)
		return
	}
	writeJSON(w, http.StatusOK, map[string]int64{"seq": seq})
}

func (h *Handler) since(w http.ResponseWriter, r *http.Request) {
	var req struct {
		SessionID string `json:"session_id"`
		Path      string `json:"path"`
		SinceSeq  int64  `json:"since_seq"`
	}
	if err := decodeJSON(w, r, &req); err != nil {
		writeError(w, http.StatusBadRequest, err)
		return
	}
	sess, err := h.lookup(sessionIDFrom(r, req.SessionID))
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	got, err := h.hub.Since(r.Context(), sess.conn, req.Path, req.SinceSeq)
	if err != nil {
		writeHubError(w, err, 0)
		return
	}
	writeJSON(w, http.StatusOK, got)
}

func (h *Handler) fileMeta(artifactID string) (collab.FileMeta, error) {
	if looksLikePath(artifactID) {
		return collab.FileMeta{}, xlsxhttp.ErrInvalidArtifactID
	}
	if h.store == nil {
		return collab.FileMeta{}, fmt.Errorf("artifact_id requires a configured artifact store")
	}
	data, err := h.store.Get(artifactID)
	if err != nil {
		return collab.FileMeta{}, err
	}
	sum := sha256.Sum256(data)
	return collab.FileMeta{
		Version: hex.EncodeToString(sum[:]),
		Size:    int64(len(data)),
	}, nil
}

func (h *Handler) lookup(id string) (*session, error) {
	if id == "" {
		return nil, collab.ErrInvalidConn
	}
	h.expireIdle()
	h.mu.Lock()
	defer h.mu.Unlock()
	sess := h.sessions[id]
	if sess == nil {
		return nil, collab.ErrInvalidConn
	}
	select {
	case <-sess.done:
		return nil, collab.ErrInvalidConn
	default:
		sess.touch()
		return sess, nil
	}
}

func (h *Handler) expireIdle() {
	now := time.Now()
	h.mu.Lock()
	var stale []*session
	for _, sess := range h.sessions {
		if sess.streaming.Load() {
			continue
		}
		last := sess.lastUsed.Load()
		if last == 0 || now.Sub(time.Unix(0, last)) <= idleTTL {
			continue
		}
		stale = append(stale, sess)
	}
	h.mu.Unlock()
	for _, sess := range stale {
		h.drop(sess)
	}
}

func (s *session) touch() {
	s.lastUsed.Store(time.Now().UnixNano())
}

func (h *Handler) drop(sess *session) {
	if sess == nil {
		return
	}
	sess.kill()
	h.mu.Lock()
	delete(h.sessions, sess.id)
	h.mu.Unlock()
	h.hub.Disconnect(sess.conn)
}

func (s *session) kill() {
	s.once.Do(func() { close(s.done) })
}

func sessionIDFrom(r *http.Request, bodyID string) string {
	if id := strings.TrimSpace(bodyID); id != "" {
		return id
	}
	if id := strings.TrimSpace(r.Header.Get(headerSessionID)); id != "" {
		return id
	}
	return strings.TrimSpace(r.URL.Query().Get("session_id"))
}

func looksLikePath(id string) bool {
	if id == "" {
		return true
	}
	if strings.ContainsAny(id, `/\`) || strings.Contains(id, "..") || strings.IndexByte(id, 0) >= 0 {
		return true
	}
	return false
}

func decodeJSON(w http.ResponseWriter, r *http.Request, dest any) error {
	r.Body = http.MaxBytesReader(w, r.Body, maxBodyBytes)
	dec := json.NewDecoder(r.Body)
	if err := dec.Decode(dest); err != nil {
		return err
	}
	return nil
}

func writeHubError(w http.ResponseWriter, err error, head int64) {
	switch {
	case errors.Is(err, collab.ErrStaleBase):
		writeJSON(w, http.StatusConflict, errorBody{Error: collab.ErrStaleBase.Error(), Head: head})
	case errors.Is(err, collab.ErrInvalidConn):
		writeError(w, http.StatusUnauthorized, err)
	case errors.Is(err, collab.ErrNotInRoom):
		writeError(w, http.StatusForbidden, err)
	case errors.Is(err, collab.ErrRoomFull):
		writeError(w, http.StatusTooManyRequests, err)
	case errors.Is(err, collab.ErrOpTooLarge), errors.Is(err, collab.ErrBatchTooLarge), errors.Is(err, collab.ErrSelectionTooLarge):
		writeError(w, http.StatusRequestEntityTooLarge, err)
	case errors.Is(err, collab.ErrInvalidArtifact):
		writeError(w, http.StatusBadRequest, err)
	default:
		writeError(w, http.StatusBadRequest, err)
	}
}

func writeStoreError(w http.ResponseWriter, err error) {
	switch {
	case errors.Is(err, xlsxhttp.ErrArtifactNotFound):
		writeError(w, http.StatusNotFound, err)
	case errors.Is(err, xlsxhttp.ErrInvalidArtifactID):
		writeError(w, http.StatusBadRequest, err)
	default:
		writeError(w, http.StatusBadRequest, err)
	}
}

func writeError(w http.ResponseWriter, status int, err error) {
	msg := "error"
	if err != nil {
		msg = err.Error()
	}
	writeJSON(w, status, errorBody{Error: msg})
}

func writeJSON(w http.ResponseWriter, status int, body any) {
	w.Header().Set("Content-Type", "application/json")
	w.WriteHeader(status)
	_ = json.NewEncoder(w).Encode(body)
}

func newToken(prefix string, n int) (string, error) {
	raw := make([]byte, n)
	if _, err := rand.Read(raw); err != nil {
		return "", err
	}
	return prefix + hex.EncodeToString(raw), nil
}
