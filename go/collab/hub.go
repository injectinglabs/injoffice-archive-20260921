package collab

import (
	"context"
	"encoding/json"
	"sort"
	"sync"
	"time"
)

type peer struct {
	info PeerInfo
	send Sender
}

type room struct {
	mu    sync.Mutex
	key   string
	peers map[string]*peer // by client id
}

// Hub holds every in-process room. Membership and seq assignment are
// single-writer per room so fan-out cannot race *peer mutations. The global
// mutex only protects the room index; it is not held across store IO or send.
type Hub struct {
	mu       sync.Mutex
	rooms    map[string]*room
	byClient map[string]map[string]struct{} // client id → room keys
	now      func() time.Time
	store    OpStore
}

// NewHub returns a hub with an in-memory op log.
func NewHub() *Hub {
	return NewHubWithStore(NewMemoryStore())
}

// NewHubWithStore returns a hub that delegates seq assignment and replay to
// store. A nil store uses the in-memory implementation.
func NewHubWithStore(store OpStore) *Hub {
	if store == nil {
		store = NewMemoryStore()
	}
	return &Hub{
		rooms:    map[string]*room{},
		byClient: map[string]map[string]struct{}{},
		now:      time.Now,
		store:    store,
	}
}

// pinRoom locks hub then room and releases the hub lock. Caller must Unlock
// the room. Nil means the room does not exist.
func (h *Hub) pinRoom(roomKey string) *room {
	h.mu.Lock()
	r := h.rooms[roomKey]
	if r == nil {
		h.mu.Unlock()
		return nil
	}
	r.mu.Lock()
	h.mu.Unlock()
	return r
}

// Join adds (or refreshes) a connection in the artifact's room. Repeated joins
// for the same connection and artifact are idempotent. The host must already
// have authorized this user for the artifact; the hub does not resolve paths.
func (h *Hub) Join(conn *Conn, artifactID, name string, file FileMeta) (JoinResult, error) {
	if !conn.usable() {
		return JoinResult{}, ErrInvalidConn
	}
	if !validArtifactID(artifactID) {
		return JoinResult{}, ErrInvalidArtifact
	}
	roomKey := RoomKey(artifactID)
	info := PeerInfo{
		ClientID: conn.clientID,
		UserID:   conn.userID,
		Name:     sanitizeName(name, conn.userID),
		Color:    ColorFor(conn.userID),
	}

	h.mu.Lock()
	r := h.rooms[roomKey]
	if r == nil {
		r = &room{key: roomKey, peers: map[string]*peer{}}
		h.rooms[roomKey] = r
	}
	r.mu.Lock()

	fresh := false
	if existing, ok := r.peers[info.ClientID]; ok {
		info.JoinedAt = existing.info.JoinedAt
		if info.Selection == nil {
			info.Selection = existing.info.Selection
		}
		existing.info = info
		existing.send = conn.send
	} else if len(r.peers) >= MaxRoomMembers {
		r.mu.Unlock()
		h.mu.Unlock()
		return JoinResult{}, ErrRoomFull
	} else {
		fresh = true
		info.JoinedAt = h.now().UnixMilli()
		r.peers[info.ClientID] = &peer{info: info, send: conn.send}
	}
	if h.byClient[info.ClientID] == nil {
		h.byClient[info.ClientID] = map[string]struct{}{}
	}
	h.byClient[info.ClientID][roomKey] = struct{}{}
	self := clonePeerInfo(r.peers[info.ClientID].info)
	peers := r.infosExcept(info.ClientID)
	senders := r.sendersExcept(info.ClientID)
	h.mu.Unlock()
	defer r.mu.Unlock()

	if fresh {
		payload := map[string]any{"room": roomKey, "peer": self}
		for _, send := range senders {
			send(EventPeerJoined, payload)
		}
	}
	return JoinResult{
		Room:  roomKey,
		Self:  self,
		Peers: peers,
		File:  file,
		Log:   h.Log(context.Background(), artifactID),
	}, nil
}

// Leave removes a connection from one artifact's room and tells the others.
func (h *Hub) Leave(conn *Conn, artifactID string) {
	if conn == nil || !validArtifactID(artifactID) {
		return
	}
	h.leaveRoom(RoomKey(artifactID), conn.clientID)
}

func (h *Hub) leaveRoom(roomKey, clientID string) {
	h.mu.Lock()
	r := h.rooms[roomKey]
	if r == nil {
		h.mu.Unlock()
		return
	}
	r.mu.Lock()
	senders, left, empty := h.removePeer(r, clientID)
	if empty {
		// Under h.mu so Join cannot recreate the room / Append until the
		// old log is gone. Forgetting after Unlock races a new session.
		_ = h.store.Forget(context.Background(), roomKey)
	}
	h.mu.Unlock()
	if left {
		payload := map[string]any{"room": roomKey, "client_id": clientID}
		for _, send := range senders {
			send(EventPeerLeft, payload)
		}
	}
	r.mu.Unlock()
}

// Disconnect drops a connection from every room — the disconnect hook.
// Protocol: disconnecting must have the same effect as leave even when no
// explicit leave arrives.
func (h *Hub) Disconnect(conn *Conn) {
	if conn == nil {
		return
	}
	h.LeaveAll(conn.clientID)
}

// LeaveAll drops a client id from every room. Prefer Disconnect with a Conn.
func (h *Hub) LeaveAll(clientID string) {
	if clientID == "" {
		return
	}
	h.mu.Lock()
	keys := make([]string, 0, len(h.byClient[clientID]))
	for k := range h.byClient[clientID] {
		keys = append(keys, k)
	}
	type pending struct {
		room    string
		r       *room
		senders []Sender
	}
	var notify []pending
	for _, k := range keys {
		r := h.rooms[k]
		if r == nil {
			continue
		}
		r.mu.Lock()
		senders, left, empty := h.removePeer(r, clientID)
		if empty {
			_ = h.store.Forget(context.Background(), k)
		}
		if left {
			notify = append(notify, pending{k, r, senders})
		} else {
			r.mu.Unlock()
		}
	}
	h.mu.Unlock()
	for _, n := range notify {
		payload := map[string]any{"room": n.room, "client_id": clientID}
		for _, send := range n.senders {
			send(EventPeerLeft, payload)
		}
		n.r.mu.Unlock()
	}
}

// Presence records a peer's selection and fans it out. Selection is opaque
// JSON, size- and depth-limited, and does not advance the operation sequence.
func (h *Hub) Presence(conn *Conn, artifactID string, selection json.RawMessage) error {
	if !conn.usable() {
		return ErrInvalidConn
	}
	if !validArtifactID(artifactID) {
		return ErrInvalidArtifact
	}
	if err := validateSelection(selection); err != nil {
		return err
	}
	selection = cloneRaw(selection)
	roomKey := RoomKey(artifactID)

	r := h.pinRoom(roomKey)
	if r == nil {
		return ErrNotInRoom
	}
	defer r.mu.Unlock()
	p, ok := r.peers[conn.clientID]
	if !ok {
		return ErrNotInRoom
	}
	p.info.Selection = selection
	senders := r.sendersExcept(conn.clientID)
	payload := map[string]any{
		"room":      roomKey,
		"client_id": conn.clientID,
		"selection": json.RawMessage(selection),
	}
	for _, send := range senders {
		send(EventPresence, payload)
	}
	return nil
}

// FileChanged tells every peer of the artifact's room (except Origin) that
// persisted bytes changed. The host should call MarkSaved or ResetLog first
// when the write is tied to the op log. No-op when nobody has the artifact open.
func (h *Hub) FileChanged(ctx context.Context, artifactID string, change FileChange) {
	if !validArtifactID(artifactID) {
		return
	}
	key := RoomKey(artifactID)
	change.Room = key
	change.Path = artifactID
	if change.Action == "" {
		change.Action = "saved"
	}
	st := h.Log(ctx, artifactID)
	change.SavedSeq = st.SavedSeq

	r := h.pinRoom(key)
	if r == nil {
		return
	}
	defer r.mu.Unlock()
	for _, send := range r.sendersExcept(change.Origin) {
		send(EventFileChanged, change)
	}
}

// NotifyFileChanged updates the log watermark then fans file.changed.
// Reset clears replay history (wholesale replace). SavedSeq > 0 marks the
// persisted bytes as covering that seq.
func (h *Hub) NotifyFileChanged(ctx context.Context, artifactID string, change FileChange) {
	if change.Reset {
		h.ResetLog(ctx, artifactID)
	} else if change.SavedSeq > 0 {
		h.MarkSaved(ctx, artifactID, change.SavedSeq)
	}
	h.FileChanged(ctx, artifactID, change)
}

// Peers lists the peers of an artifact's room (for tests and diagnostics).
func (h *Hub) Peers(artifactID string) []PeerInfo {
	r := h.pinRoom(RoomKey(artifactID))
	if r == nil {
		return nil
	}
	defer r.mu.Unlock()
	return r.infosExcept("")
}

// Forget drops retained ops for an artifact. MemoryStore deletes the log;
// a durable store may no-op so unsaved ops survive empty rooms.
func (h *Hub) Forget(ctx context.Context, artifactID string) {
	if !validArtifactID(artifactID) {
		return
	}
	h.withStore(RoomKey(artifactID), func() {
		_ = h.store.Forget(ctx, RoomKey(artifactID))
	})
}

func (h *Hub) withStore(roomKey string, fn func()) {
	if r := h.pinRoom(roomKey); r != nil {
		defer r.mu.Unlock()
	}
	fn()
}

// removePeer drops clientID from r. Caller holds hub.mu and r.mu.
func (h *Hub) removePeer(r *room, clientID string) (senders []Sender, left, empty bool) {
	if _, ok := r.peers[clientID]; !ok {
		return nil, false, false
	}
	delete(r.peers, clientID)
	if set := h.byClient[clientID]; set != nil {
		delete(set, r.key)
		if len(set) == 0 {
			delete(h.byClient, clientID)
		}
	}
	senders = r.sendersExcept("")
	empty = len(r.peers) == 0
	if empty {
		delete(h.rooms, r.key)
	}
	return senders, true, empty
}

func (r *room) sendersExcept(clientID string) []Sender {
	out := make([]Sender, 0, len(r.peers))
	for id, p := range r.peers {
		if id == clientID {
			continue
		}
		out = append(out, p.send)
	}
	return out
}

func (r *room) infosExcept(clientID string) []PeerInfo {
	out := make([]PeerInfo, 0, len(r.peers))
	for id, p := range r.peers {
		if id == clientID {
			continue
		}
		out = append(out, clonePeerInfo(p.info))
	}
	sort.Slice(out, func(i, j int) bool { return out[i].JoinedAt < out[j].JoinedAt })
	return out
}

func clonePeerInfo(info PeerInfo) PeerInfo {
	info.Selection = cloneRaw(info.Selection)
	return info
}
