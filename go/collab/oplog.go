package collab

import (
	"context"
	"encoding/json"
	"errors"
)

// OpStore is the linear operation log behind a Hub. The in-memory
// implementation is enough for v1; sqlite or postgres can implement the same
// contract later. Append must reject a stale base without writing.
type OpStore interface {
	// Append assigns the next seq to ops. baseSeq must equal the current head
	// or ErrStaleBase is returned with the current head as the seq value.
	Append(ctx context.Context, roomKey, artifactID, clientID string, ops json.RawMessage, baseSeq int64) (int64, error)
	// Since returns entries with seq > since plus the head. ok=false means
	// the gap can no longer be filled: the caller must reload and rejoin.
	Since(ctx context.Context, roomKey string, since int64) (entries []OpEntry, head int64, ok bool, err error)
	// LogState reports the head and saved_seq (0/0 for an unknown room).
	LogState(ctx context.Context, roomKey string) (head, savedSeq int64, err error)
	// MarkSaved records that persisted bytes reflect every op up to seq
	// (clamped to the head, never backwards) and drops covered entries.
	MarkSaved(ctx context.Context, roomKey string, seq int64) error
	// Reset makes persisted bytes the new base: entries are cleared and the
	// head is bumped so a client still at the old head cannot Submit onto
	// the replaced snapshot.
	Reset(ctx context.Context, roomKey string) error
	// Forget drops retained ops. MemoryStore deletes the log; durable
	// implementations may no-op so unsaved ops survive empty rooms.
	Forget(ctx context.Context, roomKey string) error
}

// Submit appends ops for an artifact and returns the assigned seq, pushing the
// entry to every other peer. The submitter must be in the room. baseSeq is the
// head the submitter generated the ops against: when it isn't the current head,
// ErrStaleBase is returned along with that head.
func (h *Hub) Submit(ctx context.Context, conn *Conn, artifactID string, ops json.RawMessage, baseSeq int64) (int64, error) {
	if !conn.usable() {
		return 0, ErrInvalidConn
	}
	if !validArtifactID(artifactID) {
		return 0, ErrInvalidArtifact
	}
	if err := validateOps(ops); err != nil {
		return 0, err
	}
	ops = cloneRaw(ops)
	roomKey := RoomKey(artifactID)

	r := h.pinRoom(roomKey)
	if r == nil {
		return 0, ErrNotInRoom
	}
	defer r.mu.Unlock()
	if _, ok := r.peers[conn.clientID]; !ok {
		return 0, ErrNotInRoom
	}
	seq, err := h.store.Append(ctx, roomKey, artifactID, conn.clientID, ops, baseSeq)
	if err != nil {
		if errors.Is(err, ErrStaleBase) {
			return seq, ErrStaleBase
		}
		return 0, err
	}
	entry := OpEntry{Room: roomKey, Seq: seq, ClientID: conn.clientID, Ops: ops}
	for _, send := range r.sendersExcept(conn.clientID) {
		send(EventOp, entry)
	}
	return seq, nil
}

// Since returns retained entries after sinceSeq. The connection must be in the
// room. Reset is true when the ring no longer covers the gap, including
// sinceSeq past the head or a missing log.
func (h *Hub) Since(ctx context.Context, conn *Conn, artifactID string, sinceSeq int64) (SinceResult, error) {
	if !conn.usable() {
		return SinceResult{}, ErrInvalidConn
	}
	if !validArtifactID(artifactID) {
		return SinceResult{}, ErrInvalidArtifact
	}
	roomKey := RoomKey(artifactID)

	r := h.pinRoom(roomKey)
	if r == nil {
		return SinceResult{}, ErrNotInRoom
	}
	if _, ok := r.peers[conn.clientID]; !ok {
		r.mu.Unlock()
		return SinceResult{}, ErrNotInRoom
	}
	r.mu.Unlock()

	entries, head, ok, err := h.store.Since(ctx, roomKey, sinceSeq)
	if err != nil {
		return SinceResult{}, err
	}
	if entries == nil {
		entries = []OpEntry{}
	}
	return SinceResult{Ops: entries, Head: head, Reset: !ok}, nil
}

// Log reports the artifact's log position (0/0 for an unknown artifact).
func (h *Hub) Log(ctx context.Context, artifactID string) LogState {
	if !validArtifactID(artifactID) {
		return LogState{}
	}
	head, saved, err := h.store.LogState(ctx, RoomKey(artifactID))
	if err != nil {
		return LogState{}
	}
	return LogState{Seq: head, SavedSeq: saved}
}

// MarkSaved records that persisted bytes now reflect every op up to seq.
func (h *Hub) MarkSaved(ctx context.Context, artifactID string, seq int64) {
	if !validArtifactID(artifactID) {
		return
	}
	h.withStore(RoomKey(artifactID), func() {
		_ = h.store.MarkSaved(ctx, RoomKey(artifactID), seq)
	})
}

// ResetLog makes persisted bytes the new base for an artifact: nothing in the
// log applies any more. The head is bumped so a client still at the old head
// gets STALE_BASE and Since(oldHead) returns reset.
func (h *Hub) ResetLog(ctx context.Context, artifactID string) {
	if !validArtifactID(artifactID) {
		return
	}
	h.withStore(RoomKey(artifactID), func() {
		_ = h.store.Reset(ctx, RoomKey(artifactID))
	})
}
