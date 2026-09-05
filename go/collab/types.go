package collab

import (
	"encoding/json"
	"errors"
)

// Event names pushed to room members. They are the collab.* frames in the
// collaboration protocol; a host transport should deliver only these.
const (
	EventPeerJoined  = "collab.peer.joined"
	EventPeerLeft    = "collab.peer.left"
	EventPresence    = "collab.presence"
	EventFileChanged = "collab.file.changed"
	EventOp          = "collab.op"
)

const (
	// MaxSelectionBytes bounds the opaque presence blob a peer may publish.
	MaxSelectionBytes = 4096
	// MaxOpBytes bounds one Submit payload.
	MaxOpBytes = 512 * 1024
	// MaxOpsPerBatch bounds the number of op records in one Submit.
	MaxOpsPerBatch = 256
	// MaxOpsPerRoom bounds retained log entries for one artifact.
	MaxOpsPerRoom = 5000
	// MaxRoomMembers bounds concurrent connections in one room.
	MaxRoomMembers = 64
	// MaxJSONDepth bounds nesting in ops and presence JSON.
	MaxJSONDepth = 32
	// MaxPeerNameRunes bounds the display name accepted on join.
	MaxPeerNameRunes = 64
	// MaxArtifactIDLen bounds an opaque artifact identifier.
	MaxArtifactIDLen = 512
)

var (
	// ErrStaleBase is returned by Submit when base_seq is not the current
	// head. The returned seq is the head so the client can catch up. The
	// error text is the stable protocol code.
	ErrStaleBase = errors.New("STALE_BASE")
	// ErrOpTooLarge is returned when a submission exceeds MaxOpBytes.
	ErrOpTooLarge = errors.New("op too large")
	// ErrBatchTooLarge is returned when a submission has more than MaxOpsPerBatch records.
	ErrBatchTooLarge = errors.New("op batch too large")
	// ErrInvalidOps is returned when ops are not a non-empty JSON array.
	ErrInvalidOps = errors.New("invalid ops")
	// ErrJSONTooDeep is returned when ops or presence nest past MaxJSONDepth.
	ErrJSONTooDeep = errors.New("json too deep")
	// ErrNotInRoom is returned when the connection is not a member of the room.
	ErrNotInRoom = errors.New("not in room")
	// ErrRoomFull is returned when a join would exceed MaxRoomMembers.
	ErrRoomFull = errors.New("room full")
	// ErrInvalidArtifact is returned when the artifact id is empty, too long, or contains NUL.
	ErrInvalidArtifact = errors.New("invalid artifact id")
	// ErrInvalidConn is returned when Join/Submit/etc. is called without a usable Conn.
	ErrInvalidConn = errors.New("invalid connection")
	// ErrInvalidSelection is returned when presence JSON is missing or malformed.
	ErrInvalidSelection = errors.New("invalid selection")
	// ErrSelectionTooLarge is returned when presence exceeds MaxSelectionBytes.
	ErrSelectionTooLarge = errors.New("selection too large")
)

// Sender delivers one event frame to a connection. It must not block: a host
// should drop or disconnect a slow consumer rather than grow an unbounded queue.
type Sender func(event string, payload any)

// PeerInfo is the wire shape of one peer (self or remote).
type PeerInfo struct {
	ClientID  string          `json:"client_id"`
	UserID    string          `json:"user_id"`
	Name      string          `json:"name"`
	Color     string          `json:"color"`
	Selection json.RawMessage `json:"selection,omitempty"`
	JoinedAt  int64           `json:"joined_at"`
}

// FileMeta is what a joiner learns about the persisted artifact bytes.
// The hub does not read storage; the host supplies this on Join.
type FileMeta struct {
	Version string `json:"version,omitempty"`
	Mtime   int64  `json:"mtime"`
	Size    int64  `json:"size"`
}

// JoinResult is the collab.join response.
type JoinResult struct {
	Room  string     `json:"room"`
	Self  PeerInfo   `json:"self"`
	Peers []PeerInfo `json:"peers"`
	File  FileMeta   `json:"file"`
	Log   LogState   `json:"log"`
}

// FileChange is what followers learn when persisted artifact bytes changed.
// Path is the opaque artifact id, never a resolved filesystem path.
type FileChange struct {
	Room     string `json:"room"`
	Path     string `json:"path"`
	Author   string `json:"author"`
	Origin   string `json:"origin,omitempty"`
	Version  string `json:"version,omitempty"`
	Mtime    int64  `json:"mtime"`
	Size     int64  `json:"size"`
	Action   string `json:"action,omitempty"`
	UserID   string `json:"user_id,omitempty"`
	UserName string `json:"user_name,omitempty"`
	SavedSeq int64  `json:"saved_seq"`
	Reset    bool   `json:"reset,omitempty"`
}

// OpEntry is one ordered submission as stored and as pushed to peers.
type OpEntry struct {
	Room     string          `json:"room"`
	Seq      int64           `json:"seq"`
	ClientID string          `json:"client_id"`
	Ops      json.RawMessage `json:"ops"`
}

// LogState is the log position a joiner learns.
type LogState struct {
	Seq      int64 `json:"seq"`
	SavedSeq int64 `json:"saved_seq"`
}

// SinceResult is the collab.op.since response. Reset means the requested
// history can no longer be replayed onto the client's snapshot.
type SinceResult struct {
	Ops   []OpEntry `json:"ops"`
	Head  int64     `json:"head"`
	Reset bool      `json:"reset"`
}
