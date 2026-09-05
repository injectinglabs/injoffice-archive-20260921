package collab

import (
	"crypto/rand"
	"crypto/sha256"
	"encoding/hex"
	"hash/fnv"
	"strings"
	"unicode/utf8"
)

// palette is indexed by a stable hash of the user id so one person has the
// same colour in every room and on every reconnect.
var palette = []string{
	"#0fa98f", "#7c6cf0", "#e0943a", "#d9534f", "#3b8bd6",
	"#c04a9b", "#5a9e3a", "#b8860b", "#2a9d8f", "#8e5ad6",
}

// Conn is one live connection. Reconnecting must allocate a new Conn — a new
// client_id. Fields are unexported so a host cannot inject client-supplied
// identity into the hub.
type Conn struct {
	clientID string
	userID   string
	send     Sender
}

// NewConn assigns a client_id. userID must come from the host's authenticator,
// never from a collab payload. send must not block.
func NewConn(userID string, send Sender) *Conn {
	return &Conn{clientID: newClientID(), userID: userID, send: send}
}

// ClientID is the server-assigned id of this live connection.
func (c *Conn) ClientID() string {
	if c == nil {
		return ""
	}
	return c.clientID
}

// UserID is the host-authenticated principal bound to this connection.
func (c *Conn) UserID() string {
	if c == nil {
		return ""
	}
	return c.userID
}

func (c *Conn) usable() bool {
	return c != nil && c.clientID != "" && c.send != nil
}

func newClientID() string {
	var b [16]byte
	if _, err := rand.Read(b[:]); err != nil {
		// crypto/rand failure is not expected in process; fall back to a
		// distinct empty-id so Join still rejects rather than panicking.
		return ""
	}
	return hex.EncodeToString(b[:])
}

// RoomKey is the canonical, opaque room id for one artifact. The input is
// hashed so a room key is never a filesystem path even if a host's artifact
// ids look like one.
func RoomKey(artifactID string) string {
	sum := sha256.Sum256([]byte("injoffice-collab-room\x00" + artifactID))
	return hex.EncodeToString(sum[:12])
}

// ColorFor picks the palette entry for a user id.
func ColorFor(userID string) string {
	h := fnv.New32a()
	_, _ = h.Write([]byte(userID))
	return palette[int(h.Sum32()%uint32(len(palette)))]
}

func sanitizeName(raw, userID string) string {
	n := strings.TrimSpace(raw)
	n = strings.Map(func(r rune) rune {
		if r < 0x20 || r == 0x7f || r == '<' || r == '>' {
			return -1
		}
		return r
	}, n)
	if utf8.RuneCountInString(n) > MaxPeerNameRunes {
		n = string([]rune(n)[:MaxPeerNameRunes])
	}
	if n == "" {
		if len(userID) > 8 {
			return "user-" + userID[:8]
		}
		if userID != "" {
			return "user-" + userID
		}
		return "guest"
	}
	return n
}

func validArtifactID(id string) bool {
	if id == "" || len(id) > MaxArtifactIDLen {
		return false
	}
	return strings.IndexByte(id, 0) < 0
}
