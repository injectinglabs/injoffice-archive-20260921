package collab

import (
	"context"
	"encoding/json"
	"strings"
	"sync"
	"testing"
	"time"
)

type recorder struct {
	mu     sync.Mutex
	events []string
	last   map[string]any
}

func (r *recorder) sender() Sender {
	return func(event string, payload any) {
		r.mu.Lock()
		defer r.mu.Unlock()
		r.events = append(r.events, event)
		b, _ := json.Marshal(payload)
		var m map[string]any
		_ = json.Unmarshal(b, &m)
		r.last = m
	}
}

func (r *recorder) names() []string {
	r.mu.Lock()
	defer r.mu.Unlock()
	return append([]string{}, r.events...)
}

func newTestHub() *Hub {
	h := NewHub()
	var tick int64
	h.now = func() time.Time { tick++; return time.UnixMilli(tick) }
	return h
}

func mustJoin(t *testing.T, h *Hub, rec *recorder, user, artifact, name string) *Conn {
	t.Helper()
	conn := NewConn(user, rec.sender())
	if _, err := h.Join(conn, artifact, name, FileMeta{Size: 1, Mtime: 1}); err != nil {
		t.Fatalf("join %s: %v", user, err)
	}
	return conn
}

func TestJoinAnnouncesToOthersAndReturnsExistingPeers(t *testing.T) {
	h := newTestHub()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "workbook-1"

	alice := NewConn("u1", aRec.sender())
	got, err := h.Join(alice, artifact, "Ann", FileMeta{})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Peers) != 0 {
		t.Fatalf("first peer should see nobody, got %d", len(got.Peers))
	}
	if got.Room == "" || got.Room == artifact || strings.Contains(got.Room, "/") {
		t.Fatalf("room must be an opaque key, got %q", got.Room)
	}
	if got.Self.ClientID != alice.ClientID() || got.Self.UserID != "u1" || got.Self.Color == "" {
		t.Fatalf("self identity must come from the conn, got %+v", got.Self)
	}

	bob := NewConn("u2", bRec.sender())
	got, err = h.Join(bob, artifact, "Bob", FileMeta{})
	if err != nil {
		t.Fatal(err)
	}
	if len(got.Peers) != 1 || got.Peers[0].ClientID != alice.ClientID() || got.Peers[0].Color == "" {
		t.Fatalf("second peer should see alice with a colour, got %+v", got.Peers)
	}
	if ev := aRec.names(); len(ev) != 1 || ev[0] != EventPeerJoined {
		t.Fatalf("alice should have been told bob joined, got %v", ev)
	}
	if ev := bRec.names(); len(ev) != 0 {
		t.Fatalf("bob must not be told about its own join, got %v", ev)
	}

	first := h.Peers(artifact)[1].JoinedAt
	if _, err := h.Join(bob, artifact, "Bob", FileMeta{}); err != nil {
		t.Fatal(err)
	}
	if ev := aRec.names(); len(ev) != 1 {
		t.Fatalf("re-join must not re-announce, got %v", ev)
	}
	if h.Peers(artifact)[1].JoinedAt != first {
		t.Fatal("re-join changed join time")
	}
}

func TestJoinRejectsClientForgedIdentity(t *testing.T) {
	h := newTestHub()
	conn := NewConn("real-user", (&recorder{}).sender())
	got, err := h.Join(conn, "artifact-1", "Ada", FileMeta{})
	if err != nil {
		t.Fatal(err)
	}
	if got.Self.UserID != "real-user" {
		t.Fatalf("user_id must come from NewConn, got %q", got.Self.UserID)
	}
	if got.Self.ClientID != conn.ClientID() || got.Self.ClientID == "forged" {
		t.Fatalf("client_id must be server-assigned, got %q", got.Self.ClientID)
	}
	if conn.ClientID() == "real-user" {
		t.Fatal("client_id must not equal user_id")
	}
}

func TestPresenceFansOutExceptSelfAndRejectsOversize(t *testing.T) {
	h := newTestHub()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "sheet-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	_ = mustJoin(t, h, bRec, "u2", artifact, "B")
	bRec.mu.Lock()
	bRec.events = nil
	bRec.mu.Unlock()

	sel := json.RawMessage(`{"sheet":"s1","ranges":[[0,0,2,2]]}`)
	if err := h.Presence(alice, artifact, sel); err != nil {
		t.Fatalf("presence from a member must succeed: %v", err)
	}
	if ev := bRec.names(); len(ev) != 1 || ev[0] != EventPresence {
		t.Fatalf("bob should receive alice's presence, got %v", ev)
	}
	if bRec.last["client_id"] != alice.ClientID() {
		t.Fatalf("presence payload wrong: %v", bRec.last)
	}
	if ev := aRec.names(); len(ev) != 1 { // only bob's join
		t.Fatalf("alice must not receive its own presence, got %v", ev)
	}

	late := NewConn("u3", (&recorder{}).sender())
	got, err := h.Join(late, artifact, "C", FileMeta{})
	if err != nil {
		t.Fatal(err)
	}
	if string(got.Peers[0].Selection) != string(sel) && string(got.Peers[1].Selection) != string(sel) {
		t.Fatalf("late joiner should see alice's selection, got %+v", got.Peers)
	}

	stranger := NewConn("zzz", (&recorder{}).sender())
	if err := h.Presence(stranger, artifact, sel); err != ErrNotInRoom {
		t.Fatalf("presence from a non-member must be rejected, got %v", err)
	}
	big := json.RawMessage(make([]byte, MaxSelectionBytes+1))
	if err := h.Presence(alice, artifact, big); err != ErrSelectionTooLarge {
		t.Fatalf("oversized selection must be rejected, got %v", err)
	}
}

func TestLeaveAllOnDisconnectNotifiesEveryRoomAndGarbageCollects(t *testing.T) {
	h := newTestHub()
	aRec, bRec, cRec := &recorder{}, &recorder{}, &recorder{}
	alice := NewConn("u1", aRec.sender())
	bob := NewConn("u2", bRec.sender())
	carol := NewConn("u3", cRec.sender())
	if _, err := h.Join(alice, "r1", "A", FileMeta{}); err != nil {
		t.Fatal(err)
	}
	if _, err := h.Join(alice, "r2", "A", FileMeta{}); err != nil {
		t.Fatal(err)
	}
	if _, err := h.Join(bob, "r1", "B", FileMeta{}); err != nil {
		t.Fatal(err)
	}
	if _, err := h.Join(carol, "r2", "C", FileMeta{}); err != nil {
		t.Fatal(err)
	}

	h.Disconnect(alice)
	for _, rec := range []*recorder{bRec, cRec} {
		ev := rec.names()
		if len(ev) == 0 || ev[len(ev)-1] != EventPeerLeft || rec.last["client_id"] != alice.ClientID() {
			t.Fatalf("peer should learn alice left, got %v / %v", ev, rec.last)
		}
	}
	if len(h.Peers("r1")) != 1 || len(h.Peers("r2")) != 1 {
		t.Fatal("alice should be gone from both rooms")
	}
	h.Disconnect(bob)
	h.Disconnect(carol)
	h.mu.Lock()
	defer h.mu.Unlock()
	if len(h.rooms) != 0 || len(h.byClient) != 0 {
		t.Fatalf("empty rooms must be collected: rooms=%d byClient=%d", len(h.rooms), len(h.byClient))
	}
}

func TestFileChangedSkipsOriginAndIsRoomScoped(t *testing.T) {
	h := newTestHub()
	aRec, bRec, otherRec := &recorder{}, &recorder{}, &recorder{}
	artifact := "artifact-x"
	other := "artifact-y"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	_ = mustJoin(t, h, bRec, "u2", artifact, "B")
	_ = mustJoin(t, h, otherRec, "u9", other, "O")
	bRec.mu.Lock()
	bRec.events = nil
	bRec.mu.Unlock()

	h.FileChanged(context.Background(), artifact, FileChange{Author: "user", Origin: alice.ClientID(), Version: "v1", Mtime: 5})
	if ev := aRec.names(); len(ev) != 1 { // only bob's join
		t.Fatalf("origin must not be notified of its own save, got %v", ev)
	}
	ev := bRec.names()
	if len(ev) != 1 || ev[0] != EventFileChanged || bRec.last["version"] != "v1" || bRec.last["action"] != "saved" {
		t.Fatalf("follower should get file.changed, got %v / %v", ev, bRec.last)
	}
	if bRec.last["path"] != artifact {
		t.Fatalf("file.changed path must be the opaque artifact id, got %v", bRec.last["path"])
	}
	if ev := otherRec.names(); len(ev) != 0 {
		t.Fatalf("other room must not be notified, got %v", ev)
	}
	h.FileChanged(context.Background(), artifact, FileChange{Author: "agent", Action: "delivered"})
	if len(aRec.names()) != 2 || len(bRec.names()) != 2 {
		t.Fatal("agent write should reach every peer")
	}
}

func TestRoomKeyIsOpaqueAndStable(t *testing.T) {
	if RoomKey("a") != RoomKey("a") || RoomKey("a") == RoomKey("b") {
		t.Fatal("room key must be a stable function of the artifact id")
	}
	looksLikePath := "/etc/passwd"
	key := RoomKey(looksLikePath)
	if key == looksLikePath || strings.Contains(key, "/") || strings.Contains(key, "..") {
		t.Fatalf("room key must not be a filesystem path, got %q", key)
	}
}

func TestColorIsStablePerUser(t *testing.T) {
	if ColorFor("u1") != ColorFor("u1") {
		t.Fatal("colour must be deterministic")
	}
}

func TestJoinBoundsRoomMembers(t *testing.T) {
	h := newTestHub()
	var conns []*Conn
	for i := 0; i < MaxRoomMembers; i++ {
		c := NewConn("u", (&recorder{}).sender())
		if _, err := h.Join(c, "full-room", "n", FileMeta{}); err != nil {
			t.Fatalf("member %d: %v", i, err)
		}
		conns = append(conns, c)
	}
	extra := NewConn("u-extra", (&recorder{}).sender())
	if _, err := h.Join(extra, "full-room", "n", FileMeta{}); err != ErrRoomFull {
		t.Fatalf("want ErrRoomFull, got %v", err)
	}
	if _, err := h.Join(conns[0], "full-room", "n", FileMeta{}); err != nil {
		t.Fatalf("re-join of an existing member must succeed: %v", err)
	}
}

func TestJoinRejectsInvalidArtifact(t *testing.T) {
	h := newTestHub()
	c := NewConn("u", (&recorder{}).sender())
	if _, err := h.Join(c, "", "n", FileMeta{}); err != ErrInvalidArtifact {
		t.Fatalf("empty id: %v", err)
	}
	if _, err := h.Join(c, string(make([]byte, MaxArtifactIDLen+1)), "n", FileMeta{}); err != ErrInvalidArtifact {
		t.Fatalf("oversize id: %v", err)
	}
	if _, err := h.Join(c, "ok\x00hidden", "n", FileMeta{}); err != ErrInvalidArtifact {
		t.Fatalf("nul id: %v", err)
	}
}

func TestSanitizeNameStripsMarkup(t *testing.T) {
	if got := sanitizeName("<script>x</script>", "u1"); strings.Contains(got, "<") || strings.Contains(got, ">") {
		t.Fatalf("name leaked markup: %q", got)
	}
	if got := sanitizeName("", "abcdefghij"); got != "user-abcdefgh" {
		t.Fatalf("default name: %q", got)
	}
}

func TestConcurrentJoinPresenceLeaveNoRace(t *testing.T) {
	h := NewHub()
	artifact := "race-1"
	var wg sync.WaitGroup
	for i := 0; i < 8; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			c := NewConn("u", (&recorder{}).sender())
			sel := json.RawMessage(`{"a":1}`)
			ops := json.RawMessage(`[{"id":"m"}]`)
			for j := 0; j < 25; j++ {
				_, _ = h.Join(c, artifact, "n", FileMeta{})
				_ = h.Presence(c, artifact, sel)
				_, _ = h.Submit(context.Background(), c, artifact, ops, 0)
				h.Leave(c, artifact)
			}
			h.Disconnect(c)
		}()
	}
	wg.Wait()
}
