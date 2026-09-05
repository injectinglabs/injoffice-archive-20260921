package collab

import (
	"context"
	"encoding/json"
	"errors"
	"strings"
	"sync"
	"testing"
	"time"
)

func TestSubmitOrdersAndFansOut(t *testing.T) {
	h := newTestHub()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "workbook-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	bob := mustJoin(t, h, bRec, "u2", artifact, "B")
	bRec.mu.Lock()
	bRec.events = nil
	bRec.mu.Unlock()

	seq, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m1"}]`), 0)
	if err != nil || seq != 1 {
		t.Fatalf("first submit should be seq 1, got %d %v", seq, err)
	}
	seq, err = h.Submit(context.Background(), bob, artifact, json.RawMessage(`[{"id":"m2"}]`), 1)
	if err != nil || seq != 2 {
		t.Fatalf("second submit should be seq 2, got %d %v", seq, err)
	}
	if ev := bRec.names(); len(ev) != 1 || ev[0] != EventOp || bRec.last["seq"].(float64) != 1 || bRec.last["client_id"] != alice.ClientID() {
		t.Fatalf("bob should get alice's op only, got %v / %v", ev, bRec.last)
	}
	if ev := aRec.names(); ev[len(ev)-1] != EventOp || aRec.last["seq"].(float64) != 2 {
		t.Fatalf("alice should get bob's op, got %v / %v", ev, aRec.last)
	}

	stranger := NewConn("zzz", (&recorder{}).sender())
	if _, err := h.Submit(context.Background(), stranger, artifact, json.RawMessage(`[{"id":"x"}]`), 2); err != ErrNotInRoom {
		t.Fatalf("non-members can't submit: %v", err)
	}
	if _, err := h.Submit(context.Background(), alice, artifact, make(json.RawMessage, MaxOpBytes+1), 2); err != ErrOpTooLarge {
		t.Fatal("oversized op must be rejected")
	}
	if st := h.Log(context.Background(), artifact); st.Seq != 2 || st.SavedSeq != 0 {
		t.Fatalf("log state wrong: %+v", st)
	}
}

func TestSubmitStaleBaseAndCatchUp(t *testing.T) {
	h := newTestHub()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "workbook-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	bob := mustJoin(t, h, bRec, "u2", artifact, "B")

	seq, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m1"}]`), 0)
	if err != nil || seq != 1 {
		t.Fatalf("base==head must be accepted: %d %v", seq, err)
	}
	head, err := h.Submit(context.Background(), bob, artifact, json.RawMessage(`[{"id":"m2"}]`), 0)
	if !errors.Is(err, ErrStaleBase) || head != 1 {
		t.Fatalf("stale base must be rejected with the head, got %d %v", head, err)
	}
	if !strings.Contains(err.Error(), "STALE_BASE") {
		t.Fatalf("error text must contain STALE_BASE, got %q", err)
	}

	gap, err := h.Since(context.Background(), bob, artifact, 0)
	if err != nil || gap.Reset || gap.Head != 1 || len(gap.Ops) != 1 || gap.Ops[0].Seq != 1 {
		t.Fatalf("catch-up should give seq 1, got %+v %v", gap, err)
	}
	seq, err = h.Submit(context.Background(), bob, artifact, json.RawMessage(`[{"id":"m2"}]`), gap.Head)
	if err != nil || seq != 2 {
		t.Fatalf("caught-up base must be accepted: %d %v", seq, err)
	}
}

func TestSinceMarkSavedAndReset(t *testing.T) {
	h := newTestHub()
	aRec := &recorder{}
	artifact := "workbook-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	for i := 0; i < 5; i++ {
		base := int64(i)
		if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), base); err != nil {
			t.Fatalf("submit %d: %v", i, err)
		}
	}
	gap, err := h.Since(context.Background(), alice, artifact, 2)
	if err != nil || gap.Reset || gap.Head != 5 || len(gap.Ops) != 3 || gap.Ops[0].Seq != 3 {
		t.Fatalf("since 2 should give 3..5, got %+v %v", gap, err)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 5); err != nil || e.Reset || len(e.Ops) != 0 {
		t.Fatalf("since head is empty, got %+v %v", e, err)
	}

	h.MarkSaved(context.Background(), artifact, 3)
	if st := h.Log(context.Background(), artifact); st.SavedSeq != 3 {
		t.Fatalf("saved seq should be 3, got %+v", st)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 1); err != nil || !e.Reset {
		t.Fatalf("a gap below saved_seq must answer reset, got %+v %v", e, err)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 3); err != nil || e.Reset || len(e.Ops) != 2 {
		t.Fatalf("since saved_seq must still work, got %+v %v", e, err)
	}

	h.MarkSaved(context.Background(), artifact, 2)
	h.MarkSaved(context.Background(), artifact, 99)
	if st := h.Log(context.Background(), artifact); st.SavedSeq != 5 {
		t.Fatalf("saved seq should clamp to head 5, got %+v", st)
	}

	if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), 5); err != nil {
		t.Fatal(err)
	}
	h.ResetLog(context.Background(), artifact)
	st := h.Log(context.Background(), artifact)
	if st.SavedSeq != 7 || st.Seq != 7 {
		t.Fatalf("reset should bump head so the old snapshot is invalid, got %+v", st)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 7); err != nil || e.Reset || len(e.Ops) != 0 {
		t.Fatalf("after reset, since new head is clean: %+v %v", e, err)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 6); err != nil || !e.Reset {
		t.Fatalf("after reset, old head must reset: %+v %v", e, err)
	}

	h.MarkSaved(context.Background(), "nope", 1)
	h.ResetLog(context.Background(), "nope")
}

func TestResetInvalidatesSubmitAtOldHead(t *testing.T) {
	h := newTestHub()
	alice := mustJoin(t, h, &recorder{}, "u1", "workbook-1", "A")
	seq, err := h.Submit(context.Background(), alice, "workbook-1", json.RawMessage(`[{"id":"m"}]`), 0)
	if err != nil || seq != 1 {
		t.Fatalf("submit: %d %v", seq, err)
	}
	h.ResetLog(context.Background(), "workbook-1")
	head, err := h.Submit(context.Background(), alice, "workbook-1", json.RawMessage(`[{"id":"old"}]`), 1)
	if !errors.Is(err, ErrStaleBase) || head != 2 {
		t.Fatalf("submit at old head after reset: %d %v", head, err)
	}
	gap, err := h.Since(context.Background(), alice, "workbook-1", 1)
	if err != nil || !gap.Reset || gap.Head != 2 {
		t.Fatalf("since old head after reset: %+v %v", gap, err)
	}
	seq, err = h.Submit(context.Background(), alice, "workbook-1", json.RawMessage(`[{"id":"new"}]`), 2)
	if err != nil || seq != 3 {
		t.Fatalf("submit against post-reset head: %d %v", seq, err)
	}
}

func TestSinceResetWhenAheadOfHeadOrMissingLog(t *testing.T) {
	h := newTestHub()
	alice := mustJoin(t, h, &recorder{}, "u1", "fresh", "A")
	if e, err := h.Since(context.Background(), alice, "fresh", 0); err != nil || e.Reset || e.Head != 0 || len(e.Ops) != 0 {
		t.Fatalf("since==head on a fresh room is current: %+v %v", e, err)
	}
	if e, err := h.Since(context.Background(), alice, "fresh", 5); err != nil || !e.Reset || e.Head != 0 {
		t.Fatalf("since > head on a fresh room must reset: %+v %v", e, err)
	}

	if _, err := h.Submit(context.Background(), alice, "fresh", json.RawMessage(`[{"id":"m"}]`), 0); err != nil {
		t.Fatal(err)
	}
	h.Disconnect(alice)
	bob := mustJoin(t, h, &recorder{}, "u2", "fresh", "B")
	if st := h.Log(context.Background(), "fresh"); st.Seq != 0 || st.SavedSeq != 0 {
		t.Fatalf("empty room should forget the in-memory log, got %+v", st)
	}
	if e, err := h.Since(context.Background(), bob, "fresh", 1); err != nil || !e.Reset || e.Head != 0 {
		t.Fatalf("rejoin after wipe, since old seq must reset: %+v %v", e, err)
	}

	store := NewMemoryStore()
	h2 := NewHubWithStore(store)
	carol := mustJoin(t, h2, &recorder{}, "u3", "wiped", "C")
	if _, err := h2.Submit(context.Background(), carol, "wiped", json.RawMessage(`[{"id":"m"}]`), 0); err != nil {
		t.Fatal(err)
	}
	h2.Forget(context.Background(), "wiped")
	if e, err := h2.Since(context.Background(), carol, "wiped", 10); err != nil || !e.Reset || e.Head != 0 {
		t.Fatalf("since past a forgotten log must reset: %+v %v", e, err)
	}
}

func TestEmptyRoomForgetsInMemoryLog(t *testing.T) {
	h := newTestHub()
	artifact := "workbook-1"
	alice := mustJoin(t, h, &recorder{}, "u1", artifact, "A")
	if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), 0); err != nil {
		t.Fatal(err)
	}
	h.Disconnect(alice)
	if len(h.Peers(artifact)) != 0 {
		t.Fatal("expected empty room")
	}
	bob := NewConn("u2", (&recorder{}).sender())
	got, err := h.Join(bob, artifact, "B", FileMeta{})
	if err != nil {
		t.Fatal(err)
	}
	if got.Log.Seq != 0 {
		t.Fatalf("in-memory log should be forgotten when the room empties, got %+v", got.Log)
	}
}

type gateForget struct {
	*MemoryStore
	started chan struct{}
	allow   chan struct{}
}

func (g *gateForget) Forget(ctx context.Context, roomKey string) error {
	select {
	case <-g.started:
	default:
		close(g.started)
	}
	<-g.allow
	return g.MemoryStore.Forget(ctx, roomKey)
}

func TestForgetDoesNotDropConcurrentJoinSubmit(t *testing.T) {
	g := &gateForget{
		MemoryStore: NewMemoryStore(),
		started:     make(chan struct{}),
		allow:       make(chan struct{}),
	}
	h := NewHubWithStore(g)
	artifact := "workbook-1"
	alice := mustJoin(t, h, &recorder{}, "u1", artifact, "A")
	if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"old"}]`), 0); err != nil {
		t.Fatal(err)
	}

	left := make(chan struct{})
	go func() {
		h.Leave(alice, artifact)
		close(left)
	}()
	<-g.started

	bob := NewConn("u2", (&recorder{}).sender())
	joined := make(chan int64, 1)
	go func() {
		if _, err := h.Join(bob, artifact, "B", FileMeta{}); err != nil {
			t.Errorf("join: %v", err)
			joined <- 0
			return
		}
		seq, err := h.Submit(context.Background(), bob, artifact, json.RawMessage(`[{"id":"new"}]`), 0)
		if err != nil {
			t.Errorf("submit: %v", err)
			joined <- 0
			return
		}
		joined <- seq
	}()

	select {
	case <-joined:
		t.Fatal("Join/Submit must wait until last-leave Forget finishes")
	case <-left:
		t.Fatal("Leave must not return until Forget is allowed")
	case <-time.After(50 * time.Millisecond):
	}

	close(g.allow)
	<-left
	seq := <-joined
	if seq == 0 {
		t.Fatal("expected a seq from the new session")
	}
	gap, err := h.Since(context.Background(), bob, artifact, 0)
	if err != nil || gap.Reset || len(gap.Ops) == 0 || gap.Ops[len(gap.Ops)-1].Seq != seq {
		t.Fatalf("new session submit was dropped: seq=%d gap=%+v err=%v", seq, gap, err)
	}
}

func TestConcurrentLastLeaveJoinSubmitKeepsOps(t *testing.T) {
	h := NewHub()
	artifact := "workbook-1"
	ctx := context.Background()
	ops := json.RawMessage(`[{"id":"m"}]`)
	var wg sync.WaitGroup
	errCh := make(chan error, 32)
	for i := 0; i < 16; i++ {
		wg.Add(1)
		go func() {
			defer wg.Done()
			c := NewConn("u", func(string, any) {})
			for j := 0; j < 20; j++ {
				if _, err := h.Join(c, artifact, "n", FileMeta{}); err != nil {
					continue
				}
				st := h.Log(ctx, artifact)
				seq, err := h.Submit(ctx, c, artifact, ops, st.Seq)
				if err != nil {
					h.Leave(c, artifact)
					continue
				}
				gap, err := h.Since(ctx, c, artifact, seq-1)
				if err != nil {
					errCh <- err
					return
				}
				found := false
				for _, e := range gap.Ops {
					if e.Seq == seq {
						found = true
						break
					}
				}
				if gap.Reset || !found {
					errCh <- errors.New("submit vanished from the log while still in the room")
					return
				}
				h.Leave(c, artifact)
			}
		}()
	}
	wg.Wait()
	close(errCh)
	for err := range errCh {
		t.Fatal(err)
	}
}

func TestRingBound(t *testing.T) {
	h := newTestHub()
	artifact := "big-workbook"
	alice := mustJoin(t, h, &recorder{}, "u1", artifact, "A")
	for i := 0; i < MaxOpsPerRoom+10; i++ {
		if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), int64(i)); err != nil {
			t.Fatalf("submit %d: %v", i, err)
		}
	}
	if e, err := h.Since(context.Background(), alice, artifact, 5); err != nil || !e.Reset {
		t.Fatalf("evicted gap must answer reset, got %+v %v", e, err)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 10); err != nil || e.Reset || len(e.Ops) != MaxOpsPerRoom {
		t.Fatalf("ring should hold the last %d, got %+v %v", MaxOpsPerRoom, e, err)
	}
}

func TestFileChangedCarriesSavedSeq(t *testing.T) {
	h := newTestHub()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "workbook-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	_ = mustJoin(t, h, bRec, "u2", artifact, "B")
	if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), 0); err != nil {
		t.Fatal(err)
	}
	if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), 1); err != nil {
		t.Fatal(err)
	}
	h.NotifyFileChanged(context.Background(), artifact, FileChange{Author: "user", Origin: alice.ClientID(), SavedSeq: 2})
	if bRec.last["saved_seq"].(float64) != 2 {
		t.Fatalf("file.changed should carry saved_seq 2, got %v", bRec.last)
	}
}

func TestNotifyFileChangedReset(t *testing.T) {
	h := newTestHub()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "workbook-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	_ = mustJoin(t, h, bRec, "u2", artifact, "B")
	if _, err := h.Submit(context.Background(), alice, artifact, json.RawMessage(`[{"id":"m"}]`), 0); err != nil {
		t.Fatal(err)
	}
	h.NotifyFileChanged(context.Background(), artifact, FileChange{Author: "agent", Action: "delivered", Reset: true})
	st := h.Log(context.Background(), artifact)
	if st.Seq != 2 || st.SavedSeq != 2 {
		t.Fatalf("reset notify should bump head and pin saved_seq, got %+v", st)
	}
	if bRec.last["reset"] != true {
		t.Fatalf("file.changed should carry reset, got %v", bRec.last)
	}
	if e, err := h.Since(context.Background(), alice, artifact, 0); err != nil || !e.Reset {
		t.Fatalf("history after reset must not replay, got %+v %v", e, err)
	}
}

func TestSubmitIgnoresClientIDInsideOps(t *testing.T) {
	h := newTestHub()
	alice := mustJoin(t, h, &recorder{}, "u1", "workbook-1", "A")
	ops := json.RawMessage(`[{"id":"edit","params":{"client_id":"forged","user_id":"forged"}}]`)
	if _, err := h.Submit(context.Background(), alice, "workbook-1", ops, 0); err != nil {
		t.Fatal(err)
	}
	gap, err := h.Since(context.Background(), alice, "workbook-1", 0)
	if err != nil || len(gap.Ops) != 1 {
		t.Fatalf("since: %+v %v", gap, err)
	}
	if gap.Ops[0].ClientID != alice.ClientID() || gap.Ops[0].ClientID == "forged" {
		t.Fatalf("entry client_id must be the conn, got %q", gap.Ops[0].ClientID)
	}
}

func TestSubmitBoundsBatchAndDepth(t *testing.T) {
	h := newTestHub()
	alice := mustJoin(t, h, &recorder{}, "u1", "workbook-1", "A")
	if _, err := h.Submit(context.Background(), alice, "workbook-1", json.RawMessage(`{"id":"not-array"}`), 0); err != ErrInvalidOps {
		t.Fatalf("non-array: %v", err)
	}
	if _, err := h.Submit(context.Background(), alice, "workbook-1", json.RawMessage(`[]`), 0); err != ErrInvalidOps {
		t.Fatalf("empty array: %v", err)
	}
	tooMany := make([]byte, 0, 64)
	tooMany = append(tooMany, '[')
	for i := 0; i < MaxOpsPerBatch+1; i++ {
		if i > 0 {
			tooMany = append(tooMany, ',')
		}
		tooMany = append(tooMany, `{"id":"m"}`...)
	}
	tooMany = append(tooMany, ']')
	if _, err := h.Submit(context.Background(), alice, "workbook-1", tooMany, 0); err != ErrBatchTooLarge {
		t.Fatalf("batch: %v", err)
	}
	deep := []byte(`[{"id":"m","params":`)
	for i := 0; i < MaxJSONDepth; i++ {
		deep = append(deep, `{"k":`...)
	}
	deep = append(deep, '1')
	for i := 0; i < MaxJSONDepth; i++ {
		deep = append(deep, '}')
	}
	deep = append(deep, `}]`...)
	if _, err := h.Submit(context.Background(), alice, "workbook-1", deep, 0); err != ErrJSONTooDeep {
		t.Fatalf("depth: %v", err)
	}
}

func TestConcurrentSubmitAtSameBaseOneWins(t *testing.T) {
	h := newTestHub()
	alice := mustJoin(t, h, &recorder{}, "u1", "workbook-1", "A")
	bob := mustJoin(t, h, &recorder{}, "u2", "workbook-1", "B")
	var wg sync.WaitGroup
	var mu sync.Mutex
	var accepted []int64
	var stale int
	submit := func(c *Conn) {
		defer wg.Done()
		seq, err := h.Submit(context.Background(), c, "workbook-1", json.RawMessage(`[{"id":"m"}]`), 0)
		mu.Lock()
		defer mu.Unlock()
		if err == nil {
			accepted = append(accepted, seq)
			return
		}
		if errors.Is(err, ErrStaleBase) {
			stale++
			return
		}
		t.Errorf("unexpected err: %v", err)
	}
	wg.Add(2)
	go submit(alice)
	go submit(bob)
	wg.Wait()
	if len(accepted) != 1 || accepted[0] != 1 || stale != 1 {
		t.Fatalf("exactly one submit must win: accepted=%v stale=%d", accepted, stale)
	}
}

func TestHubDelegatesToOpStore(t *testing.T) {
	fake := &fakeOpStore{}
	h := NewHubWithStore(fake)
	var tick int64
	h.now = func() time.Time { tick++; return time.UnixMilli(tick) }
	ctx := context.Background()
	aRec, bRec := &recorder{}, &recorder{}
	artifact := "workbook-1"
	alice := mustJoin(t, h, aRec, "u1", artifact, "A")
	_ = mustJoin(t, h, bRec, "u2", artifact, "B")
	bRec.mu.Lock()
	bRec.events = nil
	bRec.mu.Unlock()

	seq, err := h.Submit(ctx, alice, artifact, json.RawMessage(`[{"id":"m"}]`), 0)
	if err != nil || seq != 1 {
		t.Fatalf("submit via store: %d %v", seq, err)
	}
	if ev := bRec.names(); len(ev) != 1 || ev[0] != EventOp || bRec.last["seq"].(float64) != 1 {
		t.Fatalf("peer fan-out still works with the store: %v %v", ev, bRec.last)
	}
	if head, err := h.Submit(ctx, alice, artifact, json.RawMessage(`[{"id":"m"}]`), 0); err != ErrStaleBase || head != 1 {
		t.Fatalf("stale base via store: %d %v", head, err)
	}
	if st := h.Log(ctx, artifact); st.Seq != 1 || st.SavedSeq != 0 {
		t.Fatalf("log state via store: %+v", st)
	}
	h.MarkSaved(ctx, artifact, 1)
	if e, err := h.Since(ctx, alice, artifact, 0); err != nil || !e.Reset {
		t.Fatalf("gap below saved answers reset via store: %+v %v", e, err)
	}
}

type fakeOpStore struct {
	head, saved int64
	rows        []OpEntry
}

func (f *fakeOpStore) Append(_ context.Context, roomKey, _, clientID string, ops json.RawMessage, baseSeq int64) (int64, error) {
	if baseSeq != f.head {
		return f.head, ErrStaleBase
	}
	f.head++
	f.rows = append(f.rows, OpEntry{Room: roomKey, Seq: f.head, ClientID: clientID, Ops: cloneRaw(ops)})
	return f.head, nil
}

func (f *fakeOpStore) Since(_ context.Context, _ string, since int64) ([]OpEntry, int64, bool, error) {
	if since < 0 {
		since = 0
	}
	if since > f.head {
		return nil, f.head, false, nil
	}
	if since == f.head {
		return nil, f.head, true, nil
	}
	if since < f.saved {
		return nil, f.head, false, nil
	}
	var out []OpEntry
	for _, e := range f.rows {
		if e.Seq > since {
			out = append(out, e)
		}
	}
	return out, f.head, true, nil
}

func (f *fakeOpStore) LogState(context.Context, string) (int64, int64, error) {
	return f.head, f.saved, nil
}

func (f *fakeOpStore) MarkSaved(_ context.Context, _ string, seq int64) error {
	if seq > f.head {
		seq = f.head
	}
	if seq > f.saved {
		f.saved = seq
		kept := f.rows[:0]
		for _, e := range f.rows {
			if e.Seq > seq {
				kept = append(kept, e)
			}
		}
		f.rows = kept
	}
	return nil
}

func (f *fakeOpStore) Reset(context.Context, string) error {
	f.head++
	f.saved = f.head
	f.rows = nil
	return nil
}

func (f *fakeOpStore) Forget(context.Context, string) error {
	return nil
}
