package collab

import (
	"context"
	"encoding/json"
	"sync"
)

// MemoryStore is an in-process OpStore. The hub Forgets a room's log when
// the last peer leaves so a long-lived process does not retain every artifact
// ever opened. Hosts that need a different retention policy should implement
// OpStore (durable stores typically no-op Forget).
type MemoryStore struct {
	mu    sync.Mutex
	rooms map[string]*memLog
}

type memLog struct {
	entries  []OpEntry
	seq      int64
	savedSeq int64
}

// NewMemoryStore returns an empty in-memory op log.
func NewMemoryStore() *MemoryStore {
	return &MemoryStore{rooms: map[string]*memLog{}}
}

func (s *MemoryStore) log(roomKey string) *memLog {
	l := s.rooms[roomKey]
	if l == nil {
		l = &memLog{}
		s.rooms[roomKey] = l
	}
	return l
}

// Append implements OpStore.
func (s *MemoryStore) Append(ctx context.Context, roomKey, artifactID, clientID string, ops json.RawMessage, baseSeq int64) (int64, error) {
	if err := ctx.Err(); err != nil {
		return 0, err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	l := s.log(roomKey)
	if baseSeq != l.seq {
		return l.seq, ErrStaleBase
	}
	l.seq++
	l.entries = append(l.entries, OpEntry{
		Room:     roomKey,
		Seq:      l.seq,
		ClientID: clientID,
		Ops:      cloneRaw(ops),
	})
	if n := len(l.entries) - MaxOpsPerRoom; n > 0 {
		l.entries = append([]OpEntry(nil), l.entries[n:]...)
	}
	return l.seq, nil
}

// Since implements OpStore.
func (s *MemoryStore) Since(ctx context.Context, roomKey string, since int64) ([]OpEntry, int64, bool, error) {
	if err := ctx.Err(); err != nil {
		return nil, 0, false, err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	l := s.rooms[roomKey]
	head := int64(0)
	saved := int64(0)
	var entries []OpEntry
	if l != nil {
		head = l.seq
		saved = l.savedSeq
		entries = l.entries
	}
	if since < 0 {
		since = 0
	}
	if since > head {
		return nil, head, false, nil
	}
	if since == head {
		return nil, head, true, nil
	}
	// Cleared history (reset or trim) and ring eviction both mean the client
	// cannot replay onto its snapshot.
	if since < saved {
		return nil, head, false, nil
	}
	if len(entries) > 0 && entries[0].Seq > since+1 {
		return nil, head, false, nil
	}
	var out []OpEntry
	for _, e := range entries {
		if e.Seq > since {
			out = append(out, OpEntry{
				Room:     e.Room,
				Seq:      e.Seq,
				ClientID: e.ClientID,
				Ops:      cloneRaw(e.Ops),
			})
		}
	}
	return out, head, true, nil
}

// LogState implements OpStore.
func (s *MemoryStore) LogState(ctx context.Context, roomKey string) (int64, int64, error) {
	if err := ctx.Err(); err != nil {
		return 0, 0, err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	l := s.rooms[roomKey]
	if l == nil {
		return 0, 0, nil
	}
	return l.seq, l.savedSeq, nil
}

// MarkSaved implements OpStore.
func (s *MemoryStore) MarkSaved(ctx context.Context, roomKey string, seq int64) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	l := s.rooms[roomKey]
	if l == nil || seq <= l.savedSeq {
		return nil
	}
	if seq > l.seq {
		seq = l.seq
	}
	l.savedSeq = seq
	i := 0
	for i < len(l.entries) && l.entries[i].Seq <= seq {
		i++
	}
	if i > 0 {
		l.entries = append([]OpEntry(nil), l.entries[i:]...)
	}
	return nil
}

// Reset implements OpStore.
func (s *MemoryStore) Reset(ctx context.Context, roomKey string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	l := s.log(roomKey)
	l.seq++
	l.savedSeq = l.seq
	l.entries = nil
	return nil
}

// Forget implements OpStore.
func (s *MemoryStore) Forget(ctx context.Context, roomKey string) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	delete(s.rooms, roomKey)
	return nil
}
