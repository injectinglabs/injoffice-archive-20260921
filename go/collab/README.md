# collab

In-process collaboration hub and linear operation log for InjOffice artifacts. The module implements [`docs/COLLABORATION-PROTOCOL.md`](../../docs/COLLABORATION-PROTOCOL.md): join/leave/presence, `STALE_BASE`, catch-up, `saved_seq`, and reset. Artifact IDs are opaque strings — the hub never treats them as filesystem paths.

```bash
go get github.com/injectinglabs/injoffice/go/collab
```

Authentication, authorization, and persistence belong to the host. Construct a `Conn` from your authenticator (the hub assigns `client_id` and never reads `user_id` / `client_id` from a collab payload). `go/injoffice-server` hangs HTTP + SSE in front of this API for a local two-browser demo; this module does not serve the network.

## Five-minute in-process example

Two peers join one workbook, race a submit, catch up from the linear log, and disconnect:

```go
package main

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/injectinglabs/injoffice/go/collab"
)

func main() {
	hub := collab.NewHub()
	alice := collab.NewConn("alice", func(event string, payload any) {
		fmt.Printf("alice %s\n", event)
	})
	bob := collab.NewConn("bob", func(event string, payload any) {
		fmt.Printf("bob %s\n", event)
	})
	file := collab.FileMeta{Size: 4200, Mtime: 1}

	a, err := hub.Join(alice, "workbook-1", "Alice", file)
	if err != nil {
		panic(err)
	}
	if _, err := hub.Join(bob, "workbook-1", "Bob", file); err != nil {
		panic(err)
	}
	fmt.Printf("room=%s head=%d\n", a.Room, a.Log.Seq)

	seq, err := hub.Submit(context.Background(), alice, "workbook-1",
		json.RawMessage(`[{"id":"sheet.mutation.set-range-values","params":{"plain":"json"}}]`), 0)
	if err != nil {
		panic(err)
	}

	_, err = hub.Submit(context.Background(), bob, "workbook-1",
		json.RawMessage(`[{"id":"edit","params":{}}]`), 0)
	if err != collab.ErrStaleBase {
		panic(err)
	}

	gap, err := hub.Since(context.Background(), bob, "workbook-1", 0)
	if err != nil || gap.Reset {
		panic(err)
	}
	seq, err = hub.Submit(context.Background(), bob, "workbook-1",
		json.RawMessage(`[{"id":"edit","params":{}}]`), gap.Head)
	if err != nil {
		panic(err)
	}
	fmt.Printf("head=%d\n", seq)

	hub.NotifyFileChanged(context.Background(), "workbook-1", collab.FileChange{
		Author: "user", Origin: alice.ClientID(), SavedSeq: seq, Size: 4300,
	})
	hub.Disconnect(alice)
	hub.Disconnect(bob)
}
```

Run the tests (including this example) from the module root:

```bash
cd go/collab
go test ./...
```

## Contract

| RPC / event | Hub method |
|---|---|
| `collab.join` | `Join` |
| `collab.leave` / disconnect | `Leave` / `Disconnect` |
| `collab.presence` | `Presence` |
| `collab.op.submit` | `Submit` — rejects `base_seq != head` with `ErrStaleBase` (`STALE_BASE`) |
| `collab.op.since` | `Since` — `Reset` when history cannot be replayed |
| `collab.peer.joined` / `collab.peer.left` / `collab.presence` / `collab.op` / `collab.file.changed` | delivered through the `Sender` bound to each `Conn` |

`Sender` must not block. Drop or disconnect a slow consumer instead of growing an unbounded queue.

The default `OpStore` is in-memory. The hub calls `Forget` when the last peer leaves so a long-lived process does not retain every artifact ever opened; hosts may also call `Hub.Forget`. A durable sqlite/postgres store plugged in via `NewHubWithStore` should no-op `Forget` if unsaved ops must survive empty rooms.

`ResetLog` bumps the head so `Submit` at the old head returns `STALE_BASE` and `Since(oldHead)` returns `reset=true`. `Since` also returns `reset=true` when `since > head` or the log is gone.
