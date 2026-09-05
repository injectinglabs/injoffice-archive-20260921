package collab_test

import (
	"context"
	"encoding/json"
	"fmt"

	"github.com/injectinglabs/injoffice/go/collab"
)

func Example() {
	hub := collab.NewHub()
	alice := collab.NewConn("alice", func(string, any) {})
	bob := collab.NewConn("bob", func(string, any) {})
	file := collab.FileMeta{Size: 100, Mtime: 1}

	joined, err := hub.Join(alice, "workbook-1", "Alice", file)
	if err != nil {
		panic(err)
	}
	fmt.Println(joined.Log.Seq)

	if _, err := hub.Join(bob, "workbook-1", "Bob", file); err != nil {
		panic(err)
	}

	seq, err := hub.Submit(context.Background(), alice, "workbook-1", json.RawMessage(`[{"id":"edit"}]`), 0)
	if err != nil {
		panic(err)
	}
	fmt.Println(seq)

	_, err = hub.Submit(context.Background(), bob, "workbook-1", json.RawMessage(`[{"id":"edit-2"}]`), 0)
	fmt.Println(err)

	gap, err := hub.Since(context.Background(), bob, "workbook-1", 0)
	if err != nil {
		panic(err)
	}
	fmt.Println(gap.Head, gap.Reset, len(gap.Ops))

	seq, err = hub.Submit(context.Background(), bob, "workbook-1", json.RawMessage(`[{"id":"edit-2"}]`), gap.Head)
	if err != nil {
		panic(err)
	}
	fmt.Println(seq)

	hub.Disconnect(alice)
	hub.Disconnect(bob)

	// Output:
	// 0
	// 1
	// STALE_BASE
	// 1 false 1
	// 2
}
