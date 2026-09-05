package fsstore

import (
	"bytes"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

func TestDirStoreRoundTripAndOpaqueNames(t *testing.T) {
	store, err := Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	original := []byte("PK-fake-xlsx-bytes")
	id, err := store.Create(original)
	if err != nil {
		t.Fatal(err)
	}
	if !artifactIDRe.MatchString(id) {
		t.Fatalf("minted id %q is not opaque", id)
	}
	if strings.ContainsAny(id, `/\`) || strings.Contains(id, "..") || strings.Contains(id, "workspace") {
		t.Fatalf("minted id leaked a path: %q", id)
	}
	got, err := store.Get(id)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, original) {
		t.Fatal("Get returned different bytes")
	}
	entries, err := os.ReadDir(store.root)
	if err != nil {
		t.Fatal(err)
	}
	if len(entries) != 1 {
		t.Fatalf("store entries=%d, want 1", len(entries))
	}
	if entries[0].Name() == id || strings.Contains(entries[0].Name(), "workspace") {
		t.Fatalf("on-disk name %q must not be the artifact id or a workspace path", entries[0].Name())
	}

	updated := []byte("PK-mutated-xlsx-bytes")
	if err := store.Put(id, updated, original); err != nil {
		t.Fatal(err)
	}
	got, err = store.Get(id)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, updated) {
		t.Fatal("Put did not replace bytes")
	}
}

func TestDirStoreRejectsPathsAndUnknownIDs(t *testing.T) {
	store, err := Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	for _, id := range []string{
		"../xlsxpatch/testdata/excel-authored/happy-tree.xlsx",
		"/workspace/injecting/happy-tree.xlsx",
		filepath.Join("excel-authored", "happy-tree.xlsx"),
		"art_" + strings.Repeat("g", 32),
		"art_" + strings.Repeat("0", 31),
		"",
	} {
		if _, err := store.Get(id); err != xlsxhttp.ErrInvalidArtifactID {
			t.Fatalf("Get(%q) err=%v, want invalid artifact id", id, err)
		}
		if err := store.Put(id, []byte("PK"), []byte("PK")); err != xlsxhttp.ErrInvalidArtifactID {
			t.Fatalf("Put(%q) err=%v, want invalid artifact id", id, err)
		}
	}
	missing := "art_" + strings.Repeat("0", 32)
	if _, err := store.Get(missing); err != xlsxhttp.ErrArtifactNotFound {
		t.Fatalf("missing Get err=%v", err)
	}
	if err := store.Put(missing, []byte("PK-bytes"), []byte("PK-bytes")); err != xlsxhttp.ErrArtifactNotFound {
		t.Fatalf("missing Put err=%v", err)
	}
}

func TestDirStorePutCompareAndSwap(t *testing.T) {
	store, err := Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	original := []byte("PK-original-bytes")
	id, err := store.Create(original)
	if err != nil {
		t.Fatal(err)
	}
	first := []byte("PK-first-save")
	second := []byte("PK-second-save")
	if err := store.Put(id, first, original); err != nil {
		t.Fatal(err)
	}
	if err := store.Put(id, second, original); err != xlsxhttp.ErrArtifactStale {
		t.Fatalf("stale Put err=%v, want stale artifact revision", err)
	}
	got, err := store.Get(id)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(got, first) {
		t.Fatal("stale Put overwrote the first save")
	}
	if err := store.Put(id, second, first); err != nil {
		t.Fatal(err)
	}
}

func TestDirStorePutConcurrentCAS(t *testing.T) {
	store, err := Open(t.TempDir())
	if err != nil {
		t.Fatal(err)
	}
	original := []byte("PK-original-bytes")
	id, err := store.Create(original)
	if err != nil {
		t.Fatal(err)
	}
	errs := make(chan error, 2)
	go func() { errs <- store.Put(id, []byte("PK-first-save"), original) }()
	go func() { errs <- store.Put(id, []byte("PK-second-save"), original) }()
	first, second := <-errs, <-errs
	ok, stale := 0, 0
	for _, err := range []error{first, second} {
		switch err {
		case nil:
			ok++
		case xlsxhttp.ErrArtifactStale:
			stale++
		default:
			t.Fatalf("unexpected Put err %v", err)
		}
	}
	if ok != 1 || stale != 1 {
		t.Fatalf("concurrent Put ok=%d stale=%d (err %v / %v)", ok, stale, first, second)
	}
}
