// Package fsstore is a local filesystem adapter for opaque Office artifact IDs.
// IDs are minted tokens; they are never Injecting workspace paths and are never
// used as path components under the store root.
package fsstore

import (
	"crypto/rand"
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"os"
	"path/filepath"
	"regexp"
	"strings"
	"sync"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const (
	idPrefix   = "art_"
	idHexBytes = 16
	idPattern  = `^art_[0-9a-f]{32}$`
	objectSalt = "injoffice-fs-artifact\x00"
)

var maxArtifactBytes = max(
	xlsxpatch.NativeXLSXMaxPackageBytes,
	docxpatch.NativeDOCXMaxPackageBytes,
	pptxpatch.NativePPTXMaxPackageBytes,
)

var artifactIDRe = regexp.MustCompile(idPattern)

// DirStore persists Office package bytes under a directory. Object filenames
// are a hash of the opaque ID so a caller-supplied string cannot escape the root.
type DirStore struct {
	root string
	mu   sync.Mutex
}

// Open creates (if needed) and returns a directory-backed artifact store.
func Open(root string) (*DirStore, error) {
	if root == "" {
		return nil, fmt.Errorf("fsstore: directory is required")
	}
	abs, err := filepath.Abs(root)
	if err != nil {
		return nil, fmt.Errorf("fsstore: resolve directory: %w", err)
	}
	if err := os.MkdirAll(abs, 0o700); err != nil {
		return nil, fmt.Errorf("fsstore: create directory: %w", err)
	}
	return &DirStore{root: abs}, nil
}

// Create mints an opaque ID and writes the package bytes.
func (s *DirStore) Create(data []byte) (string, error) {
	if err := checkPackage(data); err != nil {
		return "", err
	}
	var raw [idHexBytes]byte
	if _, err := rand.Read(raw[:]); err != nil {
		return "", wrapIO("mint id", err)
	}
	id := idPrefix + hex.EncodeToString(raw[:])
	s.mu.Lock()
	defer s.mu.Unlock()
	path, err := s.objectPath(id)
	if err != nil {
		return "", err
	}
	if _, err := os.Stat(path); err == nil {
		return "", xlsxhttp.ErrArtifactExists
	}
	if err := writeFileAtomic(path, data); err != nil {
		return "", err
	}
	return id, nil
}

// Get returns stored package bytes for a minted opaque ID.
func (s *DirStore) Get(id string) ([]byte, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	path, err := s.objectPath(id)
	if err != nil {
		return nil, err
	}
	data, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			return nil, xlsxhttp.ErrArtifactNotFound
		}
		return nil, wrapIO("read", err)
	}
	if err := checkPackage(data); err != nil {
		return nil, err
	}
	return data, nil
}

// Put replaces bytes for an existing opaque ID only when the current object
// digest still matches expected. Unknown IDs are not created.
func (s *DirStore) Put(id string, data, expected []byte) error {
	if err := checkPackage(data); err != nil {
		return err
	}
	if len(expected) == 0 {
		return xlsxhttp.ErrArtifactStale
	}
	s.mu.Lock()
	defer s.mu.Unlock()
	path, err := s.objectPath(id)
	if err != nil {
		return err
	}
	current, err := os.ReadFile(path)
	if err != nil {
		if os.IsNotExist(err) {
			return xlsxhttp.ErrArtifactNotFound
		}
		return wrapIO("read", err)
	}
	if sha256.Sum256(current) != sha256.Sum256(expected) {
		return xlsxhttp.ErrArtifactStale
	}
	return writeFileAtomic(path, data)
}

func (s *DirStore) objectPath(id string) (string, error) {
	if !artifactIDRe.MatchString(id) {
		return "", xlsxhttp.ErrInvalidArtifactID
	}
	sum := sha256.Sum256([]byte(objectSalt + id))
	name := hex.EncodeToString(sum[:])
	path := filepath.Join(s.root, name)
	rel, err := filepath.Rel(s.root, path)
	if err != nil || rel == ".." || strings.HasPrefix(rel, ".."+string(os.PathSeparator)) {
		return "", xlsxhttp.ErrInvalidArtifactID
	}
	if filepath.Base(path) != name {
		return "", xlsxhttp.ErrInvalidArtifactID
	}
	return path, nil
}

func checkPackage(data []byte) error {
	if len(data) == 0 {
		return fmt.Errorf("fsstore: package size must be 1..%d bytes", maxArtifactBytes)
	}
	if len(data) > maxArtifactBytes {
		return xlsxhttp.ErrArtifactTooLarge
	}
	return nil
}

func wrapIO(op string, err error) error {
	return fmt.Errorf("fsstore: %s: %w: %v", op, xlsxhttp.ErrArtifactIO, err)
}

func writeFileAtomic(path string, data []byte) error {
	dir := filepath.Dir(path)
	tmp, err := os.CreateTemp(dir, ".artifact-*")
	if err != nil {
		return wrapIO("temp", err)
	}
	tmpName := tmp.Name()
	defer func() { _ = os.Remove(tmpName) }()
	if _, err := tmp.Write(data); err != nil {
		_ = tmp.Close()
		return wrapIO("write", err)
	}
	if err := tmp.Chmod(0o600); err != nil {
		_ = tmp.Close()
		return wrapIO("chmod", err)
	}
	if err := tmp.Close(); err != nil {
		return wrapIO("close", err)
	}
	if err := os.Rename(tmpName, path); err != nil {
		return wrapIO("rename", err)
	}
	return nil
}
