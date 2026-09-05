package main

import (
	"crypto/sha256"
	"encoding/binary"
	"encoding/hex"
	"hash"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func extractNativeJSON(data []byte) ([]byte, error) {
	deck, err := pptxpatch.ExtractNativePPTX(data, pptxpatch.NativePPTXExtractOptions{
		TokenFactory: browserLocalTokenFactory{},
	})
	if err != nil {
		return nil, err
	}
	return pptxpatch.MarshalNativePPTXJSON(deck)
}

func applyNativePayload(original, payload []byte, expectedRevision string) ([]byte, error) {
	return pptxpatch.ApplyNativePPTXMutationPayload(original, expectedRevision, payload)
}

// browserLocalTokenFactory creates opaque, deterministic identifiers for
// preservation references inside one browser-local source package. The tokens
// deliberately use a distinct namespace and have no resolver: they are not
// server capabilities and cannot authorize reads from a remote artifact store.
//
// The factory is transactional only to satisfy group extraction's atomic
// publication contract. Issuing a token has no side effect, so commit and
// rollback are intentionally no-ops.
type browserLocalTokenFactory struct{}

func (browserLocalTokenFactory) IssueNativePassthroughToken(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
	return browserLocalPassthroughToken(request), nil
}

func (browserLocalTokenFactory) BeginNativePassthroughTokenTransaction() (pptxpatch.NativePassthroughTokenTransaction, error) {
	return browserLocalTokenTransaction{}, nil
}

type browserLocalTokenTransaction struct{}

func (browserLocalTokenTransaction) IssueNativePassthroughToken(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
	return browserLocalPassthroughToken(request), nil
}

func (browserLocalTokenTransaction) CommitNativePassthroughTokens() error { return nil }

func (browserLocalTokenTransaction) RollbackNativePassthroughTokens() {}

func browserLocalPassthroughToken(request pptxpatch.NativePassthroughTokenRequest) string {
	digest := sha256.New()
	writeTokenField(digest, []byte("injoffice.pptx.browser-local-passthrough.v1"))
	writeTokenField(digest, []byte(request.SourceRevision))
	writeTokenField(digest, []byte(request.OwnerPart))
	writeTokenField(digest, []byte(request.ObjectID))
	writeTokenField(digest, []byte(request.FingerprintSHA256))
	var byteLength [8]byte
	binary.BigEndian.PutUint64(byteLength[:], uint64(request.ByteLength))
	writeTokenField(digest, byteLength[:])
	writeTokenField(digest, []byte(request.Reason))
	payloadDigest := sha256.Sum256(request.Payload)
	writeTokenField(digest, payloadDigest[:])
	return "browser-local-v1-" + hex.EncodeToString(digest.Sum(nil))
}

func writeTokenField(digest hash.Hash, field []byte) {
	var size [8]byte
	binary.BigEndian.PutUint64(size[:], uint64(len(field)))
	_, _ = digest.Write(size[:])
	_, _ = digest.Write(field)
}
