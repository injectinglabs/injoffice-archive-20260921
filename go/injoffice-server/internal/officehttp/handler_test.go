package officehttp

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"mime/multipart"
	"net/http"
	"net/http/httptest"
	"os"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const (
	numberingDOCXSHA = "c77d07d2f0489e1d4c02f479fac7318f42485334f4accae56835009f048757a8"
	commonPPTXSHA    = "09041bde998c7f18db24a7f9c9aae250fce53784265f928fa8c0fcc97b990ff4"
	happyTreeSHA     = "c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513"
)

func testdata(t *testing.T, parts ...string) string {
	t.Helper()
	path := filepath.Join(append([]string{"..", "..", ".."}, parts...)...)
	if _, err := os.Stat(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func readPinned(t *testing.T, wantSHA string, parts ...string) []byte {
	t.Helper()
	data, err := os.ReadFile(testdata(t, parts...))
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(data)
	if hex.EncodeToString(sum[:]) != wantSHA {
		t.Fatalf("%s sha256=%s, want %s", filepath.Join(parts...), hex.EncodeToString(sum[:]), wantSHA)
	}
	return data
}

func TestDOCXExtractMutateRoundTripHTTP(t *testing.T) {
	original := readPinned(t, numberingDOCXSHA, "docxpatch", "testdata", "native-numbering-v1.docx")
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)

	extractRes, err := http.Post(server.URL+DOCXExtractPath, DOCXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	defer extractRes.Body.Close()
	if extractRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(extractRes.Body)
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, body)
	}
	if media := extractRes.Header.Get("Content-Type"); media != docxpatch.NativeDOCXV1MediaType {
		t.Fatalf("extract content-type %q", media)
	}
	artifactID := extractRes.Header.Get(xlsxhttp.HeaderArtifactID)
	if artifactID == "" || strings.ContainsAny(artifactID, `/\`) || strings.Contains(artifactID, "..") || strings.Contains(artifactID, "docxpatch") {
		t.Fatalf("artifact id is not opaque: %q", artifactID)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	before, err := docxpatch.DecodeNativeDocumentV1(encoded)
	if err != nil {
		t.Fatal(err)
	}
	paragraph := before.Body.Blocks[0].Paragraph

	reload := postJSON(t, server.URL+DOCXExtractPath, map[string]string{"artifact_id": artifactID})
	defer reload.Body.Close()
	if reload.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(reload.Body)
		t.Fatalf("extract-by-id HTTP %d: %s", reload.StatusCode, body)
	}
	if reload.Header.Get(xlsxhttp.HeaderArtifactID) != artifactID {
		t.Fatalf("reload artifact id %q, want %q", reload.Header.Get(xlsxhttp.HeaderArtifactID), artifactID)
	}

	payload, err := json.Marshal(map[string]any{"mutations": []docxpatch.NativeDOCXTextMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256, Text: "from-server",
	}}})
	if err != nil {
		t.Fatal(err)
	}
	mutateRes := postMutation(t, server.URL+DOCXMutationsPath, payload, before.Source.PackageSHA256, artifactID)
	defer mutateRes.Body.Close()
	if mutateRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(mutateRes.Body)
		t.Fatalf("mutations HTTP %d: %s", mutateRes.StatusCode, body)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderArtifactID) != artifactID {
		t.Fatalf("mutated artifact id %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderArtifactID), artifactID)
	}
	produced, err := io.ReadAll(mutateRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(stored, produced) {
		t.Fatal("store did not persist mutated package")
	}
	afterJSON, err := extractDOCXJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := docxpatch.DecodeNativeDocumentV1(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	got := *after.Body.Blocks[0].Paragraph.Runs[0].Text
	if got != "from-server" {
		t.Fatalf("HTTP mutation did not land: %q", got)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderRevision) != after.Revision {
		t.Fatalf("mutation revision header %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderRevision), after.Revision)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderPackageSHA) != after.Source.PackageSHA256 {
		t.Fatalf("mutation package sha header %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderPackageSHA), after.Source.PackageSHA256)
	}
}

func TestPPTXExtractMutateRoundTripHTTP(t *testing.T) {
	original := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)

	extractRes, err := http.Post(server.URL+PPTXExtractPath, PPTXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	defer extractRes.Body.Close()
	if extractRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(extractRes.Body)
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, body)
	}
	if media := extractRes.Header.Get("Content-Type"); media != pptxpatch.NativePPTXV1MediaType {
		t.Fatalf("extract content-type %q", media)
	}
	artifactID := extractRes.Header.Get(xlsxhttp.HeaderArtifactID)
	if artifactID == "" || strings.ContainsAny(artifactID, `/\`) || strings.Contains(artifactID, "..") || strings.Contains(artifactID, "officecompat") {
		t.Fatalf("artifact id is not opaque: %q", artifactID)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	before, err := pptxpatch.DecodeNativePPTXJSON(encoded)
	if err != nil {
		t.Fatal(err)
	}
	target := before.Slides[0].Elements[0]
	paragraphs := pptxMutationParagraphs("from-http")
	payload, err := json.Marshal(pptxpatch.NativePPTXMutationRequest{
		ExpectedSourceRevision: *before.SourceRevision,
		Operations: []pptxpatch.NativePPTXMutation{{
			OperationID: "http-edit", Kind: pptxpatch.NativePPTXReplaceText, ElementID: target.ID,
			ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}

	reload := postJSON(t, server.URL+PPTXExtractPath, map[string]string{"artifact_id": artifactID})
	defer reload.Body.Close()
	if reload.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(reload.Body)
		t.Fatalf("extract-by-id HTTP %d: %s", reload.StatusCode, body)
	}
	if reload.Header.Get(xlsxhttp.HeaderArtifactID) != artifactID {
		t.Fatalf("reload artifact id %q, want %q", reload.Header.Get(xlsxhttp.HeaderArtifactID), artifactID)
	}

	outer := "sha256:" + commonPPTXSHA
	mutateRes := postMutation(t, server.URL+PPTXMutationsPath, payload, outer, artifactID)
	defer mutateRes.Body.Close()
	if mutateRes.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(mutateRes.Body)
		t.Fatalf("mutations HTTP %d: %s", mutateRes.StatusCode, body)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderArtifactID) != artifactID {
		t.Fatalf("mutated artifact id %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderArtifactID), artifactID)
	}
	produced, err := io.ReadAll(mutateRes.Body)
	if err != nil {
		t.Fatal(err)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(stored, produced) {
		t.Fatal("store did not persist mutated package")
	}
	afterJSON, err := extractPPTXJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := pptxpatch.DecodeNativePPTXJSON(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	got := *(*after.Slides[0].Elements[0].Paragraphs)[0].Runs[0].Text
	if got != "from-http" {
		t.Fatalf("HTTP mutation did not land: %q", got)
	}
	sum := sha256.Sum256(stored)
	wantRev := "rev-" + hex.EncodeToString(sum[:])
	wantSHA := "sha256:" + hex.EncodeToString(sum[:])
	if mutateRes.Header.Get(xlsxhttp.HeaderRevision) != wantRev {
		t.Fatalf("mutation revision header %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderRevision), wantRev)
	}
	if mutateRes.Header.Get(xlsxhttp.HeaderPackageSHA) != wantSHA {
		t.Fatalf("mutation package sha header %q, want %q", mutateRes.Header.Get(xlsxhttp.HeaderPackageSHA), wantSHA)
	}
}

func TestExtractDoesNotPersistInvalidPackages(t *testing.T) {
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	for _, path := range []string{DOCXExtractPath, PPTXExtractPath} {
		res, err := http.Post(server.URL+path, "application/octet-stream", bytes.NewReader([]byte("not-a-zip")))
		if err != nil {
			t.Fatal(err)
		}
		body, _ := io.ReadAll(res.Body)
		res.Body.Close()
		if res.StatusCode != http.StatusBadRequest {
			t.Fatalf("%s status %d, want 400: %s", path, res.StatusCode, body)
		}
		if res.Header.Get(xlsxhttp.HeaderArtifactID) != "" {
			t.Fatalf("%s failed extract minted artifact id %q", path, res.Header.Get(xlsxhttp.HeaderArtifactID))
		}
	}
	if store.len() != 0 {
		t.Fatal("failed extract persisted an artifact")
	}
}

func TestExtractRejectsWorkspacePaths(t *testing.T) {
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	for _, path := range []string{DOCXExtractPath, PPTXExtractPath, xlsxhttp.ExtractPath} {
		for _, id := range []string{
			"../docxpatch/testdata/native-numbering-v1.docx",
			"/workspace/injecting/happy-tree.xlsx",
			filepath.Join("officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx"),
		} {
			res := postJSON(t, server.URL+path, map[string]string{"artifact_id": id})
			body, _ := io.ReadAll(res.Body)
			res.Body.Close()
			if res.StatusCode != http.StatusBadRequest {
				t.Fatalf("%s path %q status %d, want 400: %s", path, id, res.StatusCode, body)
			}
			if !strings.Contains(string(body), "invalid artifact id") {
				t.Fatalf("%s path %q error %s", path, id, body)
			}
		}
	}
}

func TestXLSXRoutesStillServed(t *testing.T) {
	original := readPinned(t, happyTreeSHA, "xlsxpatch", "testdata", "excel-authored", "happy-tree.xlsx")
	server := httptest.NewServer(NewHandler(nil))
	t.Cleanup(server.Close)
	res, err := http.Post(server.URL+xlsxhttp.ExtractPath, xlsxhttp.XLSXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	defer res.Body.Close()
	if res.StatusCode != http.StatusOK {
		body, _ := io.ReadAll(res.Body)
		t.Fatalf("xlsx extract HTTP %d: %s", res.StatusCode, body)
	}
	if media := res.Header.Get("Content-Type"); media != xlsxpatch.NativeXLSXV2MediaType {
		t.Fatalf("xlsx extract content-type %q", media)
	}
	if res.Header.Get(xlsxhttp.HeaderArtifactID) != "" {
		t.Fatal("stateless helper must not mint artifact ids")
	}
}

func TestDOCXConcurrentMutationsCAS(t *testing.T) {
	original := readPinned(t, numberingDOCXSHA, "docxpatch", "testdata", "native-numbering-v1.docx")
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	extractRes, err := http.Post(server.URL+DOCXExtractPath, DOCXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	extractRes.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	if extractRes.StatusCode != http.StatusOK {
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, encoded)
	}
	artifactID := extractRes.Header.Get(xlsxhttp.HeaderArtifactID)
	before, err := docxpatch.DecodeNativeDocumentV1(encoded)
	if err != nil {
		t.Fatal(err)
	}
	payloadA := docxTextPayload(t, before.Body.Blocks[0].Paragraph, "from-a")
	payloadB := docxTextPayload(t, before.Body.Blocks[1].Paragraph, "from-b")
	rev := before.Source.PackageSHA256
	results := make(chan int, 2)
	go func() {
		res := postMutation(t, server.URL+DOCXMutationsPath, payloadA, rev, artifactID)
		results <- res.StatusCode
		res.Body.Close()
	}()
	go func() {
		res := postMutation(t, server.URL+DOCXMutationsPath, payloadB, rev, artifactID)
		results <- res.StatusCode
		res.Body.Close()
	}()
	first, second := <-results, <-results
	ok, stale := 0, 0
	for _, code := range []int{first, second} {
		switch code {
		case http.StatusOK:
			ok++
		case http.StatusBadRequest:
			stale++
		default:
			t.Fatalf("unexpected mutation status %d", code)
		}
	}
	if ok != 1 || stale != 1 {
		t.Fatalf("CAS race statuses ok=%d stale=%d (got %d and %d)", ok, stale, first, second)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	afterJSON, err := extractDOCXJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := docxpatch.DecodeNativeDocumentV1(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	a := paragraphText(after.Body.Blocks[0].Paragraph) == "from-a"
	b := paragraphText(after.Body.Blocks[1].Paragraph) == "from-b"
	if a == b {
		t.Fatalf("store should keep exactly one racing mutation: a=%v b=%v", a, b)
	}
}

func TestPPTXConcurrentMutationsCAS(t *testing.T) {
	original := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	extractRes, err := http.Post(server.URL+PPTXExtractPath, PPTXContentType, bytes.NewReader(original))
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	extractRes.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	if extractRes.StatusCode != http.StatusOK {
		t.Fatalf("extract HTTP %d: %s", extractRes.StatusCode, encoded)
	}
	artifactID := extractRes.Header.Get(xlsxhttp.HeaderArtifactID)
	before, err := pptxpatch.DecodeNativePPTXJSON(encoded)
	if err != nil {
		t.Fatal(err)
	}
	target := before.Slides[0].Elements[0]
	payloadA := pptxTextPayload(t, target, "from-a")
	payloadB := pptxTextPayload(t, target, "from-b")
	rev := "sha256:" + commonPPTXSHA
	results := make(chan int, 2)
	go func() {
		res := postMutation(t, server.URL+PPTXMutationsPath, payloadA, rev, artifactID)
		results <- res.StatusCode
		res.Body.Close()
	}()
	go func() {
		res := postMutation(t, server.URL+PPTXMutationsPath, payloadB, rev, artifactID)
		results <- res.StatusCode
		res.Body.Close()
	}()
	first, second := <-results, <-results
	ok, stale := 0, 0
	for _, code := range []int{first, second} {
		switch code {
		case http.StatusOK:
			ok++
		case http.StatusBadRequest:
			stale++
		default:
			t.Fatalf("unexpected mutation status %d", code)
		}
	}
	if ok != 1 || stale != 1 {
		t.Fatalf("CAS race statuses ok=%d stale=%d (got %d and %d)", ok, stale, first, second)
	}
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	afterJSON, err := extractPPTXJSON(stored)
	if err != nil {
		t.Fatal(err)
	}
	after, err := pptxpatch.DecodeNativePPTXJSON(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	got := firstRunText(after.Slides[0].Elements[0])
	if got != "from-a" && got != "from-b" {
		t.Fatalf("stored PPTX mutation text %q, want from-a or from-b", got)
	}
}

func TestMutationsHTTPRejectsCASMismatches(t *testing.T) {
	docxOriginal := readPinned(t, numberingDOCXSHA, "docxpatch", "testdata", "native-numbering-v1.docx")
	doc, err := docxpatch.ExtractNativeDocumentV1(docxOriginal)
	if err != nil {
		t.Fatal(err)
	}
	docxPayload := docxTextPayload(t, doc.Body.Blocks[0].Paragraph, "must-not-commit")
	pptxOriginal := readPinned(t, commonPPTXSHA, "officecompat", "corpus", "generated", "packages", "pptx-transitional-common.pptx")
	store := newMemStore()
	server := httptest.NewServer(NewHandler(store))
	t.Cleanup(server.Close)
	extractRes, err := http.Post(server.URL+PPTXExtractPath, PPTXContentType, bytes.NewReader(pptxOriginal))
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := io.ReadAll(extractRes.Body)
	extractRes.Body.Close()
	if err != nil {
		t.Fatal(err)
	}
	if extractRes.StatusCode != http.StatusOK {
		t.Fatalf("pptx extract HTTP %d: %s", extractRes.StatusCode, encoded)
	}
	artifactID := extractRes.Header.Get(xlsxhttp.HeaderArtifactID)
	before, err := pptxpatch.DecodeNativePPTXJSON(encoded)
	if err != nil {
		t.Fatal(err)
	}
	pptxPayload := pptxTextPayload(t, before.Slides[0].Elements[0], "must-not-commit")
	stale := "sha256:" + strings.Repeat("0", 64)

	docxRes := postMutationBytes(t, server.URL+DOCXMutationsPath, docxOriginal, docxPayload, stale, "")
	assertRejectedPackage(t, docxRes, DOCXContentType, "STALE_REVISION")

	pptxRes := postMutation(t, server.URL+PPTXMutationsPath, pptxPayload, stale, artifactID)
	assertRejectedPackage(t, pptxRes, PPTXContentType, "stale outer")
	stored, err := store.Get(artifactID)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(stored, pptxOriginal) {
		t.Fatal("stale PPTX mutation persisted a candidate package")
	}
}

func assertRejectedPackage(t *testing.T, res *http.Response, packageType, want string) {
	t.Helper()
	body, _ := io.ReadAll(res.Body)
	res.Body.Close()
	if res.StatusCode != http.StatusBadRequest {
		t.Fatalf("status %d, want 400: %s", res.StatusCode, body)
	}
	if res.Header.Get("Content-Type") == packageType || bytes.HasPrefix(body, []byte("PK")) {
		t.Fatalf("rejected mutation returned package bytes (%d)", len(body))
	}
	if !strings.Contains(string(body), want) && !strings.Contains(string(body), "does not match exact package revision") {
		t.Fatalf("error %s", body)
	}
}

func paragraphText(paragraph *docxpatch.NativeParagraphV1) string {
	if paragraph == nil || len(paragraph.Runs) == 0 || paragraph.Runs[0].Text == nil {
		return ""
	}
	return *paragraph.Runs[0].Text
}

func firstRunText(element pptxpatch.NativeElement) string {
	if element.Paragraphs == nil || len(*element.Paragraphs) == 0 || len((*element.Paragraphs)[0].Runs) == 0 || (*element.Paragraphs)[0].Runs[0].Text == nil {
		return ""
	}
	return *(*element.Paragraphs)[0].Runs[0].Text
}

func docxTextPayload(t *testing.T, paragraph *docxpatch.NativeParagraphV1, text string) []byte {
	t.Helper()
	payload, err := json.Marshal(map[string]any{"mutations": []docxpatch.NativeDOCXTextMutationV1{{
		TargetKind: "paragraph", TargetID: paragraph.ID, ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256, Text: text,
	}}})
	if err != nil {
		t.Fatal(err)
	}
	return payload
}

func pptxTextPayload(t *testing.T, target pptxpatch.NativeElement, text string) []byte {
	t.Helper()
	paragraphs := pptxMutationParagraphs(text)
	payload, err := json.Marshal(pptxpatch.NativePPTXMutationRequest{
		ExpectedSourceRevision: "rev-" + commonPPTXSHA,
		Operations: []pptxpatch.NativePPTXMutation{{
			OperationID: "http-" + text, Kind: pptxpatch.NativePPTXReplaceText, ElementID: target.ID,
			ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs,
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	return payload
}

func pptxMutationParagraphs(text string) []pptxpatch.NativeParagraph {
	align := pptxpatch.NativeTextAlignRight
	level := int64(2)
	bullet := false
	bold, italic := false, true
	size := int64(2800)
	color := "ABCDEF"
	family := "Aptos Display"
	return []pptxpatch.NativeParagraph{{Align: &align, Level: &level, Bullet: &bullet, Runs: []pptxpatch.NativeTextRun{{
		Text: &text, Bold: &bold, Italic: &italic, FontSizeHundredthPt: &size, Color: &color, FontFamily: &family,
	}}}}
}

func postJSON(t *testing.T, url string, body any) *http.Response {
	t.Helper()
	encoded, err := json.Marshal(body)
	if err != nil {
		t.Fatal(err)
	}
	res, err := http.Post(url, "application/json", bytes.NewReader(encoded))
	if err != nil {
		t.Fatal(err)
	}
	return res
}

func postMutation(t *testing.T, url string, payload []byte, expectedRevision, artifactID string) *http.Response {
	t.Helper()
	return postMutationBytes(t, url, nil, payload, expectedRevision, artifactID)
}

func postMutationBytes(t *testing.T, url string, original, payload []byte, expectedRevision, artifactID string) *http.Response {
	t.Helper()
	body := &bytes.Buffer{}
	writer := multipart.NewWriter(body)
	if artifactID != "" {
		if err := writer.WriteField("artifact_id", artifactID); err != nil {
			t.Fatal(err)
		}
	} else {
		originalPart, err := writer.CreateFormFile("original", "original.bin")
		if err != nil {
			t.Fatal(err)
		}
		if _, err := originalPart.Write(original); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.WriteField("payload", string(payload)); err != nil {
		t.Fatal(err)
	}
	if expectedRevision != "" {
		if err := writer.WriteField("expected_revision", expectedRevision); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	res, err := http.Post(url, writer.FormDataContentType(), body)
	if err != nil {
		t.Fatal(err)
	}
	return res
}

type memStore struct {
	mu      sync.Mutex
	next    int
	objects map[string][]byte
}

func newMemStore() *memStore {
	return &memStore{objects: map[string][]byte{}}
}

func (s *memStore) len() int {
	s.mu.Lock()
	defer s.mu.Unlock()
	return len(s.objects)
}

func (s *memStore) Create(data []byte) (string, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	s.next++
	id := "art_" + strings.Repeat("0", 31) + "0123456789abcdef"[s.next%16:s.next%16+1]
	s.objects[id] = append([]byte(nil), data...)
	return id, nil
}

func (s *memStore) Get(id string) ([]byte, error) {
	s.mu.Lock()
	defer s.mu.Unlock()
	if !strings.HasPrefix(id, "art_") || len(id) != 36 {
		return nil, xlsxhttp.ErrInvalidArtifactID
	}
	for _, r := range id[4:] {
		if r < '0' || r > '9' && (r < 'a' || r > 'f') {
			return nil, xlsxhttp.ErrInvalidArtifactID
		}
	}
	data, ok := s.objects[id]
	if !ok {
		return nil, xlsxhttp.ErrArtifactNotFound
	}
	return append([]byte(nil), data...), nil
}

func (s *memStore) Put(id string, data, expected []byte) error {
	s.mu.Lock()
	defer s.mu.Unlock()
	current, ok := s.objects[id]
	if !ok {
		return xlsxhttp.ErrArtifactNotFound
	}
	if sha256.Sum256(current) != sha256.Sum256(expected) {
		return xlsxhttp.ErrArtifactStale
	}
	s.objects[id] = append([]byte(nil), data...)
	return nil
}
