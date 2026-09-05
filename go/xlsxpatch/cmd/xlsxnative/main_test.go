package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch/xlsxhttp"
)

const happyTreeSHA = "c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513"

func testdata(t *testing.T, parts ...string) string {
	t.Helper()
	path := filepath.Join(append([]string{"..", "..", "testdata"}, parts...)...)
	if _, err := os.Stat(path); err != nil {
		t.Fatal(err)
	}
	return path
}

func readHappyTree(t *testing.T) []byte {
	t.Helper()
	data, err := os.ReadFile(testdata(t, "excel-authored", "happy-tree.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	sum := sha256.Sum256(data)
	if hex.EncodeToString(sum[:]) != happyTreeSHA {
		t.Fatalf("happy-tree.xlsx sha256=%s, want %s", hex.EncodeToString(sum[:]), happyTreeSHA)
	}
	return data
}

func TestExtractHappyTreeFixture(t *testing.T) {
	original := readHappyTree(t)
	encoded, err := xlsxhttp.ExtractNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := xlsxpatch.DecodeNativeWorkbookV2(encoded)
	if err != nil {
		t.Fatal(err)
	}
	if workbook.Version != 2 || workbook.Source.PackageSHA256 != "sha256:"+happyTreeSHA {
		t.Fatalf("unexpected extract identity: version=%d sha=%s", workbook.Version, workbook.Source.PackageSHA256)
	}
	if len(workbook.PassthroughParts) == 0 || len(workbook.Unsupported) == 0 {
		t.Fatalf("passthrough/unsupported inventory missing: passthrough=%d unsupported=%d", len(workbook.PassthroughParts), len(workbook.Unsupported))
	}
	want, err := os.ReadFile(testdata(t, "native-xlsx-v2", "valid", "excel-authored-happy-tree.json"))
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(encoded, bytes.TrimSpace(want)) {
		t.Fatal("extract JSON is not the committed native v2 happy-tree fixture")
	}

	stdout := &bytes.Buffer{}
	stderr := &bytes.Buffer{}
	if code := run([]string{"extract", testdata(t, "excel-authored", "happy-tree.xlsx")}, bytes.NewReader(nil), stdout, stderr); code != 0 {
		t.Fatalf("extract CLI failed: code=%d stderr=%s", code, stderr)
	}
	if !bytes.Equal(bytes.TrimSpace(stdout.Bytes()), encoded) {
		t.Fatal("extract CLI output disagreed with helper extract")
	}
}

func TestApplyOneCellMutationOnHappyTree(t *testing.T) {
	original := readHappyTree(t)
	extracted, err := xlsxhttp.ExtractNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	before, err := xlsxpatch.DecodeNativeWorkbookV2(extracted)
	if err != nil {
		t.Fatal(err)
	}
	payload, err := json.Marshal(xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: "edit-a1",
			SheetID:     before.Sheets[0].ID,
			Kind:        xlsxpatch.CellSetValue,
			Cell:        xlsxpatch.CellRef{Row: 0, Column: 0},
			Value:       "native-preview",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	produced, err := xlsxhttp.ApplyNativeMutation(original, payload, before.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Equal(original, produced) {
		t.Fatal("apply returned original bytes")
	}
	afterJSON, err := xlsxhttp.ExtractNativeJSON(produced)
	if err != nil {
		t.Fatal(err)
	}
	after, err := xlsxpatch.DecodeNativeWorkbookV2(afterJSON)
	if err != nil {
		t.Fatal(err)
	}
	cell := findCell(t, after, before.Sheets[0].ID, 0, 0)
	if cell.Value == nil || cell.Value.Text == nil || *cell.Value.Text != "native-preview" {
		t.Fatalf("mutated A1 mismatch: %#v", cell.Value)
	}

	dir := t.TempDir()
	payloadPath := filepath.Join(dir, "payload.json")
	if err := os.WriteFile(payloadPath, payload, 0o600); err != nil {
		t.Fatal(err)
	}
	stdout := &bytes.Buffer{}
	stderr := &bytes.Buffer{}
	code := run([]string{"apply", testdata(t, "excel-authored", "happy-tree.xlsx"), payloadPath}, bytes.NewReader(nil), stdout, stderr)
	if code != 0 {
		t.Fatalf("apply CLI failed: code=%d stderr=%s", code, stderr)
	}
	if !bytes.Equal(stdout.Bytes(), produced) {
		t.Fatal("apply CLI output disagreed with helper apply")
	}
}

func TestServeRejectsNonLoopbackAddr(t *testing.T) {
	for _, addr := range []string{"0.0.0.0:18765", "192.168.1.9:18765", ":18765"} {
		if err := requireLoopbackAddr(addr); err == nil {
			t.Fatalf("non-loopback addr %q was accepted", addr)
		}
	}
	for _, addr := range []string{"127.0.0.1:18765", "localhost:18765", "[::1]:18765"} {
		if err := requireLoopbackAddr(addr); err != nil {
			t.Fatalf("loopback addr %q was refused: %v", addr, err)
		}
	}
}

func TestUnknownCommandAndHelp(t *testing.T) {
	stderr := &bytes.Buffer{}
	if code := run(nil, bytes.NewReader(nil), io.Discard, stderr); code != 2 || !strings.Contains(stderr.String(), "xlsxnative wrap") {
		t.Fatalf("missing usage: code=%d stderr=%s", code, stderr)
	}
	stderr.Reset()
	if code := run([]string{"nope"}, bytes.NewReader(nil), io.Discard, stderr); code != 2 {
		t.Fatalf("unknown command code=%d", code)
	}
}

func findCell(t *testing.T, workbook *xlsxpatch.NativeWorkbookV2, sheetID string, row, column int) xlsxpatch.NativeWorkbookCellV2 {
	t.Helper()
	for _, sheet := range workbook.Sheets {
		if sheet.ID != sheetID {
			continue
		}
		for _, cell := range sheet.Cells {
			if cell.Row == row && cell.Column == column {
				return cell
			}
		}
		t.Fatalf("cell r=%d c=%d missing on sheet %s", row, column, sheetID)
	}
	t.Fatalf("sheet %s missing", sheetID)
	return xlsxpatch.NativeWorkbookCellV2{}
}
