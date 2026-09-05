//go:build !js

package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
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

func extractNativeJSON(t *testing.T, data []byte) []byte {
	t.Helper()
	workbook, err := xlsxpatch.ExtractNativeWorkbookV2(data)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := xlsxpatch.EncodeNativeWorkbookV2(workbook)
	if err != nil {
		t.Fatal(err)
	}
	return encoded
}

func wasmExecFromGOROOT() string {
	if env := strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM_EXEC")); env != "" {
		return env
	}
	if _, err := os.Stat(filepath.Join("dist", "wasm_exec.js")); err == nil {
		return filepath.Join("dist", "wasm_exec.js")
	}
	root := strings.TrimSpace(os.Getenv("GOROOT"))
	if root == "" {
		out, err := exec.Command("go", "env", "GOROOT").Output()
		if err != nil {
			return ""
		}
		root = strings.TrimSpace(string(out))
	}
	for _, rel := range []string{filepath.Join("misc", "wasm", "wasm_exec.js"), filepath.Join("lib", "wasm", "wasm_exec.js")} {
		candidate := filepath.Join(root, rel)
		if _, err := os.Stat(candidate); err == nil {
			return candidate
		}
	}
	return ""
}

func lookupExistingWasm() (wasm, wasmExec string, ok bool) {
	if env := strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM")); env != "" {
		wasm = env
	} else {
		wasm = filepath.Join("dist", "xlsxnative.wasm")
	}
	wasmExec = wasmExecFromGOROOT()
	if _, err := os.Stat(wasm); err != nil {
		return "", "", false
	}
	if wasmExec == "" {
		return "", "", false
	}
	if _, err := os.Stat(wasmExec); err != nil {
		return "", "", false
	}
	return wasm, wasmExec, true
}

var wasmBuild struct {
	once sync.Once
	wasm string
	exec string
	err  error
}

func ensureWasm(t *testing.T) (wasm, wasmExec string) {
	t.Helper()
	if wasm, wasmExec, ok := lookupExistingWasm(); ok {
		return wasm, wasmExec
	}
	if strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM_SKIP_BUILD")) != "" {
		if strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatal("xlsxnative.wasm is required but was not found; run ./scripts/build-xlsxnative-wasm.sh")
		}
		t.Skip("xlsxnative.wasm not built; run ./scripts/build-xlsxnative-wasm.sh")
	}
	wasmBuild.once.Do(func() {
		dir, err := os.MkdirTemp("", "xlsxnativewasm-")
		if err != nil {
			wasmBuild.err = err
			return
		}
		wasmBuild.wasm = filepath.Join(dir, "xlsxnative.wasm")
		cmd := exec.Command("go", "build", "-trimpath", "-ldflags=-s -w", "-o", wasmBuild.wasm, ".")
		cmd.Env = append(os.Environ(), "GOOS=js", "GOARCH=wasm")
		if out, buildErr := cmd.CombinedOutput(); buildErr != nil {
			wasmBuild.err = errWithOutput(buildErr, out)
			return
		}
		wasmBuild.exec = wasmExecFromGOROOT()
		if wasmBuild.exec == "" {
			wasmBuild.err = os.ErrNotExist
		}
	})
	if wasmBuild.err != nil {
		if strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatalf("GOOS=js WASM build is required: %v", wasmBuild.err)
		}
		t.Skipf("GOOS=js WASM build unavailable; run ./scripts/build-xlsxnative-wasm.sh: %v", wasmBuild.err)
	}
	if wasmBuild.wasm == "" || wasmBuild.exec == "" {
		if strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatal("xlsxnative.wasm is required but was not built; run ./scripts/build-xlsxnative-wasm.sh")
		}
		t.Skip("xlsxnative.wasm not built; run ./scripts/build-xlsxnative-wasm.sh")
	}
	return wasmBuild.wasm, wasmBuild.exec
}

func errWithOutput(err error, out []byte) error {
	if len(out) == 0 {
		return err
	}
	return &buildError{err: err, out: out}
}

type buildError struct {
	err error
	out []byte
}

func (e *buildError) Error() string {
	return e.err.Error() + "\n" + string(e.out)
}

func requireNodeHarness(t *testing.T) (wasm, wasmExec, script string) {
	t.Helper()
	if _, err := exec.LookPath("node"); err != nil {
		if strings.TrimSpace(os.Getenv("XLSXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatalf("node is required to instantiate GOOS=js WASM: %v", err)
		}
		t.Skip("node is required to instantiate GOOS=js WASM")
	}
	wasm, wasmExec = ensureWasm(t)
	script = "node_contract.mjs"
	if _, err := os.Stat(script); err != nil {
		t.Fatalf("missing %s", script)
	}
	return wasm, wasmExec, script
}

func TestWASMExtractMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := readHappyTree(t)
	want := extractNativeJSON(t, original)
	fixture, err := os.ReadFile(testdata(t, "native-xlsx-v2", "valid", "excel-authored-happy-tree.json"))
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(want, bytes.TrimSpace(fixture)) {
		t.Fatal("in-process extract JSON is not the committed native v2 happy-tree fixture")
	}

	cmd := exec.Command("node", script, "extract",
		"--wasm", wasm,
		"--wasm-exec", wasmExec,
		"--input", testdata(t, "excel-authored", "happy-tree.xlsx"),
	)
	got, err := cmd.Output()
	if err != nil {
		t.Fatalf("WASM extract failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(bytes.TrimSpace(got), want) {
		t.Fatalf("WASM extract JSON disagreed with in-process Go (%d vs %d bytes)", len(bytes.TrimSpace(got)), len(want))
	}
}

func TestWASMApplyMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := readHappyTree(t)
	beforeJSON := extractNativeJSON(t, original)
	before, err := xlsxpatch.DecodeNativeWorkbookV2(beforeJSON)
	if err != nil {
		t.Fatal(err)
	}
	payload, err := json.Marshal(xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: "wasm-edit-a1",
			SheetID:     before.Sheets[0].ID,
			Kind:        xlsxpatch.CellSetValue,
			Cell:        xlsxpatch.CellRef{Row: 0, Column: 0},
			Value:       "native-wasm",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	want, err := xlsxpatch.ApplyNativeWorkbookMutationPayloadV1(original, payload, before.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}

	dir := t.TempDir()
	payloadPath := filepath.Join(dir, "payload.json")
	if err := os.WriteFile(payloadPath, payload, 0o600); err != nil {
		t.Fatal(err)
	}
	cmd := exec.Command("node", script, "apply",
		"--wasm", wasm,
		"--wasm-exec", wasmExec,
		"--original", testdata(t, "excel-authored", "happy-tree.xlsx"),
		"--payload", payloadPath,
		"--expected-revision", before.Source.PackageSHA256,
	)
	got, err := cmd.Output()
	if err != nil {
		t.Fatalf("WASM apply failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(got, want.Package) {
		t.Fatalf("WASM apply bytes disagreed with in-process Go (%d vs %d bytes)", len(got), len(want.Package))
	}
}

func TestWASMRefusalsKeepInstanceAlive(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := readHappyTree(t)
	want := extractNativeJSON(t, original)
	before, err := xlsxpatch.DecodeNativeWorkbookV2(want)
	if err != nil {
		t.Fatal(err)
	}
	payload, err := json.Marshal(xlsxpatch.NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells: []xlsxpatch.CellMutation{{
			OperationID: "cas-edit",
			SheetID:     before.Sheets[0].ID,
			Kind:        xlsxpatch.CellSetValue,
			Cell:        xlsxpatch.CellRef{Row: 0, Column: 0},
			Value:       "must-not-commit",
		}},
	})
	if err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	payloadPath := filepath.Join(dir, "payload.json")
	if err := os.WriteFile(payloadPath, payload, 0o600); err != nil {
		t.Fatal(err)
	}
	stale := "sha256:" + strings.Repeat("0", 64)
	cmd := exec.Command("node", script, "survive",
		"--wasm", wasm,
		"--wasm-exec", wasmExec,
		"--input", testdata(t, "excel-authored", "happy-tree.xlsx"),
		"--payload", payloadPath,
		"--expected-revision", before.Source.PackageSHA256,
		"--rev-token", before.Revision,
		"--stale-revision", stale,
	)
	got, err := cmd.Output()
	if err != nil {
		t.Fatalf("WASM survive probe failed: %v\n%s", err, stderrFrom(err))
	}
	if bytes.HasPrefix(got, []byte("PK")) || bytes.Equal(bytes.TrimSpace(got), []byte("undefined")) {
		t.Fatalf("survive wrote XLSX or undefined instead of a JSON report (%d bytes)", len(got))
	}
	var report struct {
		Refusals []struct {
			Name  string `json:"name"`
			Error string `json:"error"`
			Fatal bool   `json:"fatal"`
		} `json:"refusals"`
		FatalProbe struct {
			Name  string `json:"name"`
			Error string `json:"error"`
			Fatal bool   `json:"fatal"`
		} `json:"fatalProbe"`
		LegacyProbe struct {
			Name  string `json:"name"`
			Error string `json:"error"`
			Fatal bool   `json:"fatal"`
		} `json:"legacyProbe"`
		Extract string `json:"extract"`
	}
	if err := json.Unmarshal(got, &report); err != nil {
		t.Fatalf("survive report is not JSON: %v\n%s", err, got)
	}
	wantNames := []string{"empty_extract", "empty_original", "empty_payload", "missing_revision", "rev_token", "stale_cas"}
	if len(report.Refusals) != len(wantNames) {
		t.Fatalf("refusals=%d, want %d: %+v", len(report.Refusals), len(wantNames), report.Refusals)
	}
	for i, name := range wantNames {
		item := report.Refusals[i]
		if item.Name != name {
			t.Fatalf("refusal %d name %q, want %q", i, item.Name, name)
		}
		if item.Error == "" || item.Error == "undefined" || strings.Contains(item.Error, "Go program has already exited") {
			t.Fatalf("%s did not return a catchable Go error: %q", name, item.Error)
		}
		if item.Fatal {
			t.Fatalf("ordinary refusal %s was marked fatal", name)
		}
	}
	if !strings.Contains(report.Refusals[3].Error, "expectedRevision is required") {
		t.Fatalf("missing revision error %q", report.Refusals[3].Error)
	}
	if !strings.Contains(report.Refusals[4].Error, "sha256") {
		t.Fatalf("rev-token outer CAS error %q", report.Refusals[4].Error)
	}
	if !strings.Contains(report.Refusals[5].Error, "stale outer revision") {
		t.Fatalf("stale CAS error %q", report.Refusals[5].Error)
	}
	if report.FatalProbe.Name != "synthetic_fatal" || !report.FatalProbe.Fatal || !strings.Contains(report.FatalProbe.Error, "recovered panic") {
		t.Fatalf("fatal envelope was not preserved by Node contract: %+v", report.FatalProbe)
	}
	if report.LegacyProbe.Name != "legacy_without_fatal" || report.LegacyProbe.Fatal || report.LegacyProbe.Error != "legacy refusal" {
		t.Fatalf("legacy envelope compatibility was not preserved by Node contract: %+v", report.LegacyProbe)
	}
	if report.Extract != string(want) {
		t.Fatalf("extract after refusals disagreed with in-process Go (%d vs %d bytes)", len(report.Extract), len(want))
	}
}

func TestWASMExtractEmptyDoesNotWriteUndefined(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	empty := filepath.Join(t.TempDir(), "empty.xlsx")
	if err := os.WriteFile(empty, nil, 0o600); err != nil {
		t.Fatal(err)
	}
	cmd := exec.Command("node", script, "extract",
		"--wasm", wasm,
		"--wasm-exec", wasmExec,
		"--input", empty,
	)
	stdout, err := cmd.Output()
	if err == nil {
		t.Fatalf("empty extract succeeded: %q", stdout)
	}
	stderr := stderrFrom(err)
	if stderr == "" || stderr == "undefined\n" || strings.Contains(stderr, "Go program has already exited") {
		t.Fatalf("empty extract did not print a catchable error: stdout=%q stderr=%q err=%v", stdout, stderr, err)
	}
	if len(stdout) != 0 || bytes.Equal(bytes.TrimSpace(stdout), []byte("undefined")) {
		t.Fatalf("empty extract wrote stdout %q", stdout)
	}

	want := extractNativeJSON(t, readHappyTree(t))
	okCmd := exec.Command("node", script, "extract",
		"--wasm", wasm,
		"--wasm-exec", wasmExec,
		"--input", testdata(t, "excel-authored", "happy-tree.xlsx"),
	)
	got, err := okCmd.Output()
	if err != nil {
		t.Fatalf("happy-tree extract after empty failure (new process) failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(bytes.TrimSpace(got), want) {
		t.Fatal("happy-tree extract JSON disagreed with in-process Go")
	}
}

func stderrFrom(err error) string {
	if ee, ok := err.(*exec.ExitError); ok {
		return string(ee.Stderr)
	}
	return ""
}
