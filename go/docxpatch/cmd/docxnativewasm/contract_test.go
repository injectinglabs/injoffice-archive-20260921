//go:build !js

package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

const (
	wordprocessingML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	packageRels      = "http://schemas.openxmlformats.org/package/2006/relationships"
	officeRels       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
)

func buildContractDOCX(t *testing.T) []byte {
	t.Helper()
	parts := []struct {
		name string
		data string
	}{
		{
			name: "[Content_Types].xml",
			data: `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
				`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
				`<Default Extension="xml" ContentType="application/xml"/>` +
				`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
				`</Types>`,
		},
		{
			name: "_rels/.rels",
			data: `<Relationships xmlns="` + packageRels + `"><Relationship Id="office" Type="` + officeRels + `officeDocument" Target="word/document.xml"/></Relationships>`,
		},
		{
			name: "word/document.xml",
			data: `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="` + wordprocessingML + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="01020304"><w:r><w:t>Before</w:t></w:r></w:p></w:body></w:document>`,
		},
	}
	var buffer bytes.Buffer
	writer := zip.NewWriter(&buffer)
	for _, part := range parts {
		header := &zip.FileHeader{Name: part.name, Method: zip.Store}
		entry, err := writer.CreateHeader(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := entry.Write([]byte(part.data)); err != nil {
			t.Fatal(err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatal(err)
	}
	return buffer.Bytes()
}

func extractNativeJSON(t *testing.T, data []byte) (*docxpatch.NativeDocumentV1, []byte) {
	t.Helper()
	document, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := docxpatch.EncodeNativeDocumentV1(document)
	if err != nil {
		t.Fatal(err)
	}
	return document, encoded
}

func mutationPayload(t *testing.T, document *docxpatch.NativeDocumentV1, text string) []byte {
	t.Helper()
	if len(document.Body.Blocks) != 1 || document.Body.Blocks[0].Paragraph == nil {
		t.Fatalf("contract fixture body = %#v", document.Body.Blocks)
	}
	paragraph := document.Body.Blocks[0].Paragraph
	if len(paragraph.Runs) != 1 || paragraph.Runs[0].Text == nil {
		t.Fatalf("contract fixture paragraph = %#v", paragraph)
	}
	run := paragraph.Runs[0]
	payload, err := json.Marshal(struct {
		Mutations []docxpatch.NativeDOCXTextMutationV1 `json:"mutations"`
	}{Mutations: []docxpatch.NativeDOCXTextMutationV1{{
		TargetKind:        "run",
		TargetID:          run.ID,
		ExpectedXMLSHA256: run.Anchor.XMLSHA256,
		Text:              text,
	}}})
	if err != nil {
		t.Fatal(err)
	}
	return payload
}

func paragraphMutationPayload(t *testing.T, document *docxpatch.NativeDocumentV1, text string) []byte {
	t.Helper()
	if len(document.Body.Blocks) != 1 || document.Body.Blocks[0].Paragraph == nil {
		t.Fatalf("contract fixture body = %#v", document.Body.Blocks)
	}
	paragraph := document.Body.Blocks[0].Paragraph
	if len(paragraph.Runs) != 1 || paragraph.Runs[0].Kind != "text" || paragraph.Runs[0].Text == nil {
		t.Fatalf("contract fixture is not a single-text-run paragraph: %#v", paragraph)
	}
	payload, err := json.Marshal(struct {
		Mutations []docxpatch.NativeDOCXTextMutationV1 `json:"mutations"`
	}{Mutations: []docxpatch.NativeDOCXTextMutationV1{{
		TargetKind:        "paragraph",
		TargetID:          paragraph.ID,
		ExpectedXMLSHA256: paragraph.Anchor.XMLSHA256,
		Text:              text,
	}}})
	if err != nil {
		t.Fatal(err)
	}
	return payload
}

func wasmExecFromGOROOT() string {
	if configured := strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM_EXEC")); configured != "" {
		return configured
	}
	if _, err := os.Stat(filepath.Join("dist", "wasm_exec.js")); err == nil {
		return filepath.Join("dist", "wasm_exec.js")
	}
	root := strings.TrimSpace(os.Getenv("GOROOT"))
	if root == "" {
		output, err := exec.Command("go", "env", "GOROOT").Output()
		if err != nil {
			return ""
		}
		root = strings.TrimSpace(string(output))
	}
	for _, relative := range []string{filepath.Join("misc", "wasm", "wasm_exec.js"), filepath.Join("lib", "wasm", "wasm_exec.js")} {
		candidate := filepath.Join(root, relative)
		if _, err := os.Stat(candidate); err == nil {
			return candidate
		}
	}
	return ""
}

func lookupExistingWASM() (wasm, wasmExec string, ok bool) {
	if configured := strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM")); configured != "" {
		wasm = configured
	} else {
		wasm = filepath.Join("dist", "docxnative.wasm")
	}
	wasmExec = wasmExecFromGOROOT()
	if _, err := os.Stat(wasm); err != nil || wasmExec == "" {
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

func ensureWASM(t *testing.T) (wasm, wasmExec string) {
	t.Helper()
	if wasm, wasmExec, ok := lookupExistingWASM(); ok {
		return wasm, wasmExec
	}
	if strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM_SKIP_BUILD")) != "" {
		if strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatal("docxnative.wasm is required but was not found; run ./build.sh")
		}
		t.Skip("docxnative.wasm not built; run ./build.sh")
	}
	wasmBuild.once.Do(func() {
		directory, err := os.MkdirTemp("", "docxnativewasm-")
		if err != nil {
			wasmBuild.err = err
			return
		}
		wasmBuild.wasm = filepath.Join(directory, "docxnative.wasm")
		command := exec.Command("go", "build", "-trimpath", "-ldflags=-s -w", "-o", wasmBuild.wasm, ".")
		command.Env = append(os.Environ(), "GOOS=js", "GOARCH=wasm")
		if output, buildErr := command.CombinedOutput(); buildErr != nil {
			wasmBuild.err = fmt.Errorf("%w\n%s", buildErr, output)
			return
		}
		wasmBuild.exec = wasmExecFromGOROOT()
		if wasmBuild.exec == "" {
			wasmBuild.err = os.ErrNotExist
		}
	})
	if wasmBuild.err != nil {
		if strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatalf("GOOS=js WASM build is required: %v", wasmBuild.err)
		}
		t.Skipf("GOOS=js WASM build unavailable; run ./build.sh: %v", wasmBuild.err)
	}
	if wasmBuild.wasm == "" || wasmBuild.exec == "" {
		if strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatal("docxnative.wasm is required but was not built; run ./build.sh")
		}
		t.Skip("docxnative.wasm not built; run ./build.sh")
	}
	return wasmBuild.wasm, wasmBuild.exec
}

func requireNodeHarness(t *testing.T) (wasm, wasmExec, script string) {
	t.Helper()
	if _, err := exec.LookPath("node"); err != nil {
		if strings.TrimSpace(os.Getenv("DOCXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatalf("node is required to instantiate GOOS=js WASM: %v", err)
		}
		t.Skip("node is required to instantiate GOOS=js WASM")
	}
	wasm, wasmExec = ensureWASM(t)
	script = "node_contract.mjs"
	if _, err := os.Stat(script); err != nil {
		t.Fatalf("missing %s", script)
	}
	return wasm, wasmExec, script
}

func writeContractInputs(t *testing.T, original, payload []byte) (originalPath, payloadPath string) {
	t.Helper()
	directory := t.TempDir()
	originalPath = filepath.Join(directory, "contract.docx")
	payloadPath = filepath.Join(directory, "payload.json")
	if err := os.WriteFile(originalPath, original, 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(payloadPath, payload, 0o600); err != nil {
		t.Fatal(err)
	}
	return originalPath, payloadPath
}

func TestWASMExtractMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := buildContractDOCX(t)
	_, wanted := extractNativeJSON(t, original)
	originalPath, _ := writeContractInputs(t, original, []byte(`{}`))
	command := exec.Command("node", script, "extract", "--wasm", wasm, "--wasm-exec", wasmExec, "--input", originalPath)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM extract failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(bytes.TrimSpace(got), wanted) {
		t.Fatalf("WASM extract JSON disagreed with in-process Go (%d vs %d bytes)", len(bytes.TrimSpace(got)), len(wanted))
	}
}

func TestWASMApplyMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := buildContractDOCX(t)
	document, _ := extractNativeJSON(t, original)
	payload := mutationPayload(t, document, "After & exact")
	wanted, err := docxpatch.ApplyNativeTextMutationPayloadV1(original, payload, document.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	originalPath, payloadPath := writeContractInputs(t, original, payload)
	command := exec.Command("node", script, "apply", "--wasm", wasm, "--wasm-exec", wasmExec, "--original", originalPath, "--payload", payloadPath, "--expected-revision", document.Source.PackageSHA256)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM apply failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(got, wanted.Package) {
		t.Fatalf("WASM apply bytes disagreed with in-process Go (%d vs %d bytes)", len(got), len(wanted.Package))
	}
	if _, err := docxpatch.ExtractNativeDocumentV1(got); err != nil {
		t.Fatalf("WASM output did not reopen: %v", err)
	}
}

func TestWASMApplySingleTextRunParagraphMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := buildContractDOCX(t)
	document, _ := extractNativeJSON(t, original)
	payload := paragraphMutationPayload(t, document, "Paragraph target")
	wanted, err := docxpatch.ApplyNativeTextMutationPayloadV1(original, payload, document.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	originalPath, payloadPath := writeContractInputs(t, original, payload)
	command := exec.Command("node", script, "apply", "--wasm", wasm, "--wasm-exec", wasmExec, "--original", originalPath, "--payload", payloadPath, "--expected-revision", document.Source.PackageSHA256)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM paragraph apply failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(got, wanted.Package) {
		t.Fatalf("WASM paragraph apply bytes disagreed with in-process Go (%d vs %d bytes)", len(got), len(wanted.Package))
	}
}

func TestWASMRefusalsKeepInstanceAlive(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := buildContractDOCX(t)
	document, wanted := extractNativeJSON(t, original)
	payload := mutationPayload(t, document, "Must not commit")
	originalPath, payloadPath := writeContractInputs(t, original, payload)
	stale := "sha256:" + strings.Repeat("0", 64)
	command := exec.Command("node", script, "survive",
		"--wasm", wasm, "--wasm-exec", wasmExec,
		"--input", originalPath, "--payload", payloadPath,
		"--expected-revision", document.Source.PackageSHA256,
		"--rev-token", document.Revision, "--stale-revision", stale,
	)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM survive probe failed: %v\n%s", err, stderrFrom(err))
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
	wantedNames := []string{"wrong_extract_arity", "wrong_apply_arity", "empty_extract", "empty_original", "empty_payload", "missing_revision", "rev_token", "stale_cas"}
	if len(report.Refusals) != len(wantedNames) {
		t.Fatalf("refusals=%d, want %d: %+v", len(report.Refusals), len(wantedNames), report.Refusals)
	}
	for index, name := range wantedNames {
		refusal := report.Refusals[index]
		if refusal.Name != name {
			t.Fatalf("refusal %d name %q, want %q", index, refusal.Name, name)
		}
		if refusal.Error == "" || refusal.Error == "undefined" || strings.Contains(refusal.Error, "Go program has already exited") {
			t.Fatalf("%s did not return a catchable error: %q", name, refusal.Error)
		}
		if refusal.Fatal {
			t.Fatalf("ordinary refusal %s was marked fatal", name)
		}
	}
	if report.FatalProbe.Name != "synthetic_fatal" || !report.FatalProbe.Fatal || !strings.Contains(report.FatalProbe.Error, "recovered panic") {
		t.Fatalf("fatal envelope was not preserved by Node contract: %+v", report.FatalProbe)
	}
	if report.LegacyProbe.Name != "legacy_without_fatal" || report.LegacyProbe.Fatal || report.LegacyProbe.Error != "legacy refusal" {
		t.Fatalf("legacy envelope compatibility was not preserved by Node contract: %+v", report.LegacyProbe)
	}
	if report.Extract != string(wanted) {
		t.Fatalf("extract after refusals disagreed with in-process Go (%d vs %d bytes)", len(report.Extract), len(wanted))
	}
}

func stderrFrom(err error) string {
	if exitError, ok := err.(*exec.ExitError); ok {
		return string(exitError.Stderr)
	}
	return ""
}
