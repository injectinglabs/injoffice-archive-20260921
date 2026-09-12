//go:build !js

package main

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"sync"
	"testing"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func TestBrowserLocalTokenIsDeterministicBoundAndNonAuthorizing(t *testing.T) {
	request := pptxpatch.NativePassthroughTokenRequest{
		SourceRevision: "rev-" + strings.Repeat("a", 64),
		OwnerPart:      "relocated/slides/slide-a.xml", ObjectID: "cNvPr-3",
		FingerprintSHA256: strings.Repeat("b", 64), ByteLength: 6,
		Reason: "pptx.autoshape-refused", Payload: []byte("opaque"),
	}
	want := browserLocalPassthroughToken(request)
	if got := browserLocalPassthroughToken(request); got != want {
		t.Fatalf("browser-local token is not deterministic: %q != %q", got, want)
	}
	if !strings.HasPrefix(want, "browser-local-v1-") || strings.HasPrefix(want, "http-") {
		t.Fatalf("browser-local token entered a server capability namespace: %q", want)
	}
	for _, secret := range []string{request.OwnerPart, request.ObjectID, request.Reason, string(request.Payload)} {
		if strings.Contains(want, secret) {
			t.Fatalf("browser-local token exposed source material %q: %q", secret, want)
		}
	}
	changed := request
	changed.Payload = []byte("opaqvf")
	if got := browserLocalPassthroughToken(changed); got == want {
		t.Fatal("browser-local token did not bind the exact passthrough payload")
	}
	changed = request
	changed.SourceRevision = "rev-" + strings.Repeat("c", 64)
	if got := browserLocalPassthroughToken(changed); got == want {
		t.Fatal("browser-local token did not bind the source revision")
	}
	transaction, err := (browserLocalTokenFactory{}).BeginNativePassthroughTokenTransaction()
	if err != nil {
		t.Fatal(err)
	}
	if got, err := transaction.IssueNativePassthroughToken(request); err != nil || got != want {
		t.Fatalf("transactional token drifted: token=%q err=%v", got, err)
	}
	if err := transaction.CommitNativePassthroughTokens(); err != nil {
		t.Fatal(err)
	}
	transaction.RollbackNativePassthroughTokens()
}

func TestWASMExtractMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := contractPPTX(t)
	want, err := extractNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	deck, err := pptxpatch.DecodeNativePPTXJSON(want)
	if err != nil {
		t.Fatal(err)
	}
	if !deckHasBrowserLocalToken(deck) || bytes.Contains(want, []byte(`"token":"http-`)) {
		t.Fatal("extract did not use only the browser-local passthrough namespace")
	}

	dir := t.TempDir()
	input := filepath.Join(dir, "contract.pptx")
	if err := os.WriteFile(input, original, 0o600); err != nil {
		t.Fatal(err)
	}
	command := exec.Command("node", script, "extract", "--wasm", wasm, "--wasm-exec", wasmExec, "--input", input)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM extract failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(bytes.TrimSpace(got), want) {
		t.Fatalf("WASM extract disagreed with in-process Go (%d vs %d bytes)", len(bytes.TrimSpace(got)), len(want))
	}
}

func TestWASMInspectionMatchesReadOnlyGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := contractPPTX(t)
	want, err := inspectNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	if bytes.Contains(want, []byte(`"token"`)) || bytes.Contains(want, []byte(`"passthrough"`)) {
		t.Fatal("inspection published preservation capabilities")
	}
	input := filepath.Join(t.TempDir(), "contract.pptx")
	if err := os.WriteFile(input, original, 0o600); err != nil {
		t.Fatal(err)
	}
	command := exec.Command("node", script, "inspect", "--wasm", wasm, "--wasm-exec", wasmExec, "--input", input)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM inspection failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(bytes.TrimSpace(got), want) {
		t.Fatal("WASM inspection disagreed with Go")
	}
}

func TestWASMApplyMatchesInProcessGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := contractPPTX(t)
	payload, outerRevision, _ := contractMutation(t, original)
	want, err := applyNativePayload(original, payload, outerRevision)
	if err != nil {
		t.Fatal(err)
	}

	dir := t.TempDir()
	input := filepath.Join(dir, "contract.pptx")
	payloadPath := filepath.Join(dir, "mutation.json")
	if err := os.WriteFile(input, original, 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(payloadPath, payload, 0o600); err != nil {
		t.Fatal(err)
	}
	command := exec.Command("node", script, "apply",
		"--wasm", wasm, "--wasm-exec", wasmExec,
		"--original", input, "--payload", payloadPath, "--expected-revision", outerRevision,
	)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM apply failed: %v\n%s", err, stderrFrom(err))
	}
	if !bytes.Equal(got, want) {
		t.Fatalf("WASM apply bytes disagreed with in-process Go (%d vs %d bytes)", len(got), len(want))
	}
}

func TestWASMRefusalsKeepInstanceAlive(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	original := contractPPTX(t)
	want, err := extractNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	payload, outerRevision, nativeRevision := contractMutation(t, original)

	dir := t.TempDir()
	input := filepath.Join(dir, "contract.pptx")
	payloadPath := filepath.Join(dir, "mutation.json")
	if err := os.WriteFile(input, original, 0o600); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(payloadPath, payload, 0o600); err != nil {
		t.Fatal(err)
	}
	command := exec.Command("node", script, "survive",
		"--wasm", wasm, "--wasm-exec", wasmExec, "--input", input, "--payload", payloadPath,
		"--expected-revision", outerRevision, "--rev-token", nativeRevision,
		"--stale-revision", "sha256:"+strings.Repeat("0", 64),
	)
	got, err := command.Output()
	if err != nil {
		t.Fatalf("WASM survival probe failed: %v\n%s", err, stderrFrom(err))
	}
	var report struct {
		Refusals []struct {
			Name  string `json:"name"`
			Error string `json:"error"`
		} `json:"refusals"`
		Extract string `json:"extract"`
	}
	if err := json.Unmarshal(got, &report); err != nil {
		t.Fatalf("survival report is not JSON: %v\n%s", err, got)
	}
	wantNames := []string{"empty_extract", "empty_original", "empty_payload", "missing_revision", "rev_token", "stale_cas"}
	if len(report.Refusals) != len(wantNames) {
		t.Fatalf("refusals=%d, want %d: %+v", len(report.Refusals), len(wantNames), report.Refusals)
	}
	for index, name := range wantNames {
		item := report.Refusals[index]
		if item.Name != name || item.Error == "" || item.Error == "undefined" || strings.Contains(item.Error, "Go program has already exited") {
			t.Fatalf("refusal %d did not return a catchable error: %+v", index, item)
		}
	}
	if !strings.Contains(report.Refusals[3].Error, "expectedRevision is required") ||
		!strings.Contains(report.Refusals[4].Error, "stale outer source revision") ||
		!strings.Contains(report.Refusals[5].Error, "stale outer source revision") {
		t.Fatalf("refusal errors lost policy detail: %+v", report.Refusals)
	}
	if report.Extract != string(want) {
		t.Fatalf("extract after refusals disagreed with in-process Go (%d vs %d bytes)", len(report.Extract), len(want))
	}
}

func contractMutation(t *testing.T, original []byte) (payload []byte, outerRevision, nativeRevision string) {
	t.Helper()
	encoded, err := extractNativeJSON(original)
	if err != nil {
		t.Fatal(err)
	}
	deck, err := pptxpatch.DecodeNativePPTXJSON(encoded)
	if err != nil {
		t.Fatal(err)
	}
	var target *pptxpatch.NativeElement
	for slideIndex := range deck.Slides {
		for elementIndex := range deck.Slides[slideIndex].Elements {
			element := &deck.Slides[slideIndex].Elements[elementIndex]
			if element.Kind == pptxpatch.NativeElementKindText && element.Paragraphs != nil && element.Source != nil {
				target = element
				break
			}
		}
	}
	if target == nil {
		t.Fatal("contract fixture has no editable source text")
	}
	paragraphJSON, err := json.Marshal(*target.Paragraphs)
	if err != nil {
		t.Fatal(err)
	}
	var paragraphs []pptxpatch.NativeParagraph
	if err := json.Unmarshal(paragraphJSON, &paragraphs); err != nil {
		t.Fatal(err)
	}
	value := "edited in pptxnative WASM"
	paragraphs[0].Runs[0].Text = &value
	request := pptxpatch.NativePPTXMutationRequest{
		ExpectedSourceRevision: *deck.SourceRevision,
		Operations: []pptxpatch.NativePPTXMutation{{
			OperationID: "wasm-replace-title", Kind: pptxpatch.NativePPTXReplaceText,
			ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256,
			Paragraphs: &paragraphs,
		}},
	}
	payload, err = json.Marshal(request)
	if err != nil {
		t.Fatal(err)
	}
	digest := sha256.Sum256(original)
	return payload, "sha256:" + hex.EncodeToString(digest[:]), *deck.SourceRevision
}

func deckHasBrowserLocalToken(deck pptxpatch.NativePPTXDeck) bool {
	for _, asset := range deck.Assets {
		for _, ref := range asset.Passthrough {
			if strings.HasPrefix(ref.Token, "browser-local-v1-") {
				return true
			}
		}
	}
	var elementHasToken func(pptxpatch.NativeElement) bool
	elementHasToken = func(element pptxpatch.NativeElement) bool {
		for _, ref := range element.Passthrough {
			if strings.HasPrefix(ref.Token, "browser-local-v1-") {
				return true
			}
		}
		for _, child := range element.Children {
			if elementHasToken(child) {
				return true
			}
		}
		return false
	}
	for _, slide := range deck.Slides {
		for _, ref := range slide.Passthrough {
			if strings.HasPrefix(ref.Token, "browser-local-v1-") {
				return true
			}
		}
		for _, element := range slide.Elements {
			if elementHasToken(element) {
				return true
			}
		}
	}
	return false
}

func contractPPTX(t *testing.T) []byte {
	t.Helper()
	const (
		contentTypes = "http://schemas.openxmlformats.org/package/2006/content-types"
		packageRels  = "http://schemas.openxmlformats.org/package/2006/relationships"
		presentation = "http://schemas.openxmlformats.org/presentationml/2006/main"
		drawing      = "http://schemas.openxmlformats.org/drawingml/2006/main"
		officeRels   = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	)
	text := `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="1" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t xml:space="preserve">Hello </a:t></a:r><a:r><a:rPr b="0" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>world</a:t></a:r></a:p></p:txBody></p:sp>`
	unsupported := `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Opaque Shape"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="100" y="200"/><a:ext cx="300000" cy="200000"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst/></a:custGeom><a:gradFill/><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:prstDash val="dash"/><a:round/></a:ln><a:effectLst><a:outerShdw/></a:effectLst></p:spPr><p:style/></p:sp>`
	parts := []struct{ name, data string }{
		{"[Content_Types].xml", fmt.Sprintf(`<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/relocated/deck.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/relocated/slides/slide-a.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/relocated/layouts/layout.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/relocated/masters/master.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/relocated/themes/theme.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`, contentTypes)},
		{"_rels/.rels", fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdRoot" Type="%s/officeDocument" Target="relocated/deck.xml"/></Relationships>`, packageRels, officeRels)},
		{"relocated/deck.xml", fmt.Sprintf(`<p:presentation xmlns:p="%s" xmlns:r="%s"><p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`, presentation, officeRels)},
		{"relocated/_rels/deck.xml.rels", fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rId7" Type="%s/slide" Target="slides/slide-a.xml"/></Relationships>`, packageRels, officeRels)},
		{"relocated/slides/slide-a.xml", fmt.Sprintf(`<p:sld xmlns:p="%s" xmlns:a="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>%s%s</p:spTree></p:cSld></p:sld>`, presentation, drawing, text, unsupported)},
		{"relocated/slides/_rels/slide-a.xml.rels", fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdLayout" Type="%s/slideLayout" Target="../layouts/layout.xml"/></Relationships>`, packageRels, officeRels)},
		{"relocated/layouts/layout.xml", fmt.Sprintf(`<p:sldLayout xmlns:p="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`, presentation)},
		{"relocated/layouts/_rels/layout.xml.rels", fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdMaster" Type="%s/slideMaster" Target="../masters/master.xml"/></Relationships>`, packageRels, officeRels)},
		{"relocated/masters/master.xml", fmt.Sprintf(`<p:sldMaster xmlns:p="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>`, presentation)},
		{"relocated/masters/_rels/master.xml.rels", fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdTheme" Type="%s/theme" Target="../themes/theme.xml"/></Relationships>`, packageRels, officeRels)},
		{"relocated/themes/theme.xml", fmt.Sprintf(`<a:theme xmlns:a="%s" name="Fixture"><a:themeElements/></a:theme>`, drawing)},
	}
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	for _, part := range parts {
		header := &zip.FileHeader{Name: part.name, Method: zip.Store}
		header.SetMode(0o600)
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
	return output.Bytes()
}

func wasmExecFromGOROOT() string {
	if configured := strings.TrimSpace(os.Getenv("PPTXNATIVE_WASM_EXEC")); configured != "" {
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

var wasmBuild struct {
	once sync.Once
	wasm string
	exec string
	err  error
}

func lookupExistingWasm() (wasm, wasmExec string, ok bool) {
	if configured := strings.TrimSpace(os.Getenv("PPTXNATIVE_WASM")); configured != "" {
		wasm = configured
	} else {
		wasm = filepath.Join("dist", "pptxnative.wasm")
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

func ensureWasm(t *testing.T) (wasm, wasmExec string) {
	t.Helper()
	if wasm, wasmExec, ok := lookupExistingWasm(); ok {
		return wasm, wasmExec
	}
	if strings.TrimSpace(os.Getenv("PPTXNATIVE_WASM_SKIP_BUILD")) != "" {
		if strings.TrimSpace(os.Getenv("PPTXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatal("pptxnative.wasm is required but was not found; run ./cmd/pptxnativewasm/build.sh")
		}
		t.Skip("pptxnative.wasm not built; run ./cmd/pptxnativewasm/build.sh")
	}
	wasmBuild.once.Do(func() {
		directory, err := os.MkdirTemp("", "pptxnativewasm-")
		if err != nil {
			wasmBuild.err = err
			return
		}
		wasmBuild.wasm = filepath.Join(directory, "pptxnative.wasm")
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
		t.Fatalf("GOOS=js WASM build unavailable: %v", wasmBuild.err)
	}
	if wasmBuild.wasm == "" || wasmBuild.exec == "" {
		t.Fatal("pptxnative.wasm was not built; run ./cmd/pptxnativewasm/build.sh")
	}
	return wasmBuild.wasm, wasmBuild.exec
}

func requireNodeHarness(t *testing.T) (wasm, wasmExec, script string) {
	t.Helper()
	if _, err := exec.LookPath("node"); err != nil {
		if strings.TrimSpace(os.Getenv("PPTXNATIVE_WASM_REQUIRED")) != "" {
			t.Fatalf("node is required to instantiate GOOS=js WASM: %v", err)
		}
		t.Skip("node is required to instantiate GOOS=js WASM")
	}
	wasm, wasmExec = ensureWasm(t)
	script = "node_contract.mjs"
	if _, err := os.Stat(script); err != nil {
		t.Fatal(err)
	}
	return wasm, wasmExec, script
}

func stderrFrom(err error) string {
	if exit, ok := err.(*exec.ExitError); ok {
		return string(exit.Stderr)
	}
	return ""
}
