package officecompat_test

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func TestPPTXCompatibilityCorpus(t *testing.T) {
	t.Parallel()

	manifestData, err := os.ReadFile(filepath.Join("corpus", "manifest.json"))
	if err != nil {
		t.Fatal(err)
	}
	var manifest corpus.Manifest
	if err := json.Unmarshal(manifestData, &manifest); err != nil {
		t.Fatal(err)
	}

	want := map[string]bool{
		"pptx-adversarial-encoded-traversal":     false,
		"pptx-duplicate-relationship-id":         false,
		"pptx-malformed-dialect-mix":             false,
		"pptx-resource-depth-refusal":            false,
		"pptx-strict-picture":                    false,
		"pptx-transitional-common":               false,
		"pptx-transitional-fidelity-diagnostics": false,
	}
	for _, fixture := range manifest.Fixtures {
		if fixture.Format != "pptx" {
			continue
		}
		fixture := fixture
		if _, owned := want[fixture.ID]; owned {
			want[fixture.ID] = true
		}
		t.Run(fixture.ID, func(t *testing.T) {
			t.Parallel()
			testPPTXCompatibilityFixture(t, fixture)
		})
	}
	for id, seen := range want {
		if !seen {
			t.Errorf("PPTX compatibility fixture %q is absent from the generated manifest", id)
		}
	}
}

func TestPPTXCorpusSpecsContainNoLegacyReconstructionAuthority(t *testing.T) {
	forbidden := []string{"mammoth", "luckyexcel", "domparser", "jszip", "innerhtml", "<html", "browser zip"}
	paths, err := filepath.Glob(filepath.Join("corpus", "specs", "pptx", "*.json"))
	if err != nil {
		t.Fatal(err)
	}
	if len(paths) != 7 {
		t.Fatalf("PPTX spec count = %d, want 7", len(paths))
	}
	for _, path := range paths {
		data, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		lower := strings.ToLower(string(data))
		for _, token := range forbidden {
			if strings.Contains(lower, token) {
				t.Errorf("%s contains forbidden legacy reconstruction token %q", path, token)
			}
		}
		if !strings.Contains(string(data), `"license": "CC0-1.0"`) {
			t.Errorf("%s does not carry CC0-1.0 provenance", path)
		}
	}
}

type pptxCapturedCapability struct {
	request pptxpatch.NativePassthroughTokenRequest
	token   string
}

func testPPTXCompatibilityFixture(t *testing.T, fixture corpus.FixtureRecord) {
	t.Helper()

	packageData, err := os.ReadFile(filepath.Join("corpus", filepath.FromSlash(fixture.Package)))
	if err != nil {
		t.Fatal(err)
	}
	expectedData, err := os.ReadFile(filepath.Join("corpus", filepath.FromSlash(fixture.Expected)))
	if err != nil {
		t.Fatal(err)
	}
	var expected corpus.Expectation
	if err := json.Unmarshal(expectedData, &expected); err != nil {
		t.Fatal(err)
	}
	if expected.FixtureID != fixture.ID || expected.Format != "pptx" || expected.Dialect != fixture.Dialect || expected.Outcome != fixture.Outcome {
		t.Fatalf("expectation envelope does not match manifest record: expectation=%+v fixture=%+v", expected, fixture)
	}
	if got := sha256.Sum256(packageData); hex.EncodeToString(got[:]) != fixture.SHA256 {
		t.Fatalf("package SHA-256 does not match manifest for %q", fixture.ID)
	}
	if got := sha256.Sum256(expectedData); hex.EncodeToString(got[:]) != fixture.ExpectedSHA256 {
		t.Fatalf("expectation SHA-256 does not match manifest for %q", fixture.ID)
	}
	if fixture.Provenance.Kind != "generated" || fixture.Provenance.License != "CC0-1.0" {
		t.Fatalf("fixture lacks deterministic redistributable provenance: %+v", fixture.Provenance)
	}

	var captured []pptxCapturedCapability
	deck, extractErr := pptxpatch.ExtractNativePPTX(packageData, pptxpatch.NativePPTXExtractOptions{
		TokenFactory: pptxpatch.NativePassthroughTokenFactoryFunc(func(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
			token := pptxCorpusCapabilityToken(request)
			captured = append(captured, pptxCapturedCapability{request: request, token: token})
			return token, nil
		}),
	})
	if fixture.Outcome == "refused" {
		if extractErr == nil {
			t.Fatal("explicitly refused PPTX fixture was accepted")
		}
		if got := classifyPPTXCorpusRefusal(extractErr); got != expected.Refusal.Class {
			t.Fatalf("refusal class = %q, want %q: %v", got, expected.Refusal.Class, extractErr)
		}
		if !strings.Contains(extractErr.Error(), expected.Refusal.Contains) {
			t.Fatalf("refusal %q does not contain stable marker %q", extractErr, expected.Refusal.Contains)
		}
		return
	}
	if extractErr != nil {
		t.Fatalf("accepted PPTX fixture was refused: %v", extractErr)
	}
	if issues := pptxpatch.ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("extractor returned an invalid native deck: %+v", issues)
	}

	actualJSON, err := pptxpatch.MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	roundTripped, err := pptxpatch.DecodeNativePPTXJSON(actualJSON)
	if err != nil {
		t.Fatal(err)
	}
	roundTripJSON, err := pptxpatch.MarshalNativePPTXJSON(roundTripped)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(actualJSON, roundTripJSON) {
		t.Fatal("native PPTX JSON changed across decode/marshal round trip")
	}
	expectedDeck, err := pptxpatch.DecodeNativePPTXJSON(expected.Native)
	if err != nil {
		t.Fatalf("decode expected native JSON: %v", err)
	}
	expectedJSON, err := pptxpatch.MarshalNativePPTXJSON(expectedDeck)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(actualJSON, expectedJSON) {
		t.Fatalf("native PPTX expectation drifted\nactual: %s\nexpected: %s", actualJSON, expectedJSON)
	}
	again, err := pptxpatch.ExtractNativePPTX(packageData, pptxCorpusExtractOptions(nil))
	if err != nil {
		t.Fatalf("repeat native PPTX extraction: %v", err)
	}
	againJSON, err := pptxpatch.MarshalNativePPTXJSON(again)
	if err != nil || !bytes.Equal(actualJSON, againJSON) {
		t.Fatalf("repeated native PPTX extraction drifted: err=%v\nfirst: %s\nagain: %s", err, actualJSON, againJSON)
	}

	assertPPTXCapabilitiesExactlyBound(t, deck, captured)
	assertPPTXFixtureCoverage(t, fixture.ID, deck, captured)
	if fixture.ID == "pptx-strict-picture" {
		assertPPTXRelationshipClosure(t, captured)
		assertPPTXRepackStability(t, packageData, deck, actualJSON)
		if refused, err := pptxpatch.ExtractNativePPTX(packageData, pptxpatch.NativePPTXExtractOptions{}); err == nil || len(refused.Slides) != 0 || !strings.Contains(err.Error(), "requires a trusted passthrough token factory") {
			t.Fatalf("source-dependent PPTX was not refused without capability authority: deck=%+v err=%v", refused, err)
		}
	}
}

func pptxCorpusExtractOptions(captured *[]pptxCapturedCapability) pptxpatch.NativePPTXExtractOptions {
	return pptxpatch.NativePPTXExtractOptions{TokenFactory: pptxpatch.NativePassthroughTokenFactoryFunc(func(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
		token := pptxCorpusCapabilityToken(request)
		if captured != nil {
			*captured = append(*captured, pptxCapturedCapability{request: request, token: token})
		}
		return token, nil
	})}
}

func assertPPTXRelationshipClosure(t *testing.T, captured []pptxCapturedCapability) {
	t.Helper()
	want := map[string]bool{
		"_rels/.rels":                             false,
		"relocated/_rels/deck.xml.rels":           false,
		"relocated/slides/_rels/slide-a.xml.rels": false,
		"relocated/layouts/_rels/layout.xml.rels": false,
		"relocated/masters/_rels/master.xml.rels": false,
	}
	for _, capability := range captured {
		if capability.request.Reason == "pptx.relationship-map-preserve" {
			if _, expected := want[capability.request.OwnerPart]; expected {
				want[capability.request.OwnerPart] = true
			}
		}
	}
	for part, seen := range want {
		if !seen {
			t.Errorf("relationship closure part %q was not capability-bound", part)
		}
	}
}

func assertPPTXRepackStability(t *testing.T, packageData []byte, original pptxpatch.NativePPTXDeck, originalJSON []byte) {
	t.Helper()
	repacked := repackWithStoredEntries(t, packageData)
	if bytes.Equal(packageData, repacked) {
		t.Fatal("PPTX repack fixture did not alter ZIP container bytes")
	}
	before, err := officecompat.Inspect(packageData)
	if err != nil {
		t.Fatal(err)
	}
	after, err := officecompat.Inspect(repacked)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(before, after) {
		t.Fatalf("PPTX OPC inventory changed across order/compression-only repack:\nbefore: %+v\nafter: %+v", before, after)
	}

	var captured []pptxCapturedCapability
	repackedDeck, err := pptxpatch.ExtractNativePPTX(repacked, pptxCorpusExtractOptions(&captured))
	if err != nil {
		t.Fatalf("extract repacked PPTX: %v", err)
	}
	if original.SourceRevision == nil || repackedDeck.SourceRevision == nil || *original.SourceRevision == *repackedDeck.SourceRevision {
		t.Fatalf("exact package revision did not distinguish ZIP container bytes: original=%v repacked=%v", original.SourceRevision, repackedDeck.SourceRevision)
	}
	assertPPTXCapabilitiesExactlyBound(t, repackedDeck, captured)
	assertPPTXRelationshipClosure(t, captured)
	repackedDeck.SourceRevision = original.SourceRevision
	repackedJSON, err := pptxpatch.MarshalNativePPTXJSON(repackedDeck)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(originalJSON, repackedJSON) {
		t.Fatalf("native PPTX semantics changed across order/compression-only repack\noriginal: %s\nrepacked: %s", originalJSON, repackedJSON)
	}
}

func pptxCorpusCapabilityToken(request pptxpatch.NativePassthroughTokenRequest) string {
	digest := sha256.Sum256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))
	return "token-" + hex.EncodeToString(digest[:])[:24]
}

func assertPPTXCapabilitiesExactlyBound(t *testing.T, deck pptxpatch.NativePPTXDeck, captured []pptxCapturedCapability) {
	t.Helper()
	if deck.SourceRevision == nil || len(captured) == 0 {
		t.Fatalf("parsed fixture lacks source revision or capability requests: revision=%v requests=%d", deck.SourceRevision, len(captured))
	}
	refs := map[string]pptxpatch.NativePassthroughRef{}
	for _, asset := range deck.Assets {
		for _, ref := range asset.Passthrough {
			refs[ref.Token] = ref
		}
	}
	var walk func([]pptxpatch.NativeElement)
	walk = func(elements []pptxpatch.NativeElement) {
		for _, element := range elements {
			for _, ref := range element.Passthrough {
				refs[ref.Token] = ref
			}
			walk(element.Children)
		}
	}
	for _, slide := range deck.Slides {
		for _, ref := range slide.Passthrough {
			refs[ref.Token] = ref
		}
		walk(slide.Elements)
	}
	for _, capability := range captured {
		request := capability.request
		fingerprint := sha256.Sum256(request.Payload)
		if request.SourceRevision != *deck.SourceRevision || request.ByteLength != int64(len(request.Payload)) || request.FingerprintSHA256 != hex.EncodeToString(fingerprint[:]) {
			t.Fatalf("capability request is not exactly source-bound: %+v", request)
		}
		ref, ok := refs[capability.token]
		if !ok || ref.OwnerPart != request.OwnerPart || ref.FingerprintSHA256 != request.FingerprintSHA256 || ref.Disposition != pptxpatch.NativePassthroughDispositionPreserve {
			t.Fatalf("issued capability is not represented by an exact native passthrough ref: request=%+v ref=%+v found=%v", request, ref, ok)
		}
	}
}

func assertPPTXFixtureCoverage(t *testing.T, fixtureID string, deck pptxpatch.NativePPTXDeck, captured []pptxCapturedCapability) {
	t.Helper()
	switch fixtureID {
	case "pptx-transitional-common":
		if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 5 {
			t.Fatalf("common fixture lost text/AutoShapes: %+v", deck.Slides)
		}
		text := deck.Slides[0].Elements[0]
		if text.Paragraphs == nil || len(*text.Paragraphs) != 1 || len((*text.Paragraphs)[0].Runs) != 2 || (*text.Paragraphs)[0].Runs[0].Text == nil || *(*text.Paragraphs)[0].Runs[0].Text != " Native " {
			t.Fatalf("xml:space lexical text was not retained: %+v", text)
		}
		if text.TextBody == nil || text.TextBody.LeftInsetEMU == nil || *text.TextBody.LeftInsetEMU != 91440 || text.TextBody.TopInsetEMU == nil || *text.TextBody.TopInsetEMU != 45720 || text.TextBody.Wrap != pptxpatch.NativeTextWrapSquare || text.TextBody.VerticalAnchor != pptxpatch.NativeTextVerticalAnchorTop {
			t.Fatalf("default text-body layout was not materialized exactly: %+v", text.TextBody)
		}
	case "pptx-strict-picture":
		if len(deck.Assets) != 1 || deck.Assets[0].Source == nil || deck.Assets[0].Source.PartName != "RELOCATED/MEDIA/%49MAGE.PNG" {
			t.Fatalf("Strict picture did not retain actual OPC spelling: %+v", deck.Assets)
		}
	case "pptx-transitional-fidelity-diagnostics":
		if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 2 || deck.Slides[0].Elements[0].Compatibility.Status != pptxpatch.NativeCompatibilityStatusRefused || deck.Slides[0].Elements[1].Compatibility.Status != pptxpatch.NativeCompatibilityStatusPreserveOnly {
			t.Fatalf("unsupported content was approximated or lost: %+v", deck.Slides)
		}
		for _, code := range []string{"pptx.autoshape-geometry-unavailable", "pptx.autoshape-effects-unavailable", "pptx.picture-crop-unavailable", "pptx.picture-effects-unavailable"} {
			if !pptxDeckHasDiagnostic(deck, code) {
				t.Fatalf("missing explicit fidelity diagnostic %q", code)
			}
		}
		shapeRequest := findPPTXCapabilityRequest(captured, "pptx.autoshape-refused")
		pictureRequest := findPPTXCapabilityRequest(captured, "pptx.picture-preserve-only")
		if shapeRequest == nil || !bytes.HasPrefix(shapeRequest.Payload, []byte(`<p:sp data-lexical="preserve">`)) || pictureRequest == nil || !bytes.HasPrefix(pictureRequest.Payload, []byte("<p:pic>")) {
			t.Fatalf("fidelity capabilities are not bound to exact object-local subtrees: shape=%+v picture=%+v", shapeRequest, pictureRequest)
		}
	}
}

func findPPTXCapabilityRequest(captured []pptxCapturedCapability, reason string) *pptxpatch.NativePassthroughTokenRequest {
	for index := range captured {
		if captured[index].request.Reason == reason {
			return &captured[index].request
		}
	}
	return nil
}

func pptxDeckHasDiagnostic(deck pptxpatch.NativePPTXDeck, code string) bool {
	for _, diagnostic := range deck.Compatibility.Diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

func classifyPPTXCorpusRefusal(err error) string {
	message := err.Error()
	switch {
	case strings.Contains(message, "invalid or duplicate relationship"):
		return "ambiguous-relationship-map"
	case strings.Contains(message, "encoded traversal or separator"):
		return "invalid-opc-target"
	case strings.Contains(message, "relationship dialect does not match presentation root"):
		return "mixed-ooxml-dialect"
	case strings.Contains(message, "resource budget exceeded"):
		return "xml-resource-budget"
	default:
		return fmt.Sprintf("unclassified:%T", err)
	}
}
