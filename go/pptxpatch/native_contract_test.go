package pptxpatch

import (
	"crypto/sha256"
	"encoding/json"
	"errors"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"
)

func TestNativeContractSharedFixtures(t *testing.T) {
	valid, err := filepath.Glob("testdata/native-contract/valid/*.json")
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range valid {
		t.Run("valid/"+filepath.Base(name), func(t *testing.T) {
			data, err := os.ReadFile(name)
			if err != nil {
				t.Fatal(err)
			}
			if _, err := DecodeNativePPTXJSON(data); err != nil {
				t.Fatalf("valid fixture refused: %v", err)
			}
		})
	}

	invalid, err := filepath.Glob("testdata/native-contract/invalid/*.json")
	if err != nil {
		t.Fatal(err)
	}
	for _, name := range invalid {
		t.Run("invalid/"+filepath.Base(name), func(t *testing.T) {
			data, err := os.ReadFile(name)
			if err != nil {
				t.Fatal(err)
			}
			if _, err := DecodeNativePPTXJSON(data); err == nil {
				t.Fatal("invalid fixture was accepted")
			}
		})
	}
}

func TestNativeContractRejectsMetadataOnlyParsedAssetWithoutReadCapability(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/parsed-source-asset-missing-capability.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	if err == nil {
		t.Fatal("metadata-only parsed source asset was accepted")
	}
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected validation error, got %T: %v", err, err)
	}
	found := false
	for _, issue := range validation.Issues {
		found = found || issue.Code == "native.assetReadCapability"
	}
	if !found {
		t.Fatalf("missing asset read-capability issue: %#v", validation.Issues)
	}
}

func TestNativeContractAllowsMissingShapePresetOnlyForRefusedPlaceholder(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/editable-shape-missing-preset.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected validation error, got %T: %v", err, err)
	}
	found := false
	for _, issue := range validation.Issues {
		found = found || issue.Code == "native.shapePreset"
	}
	if !found {
		t.Fatalf("missing shape-preset issue: %#v", validation.Issues)
	}
}

func TestNativeContractRejectsPartialStrokeSemantics(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/partial-stroke-semantics.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected validation error, got %T: %v", err, err)
	}
	found := false
	for _, issue := range validation.Issues {
		found = found || issue.Code == "native.strokeMetadata"
	}
	if !found {
		t.Fatalf("missing stroke metadata issue: %#v", validation.Issues)
	}
}

func TestNativeContractRejectsOversizedStrokeAndTableBorder(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/oversized-stroke-width.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected validation error, got %T: %v", err, err)
	}
	want := map[string]bool{
		"$.slides[0].elements[0].stroke.widthEmu":                  true,
		"$.slides[0].elements[0].stroke.miterLimit":                true,
		"$.slides[0].elements[1].table.rows[0][0].border.widthEmu": true,
	}
	for _, issue := range validation.Issues {
		delete(want, issue.Path)
	}
	if len(want) != 0 {
		t.Fatalf("missing bounded-stroke refusals: %#v; issues: %#v", want, validation.Issues)
	}
	validData, err := os.ReadFile("testdata/native-contract/valid/parsed-full.json")
	if err != nil {
		t.Fatal(err)
	}
	deck, err := DecodeNativePPTXJSON(validData)
	if err != nil {
		t.Fatal(err)
	}
	join := NativeStrokeJoinMiter
	deck.Slides[0].Elements[1].Stroke.WidthEMU = int64Pointer(nativeMaxLineWidthEmu)
	deck.Slides[0].Elements[1].Stroke.Join = &join
	deck.Slides[0].Elements[1].Stroke.MiterLimit = int64Pointer(nativeMaxDrawingPercentage)
	deck.Slides[0].Elements[4].Table.Rows[1][1].Border.WidthEMU = int64Pointer(nativeMaxLineWidthEmu)
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("exact DrawingML numeric boundaries were rejected: %#v", issues)
	}
}

func TestNativeContractCanonicalGolden(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/valid/parsed-full.json")
	if err != nil {
		t.Fatal(err)
	}
	deck, err := DecodeNativePPTXJSON(data)
	if err != nil {
		t.Fatal(err)
	}
	got, err := MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	want, err := os.ReadFile("testdata/native-contract/parsed-full.canonical.json")
	if err != nil {
		t.Fatal(err)
	}
	if string(got) != string(want) {
		t.Fatalf("canonical JSON differs\ngot:  %s\nwant: %s", got, want)
	}
	if deck.Assets[0].ID != "z-picture" {
		t.Fatal("marshal mutated source asset order")
	}
}

func TestNativeContractRejectsUnknownNestedField(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/valid/parsed-full.json")
	if err != nil {
		t.Fatal(err)
	}
	data = []byte(strings.Replace(string(data), `"widthEmu": 12700`, `"widthEmu": 12700, "legacyWidth": 1`, 1))
	if _, err := DecodeNativePPTXJSON(data); err == nil || !strings.Contains(err.Error(), "unknown field") {
		t.Fatalf("expected nested unknown field refusal, got %v", err)
	}
}

func TestNativeContractRejectsNullAndDuplicateKeys(t *testing.T) {
	minimal, err := os.ReadFile("testdata/native-contract/valid/authored-minimal.json")
	if err != nil {
		t.Fatal(err)
	}
	withNull := []byte(strings.Replace(string(minimal), `"origin": "authored",`, `"origin": "authored", "sourceRevision": null,`, 1))
	if _, err := DecodeNativePPTXJSON(withNull); err == nil || !strings.Contains(err.Error(), "null is not allowed") {
		t.Fatalf("expected null refusal, got %v", err)
	}
	duplicate := []byte(strings.Replace(string(minimal), `"documentId": "deck-authored-1",`, `"documentId": "deck-authored-1", "documentId": "deck-other",`, 1))
	if _, err := DecodeNativePPTXJSON(duplicate); err == nil || !strings.Contains(err.Error(), "duplicate object key") {
		t.Fatalf("expected duplicate-key refusal, got %v", err)
	}
}

func TestNativeContractRejectsNegativeZeroLexically(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/negative-zero.json")
	if err != nil {
		t.Fatal(err)
	}
	if _, err := DecodeNativePPTXJSON(data); err == nil || !strings.Contains(err.Error(), "negative zero") {
		t.Fatalf("expected negative-zero refusal, got %v", err)
	}
}

func TestNativeContractRefusesExcessiveGroupDepth(t *testing.T) {
	zero, one := int64(0), int64(1)
	transform := NativeTransform{X: &zero, Y: &zero, Cx: &one, Cy: &one}
	compatibility := NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}}
	element := NativeElement{
		Kind: NativeElementKindConnector, ID: "leaf", Provenance: NativeProvenanceAuthored,
		Transform: transform, Passthrough: []NativePassthroughRef{}, Compatibility: compatibility,
	}
	for depth := 0; depth <= nativeMaxDepth; depth++ {
		element = NativeElement{
			Kind: NativeElementKindGroup, ID: fmt.Sprintf("group-%d", depth), Provenance: NativeProvenanceAuthored,
			Transform: transform, Children: []NativeElement{element}, Passthrough: []NativePassthroughRef{}, Compatibility: compatibility,
		}
	}
	deck := NativePPTXDeck{
		ContractVersion: NativePPTXContractVersion, DocumentID: "deck-depth", Origin: NativeOriginAuthored,
		Size: NativeSize{Cx: &one, Cy: &one}, Assets: []NativeAsset{},
		Slides: []NativeSlide{{
			ID: "slide-depth", Provenance: NativeProvenanceAuthored, Elements: []NativeElement{element},
			Passthrough: []NativePassthroughRef{}, Compatibility: compatibility,
		}},
		Compatibility: compatibility,
	}
	issues := ValidateNativePPTX(deck)
	for _, issue := range issues {
		if issue.Code == "native.resourceDepth" {
			return
		}
	}
	t.Fatalf("expected group depth refusal, got %#v", issues)
}

func TestNativeContractValidationErrorIsMachineReadable(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/duplicate-id.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected NativeContractValidationError, got %T: %v", err, err)
	}
	if len(validation.Issues) == 0 || validation.Issues[0].Path == "" || validation.Issues[0].Code == "" {
		t.Fatalf("missing structured issue: %#v", validation.Issues)
	}
	if _, err := json.Marshal(validation.Issues); err != nil {
		t.Fatalf("issues are not JSON-safe: %v", err)
	}
}

func TestNativeContractRejectsEveryOmittedRequiredZeroValue(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/missing-required-zero-values.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected NativeContractValidationError, got %T: %v", err, err)
	}
	want := map[string]bool{
		"$.assets[0].byteLength":                                   true,
		"$.slides[0].elements[0].transform.x":                      true,
		"$.slides[0].elements[0].paragraphs[0].runs[0].text":       true,
		"$.slides[0].elements[1].stroke.widthEmu":                  true,
		"$.slides[0].elements[2].table.rows[0][0].text":            true,
		"$.slides[0].elements[2].table.rows[0][0].border.widthEmu": true,
	}
	for _, issue := range validation.Issues {
		delete(want, issue.Path)
	}
	if len(want) != 0 {
		t.Fatalf("missing required-field refusals: %#v; issues: %#v", want, validation.Issues)
	}
}

func TestNativeContractRejectsMismatchedChartReferences(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/invalid/chart-reference-mismatch.json")
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativePPTXJSON(data)
	var validation NativeContractValidationError
	if !errors.As(err, &validation) {
		t.Fatalf("expected NativeContractValidationError, got %T: %v", err, err)
	}
	want := map[string]bool{
		"$.slides[0].elements[0].chart.opaqueRef.ownerPart": true,
		"$.slides[0].elements[0].chart.relationshipId":      true,
	}
	for _, issue := range validation.Issues {
		if issue.Code == "native.chartReference" {
			delete(want, issue.Path)
		}
	}
	if len(want) != 0 {
		t.Fatalf("missing chart reference refusals: %#v; issues: %#v", want, validation.Issues)
	}
}

func TestNativeContractAnchorAndScopeOwnershipVectors(t *testing.T) {
	vectors := map[string][]string{
		"chart-part-preview-type.json":             {"native.assetType"},
		"cross-slide-element-source.json":          {"native.sourcePart"},
		"parsed-element-under-authored-slide.json": {"native.sourceOwnership"},
		"diagnostic-cross-slide-scope.json":        {"native.scopeOwnership"},
	}
	for name, codes := range vectors {
		t.Run(name, func(t *testing.T) {
			data, err := os.ReadFile(filepath.Join("testdata/native-contract/invalid", name))
			if err != nil {
				t.Fatal(err)
			}
			_, err = DecodeNativePPTXJSON(data)
			var validation NativeContractValidationError
			if !errors.As(err, &validation) {
				t.Fatalf("expected NativeContractValidationError, got %T: %v", err, err)
			}
			for _, code := range codes {
				found := false
				for _, issue := range validation.Issues {
					found = found || issue.Code == code
				}
				if !found {
					t.Errorf("missing %s refusal: %#v", code, validation.Issues)
				}
			}
			if name == "chart-part-preview-type.json" {
				for _, issue := range validation.Issues {
					if issue.Code == "native.chartPart" {
						t.Errorf("folder convention must not reject a relationship-resolved chart part: %#v", validation.Issues)
					}
				}
			}
		})
	}
}

func TestNativeContractRejectsEncodedPartNameAttacks(t *testing.T) {
	for _, name := range []string{
		"unsafe-part-encoded-parent.json", "unsafe-part-encoded-slash.json",
		"unsafe-part-encoded-backslash.json", "unsafe-part-encoded-control.json",
		"unsafe-part-trailing-dot.json", "unsafe-part-encoded-trailing-dot.json",
	} {
		t.Run(name, func(t *testing.T) {
			data, err := os.ReadFile(filepath.Join("testdata/native-contract/invalid", name))
			if err != nil {
				t.Fatal(err)
			}
			if _, err := DecodeNativePPTXJSON(data); err == nil || !strings.Contains(err.Error(), "canonical package-relative part name") {
				t.Fatalf("expected encoded part-name refusal, got %v", err)
			}
		})
	}
}

func TestNativeContractSchemaHash(t *testing.T) {
	data, err := os.ReadFile("../../schemas/pptx-native-v1.schema.json")
	if err != nil {
		t.Fatal(err)
	}
	got := fmt.Sprintf("%x", sha256.Sum256(data))
	if got != NativePPTXSchemaSHA256 {
		t.Fatalf("generated bindings are stale: got %s want %s", got, NativePPTXSchemaSHA256)
	}
}

func TestNativeContractGeneratedBindingShapesMatchGo(t *testing.T) {
	bindings := map[string]reflect.Type{
		"NativePptxDeck":        reflect.TypeOf(NativePPTXDeck{}),
		"NativeSize":            reflect.TypeOf(NativeSize{}),
		"NativeTransform":       reflect.TypeOf(NativeTransform{}),
		"NativeSourceAnchor":    reflect.TypeOf(NativeSourceAnchor{}),
		"NativePassthroughRef":  reflect.TypeOf(NativePassthroughRef{}),
		"NativeDiagnosticScope": reflect.TypeOf(NativeDiagnosticScope{}),
		"NativeDiagnostic":      reflect.TypeOf(NativeDiagnostic{}),
		"NativeCompatibility":   reflect.TypeOf(NativeCompatibility{}),
		"NativeAsset":           reflect.TypeOf(NativeAsset{}),
		"NativeTextRun":         reflect.TypeOf(NativeTextRun{}),
		"NativeParagraph":       reflect.TypeOf(NativeParagraph{}),
		"NativeTextBodyLayout":  reflect.TypeOf(NativeTextBodyLayout{}),
		"NativeStroke":          reflect.TypeOf(NativeStroke{}),
		"NativeAnimation":       reflect.TypeOf(NativeAnimation{}),
		"NativeTransition":      reflect.TypeOf(NativeTransition{}),
		"NativeTableBorder":     reflect.TypeOf(NativeTableBorder{}),
		"NativeTableCell":       reflect.TypeOf(NativeTableCell{}),
		"NativeTable":           reflect.TypeOf(NativeTable{}),
		"NativeOpaqueChart":     reflect.TypeOf(NativeOpaqueChart{}),
		"NativeSlide":           reflect.TypeOf(NativeSlide{}),
	}
	for name, binding := range bindings {
		got := reflectedBindingShape(binding)
		want, ok := nativePPTXBindingShapes[name]
		if !ok {
			t.Fatalf("schema generator did not emit %s", name)
		}
		if !reflect.DeepEqual(got, want) {
			t.Errorf("%s binding drifted from schema: got %#v want %#v", name, got, want)
		}
	}

	// NativeElement is one Go discriminated-union struct. Its complete field
	// set must equal the union of the generated variant schemas, while fields
	// without omitempty must equal their common required-field intersection.
	elementNames := []string{
		"NativeTextElement", "NativeShapeElement", "NativeConnectorElement",
		"NativePictureElement", "NativeTableElement", "NativeChartElement", "NativeGroupElement",
	}
	propertyUnion := map[string]bool{}
	requiredIntersection := map[string]bool{}
	for index, name := range elementNames {
		shape := nativePPTXBindingShapes[name]
		for _, property := range shape.Properties {
			propertyUnion[property] = true
		}
		if index == 0 {
			for _, property := range shape.Required {
				requiredIntersection[property] = true
			}
			continue
		}
		required := stringSet(shape.Required)
		for property := range requiredIntersection {
			if !required[property] {
				delete(requiredIntersection, property)
			}
		}
	}
	want := nativePPTXBindingShape{
		Properties: sortedSet(propertyUnion),
		Required:   sortedSet(requiredIntersection),
	}
	if got := reflectedBindingShape(reflect.TypeOf(NativeElement{})); !reflect.DeepEqual(got, want) {
		t.Fatalf("NativeElement union binding drifted from schema: got %#v want %#v", got, want)
	}
}

func reflectedBindingShape(binding reflect.Type) nativePPTXBindingShape {
	shape := nativePPTXBindingShape{Properties: []string{}, Required: []string{}}
	for index := 0; index < binding.NumField(); index++ {
		tag := binding.Field(index).Tag.Get("json")
		parts := strings.Split(tag, ",")
		if parts[0] == "" || parts[0] == "-" {
			continue
		}
		shape.Properties = append(shape.Properties, parts[0])
		if len(parts) == 1 || parts[1] != "omitempty" {
			shape.Required = append(shape.Required, parts[0])
		}
	}
	sort.Strings(shape.Properties)
	sort.Strings(shape.Required)
	return shape
}

func stringSet(values []string) map[string]bool {
	result := make(map[string]bool, len(values))
	for _, value := range values {
		result[value] = true
	}
	return result
}

func sortedSet(values map[string]bool) []string {
	result := make([]string, 0, len(values))
	for value := range values {
		result = append(result, value)
	}
	sort.Strings(result)
	return result
}

func TestNativeContractRefusesContradictoryCompatibility(t *testing.T) {
	data, err := os.ReadFile("testdata/native-contract/valid/parsed-full.json")
	if err != nil {
		t.Fatal(err)
	}
	deck, err := DecodeNativePPTXJSON(data)
	if err != nil {
		t.Fatal(err)
	}
	deck.Compatibility = NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}}
	deck.Slides[0].Compatibility = NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}}
	issues := ValidateNativePPTX(deck)
	count := 0
	for _, issue := range issues {
		if issue.Code == "native.compatibilityAggregate" {
			count++
		}
	}
	if count != 2 {
		t.Fatalf("expected deck and slide aggregate refusals, got %#v", issues)
	}
}
