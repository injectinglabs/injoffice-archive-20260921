package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"os"
	"path/filepath"
	"reflect"
	"testing"
)

func TestNativeXLSXV2SharedFixtureRoundTripsCanonicalGoEncoding(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v2/valid/lexical-render.json")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := DecodeNativeWorkbookV2(data)
	if err != nil {
		t.Fatalf("shared Go/TS v2 fixture does not satisfy Go: %v", err)
	}
	encoded, err := EncodeNativeWorkbookV2(workbook)
	if err != nil {
		t.Fatal(err)
	}
	if os.Getenv("UPDATE_XLSX_NATIVE_V2_FIXTURES") == "1" {
		if err := os.WriteFile("testdata/native-xlsx-v2/valid/lexical-render.json", append(encoded, '\n'), 0o644); err != nil {
			t.Fatal(err)
		}
		return
	}
	if !bytes.Equal(encoded, bytes.TrimSpace(data)) {
		t.Fatalf("shared v2 fixture is not canonical Go encoding\n got: %s\nwant: %s", encoded, bytes.TrimSpace(data))
	}
}

func TestExcelAuthoredNativeV2FixtureMatchesExtract(t *testing.T) {
	original := readExcelAuthoredFixture(t, "happy-tree.xlsx", "c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513")
	workbook, err := ExtractNativeWorkbookV2(original)
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := EncodeNativeWorkbookV2(workbook)
	if err != nil {
		t.Fatal(err)
	}
	if os.Getenv("UPDATE_XLSX_NATIVE_V2_FIXTURES") == "1" {
		if err := os.WriteFile("testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json", append(encoded, '\n'), 0o644); err != nil {
			t.Fatal(err)
		}
		return
	}
	expected, err := os.ReadFile("testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(append(encoded, '\n'), expected) {
		t.Fatal("committed TypeScript v2 bridge is not the exact Go Extract/Encode output")
	}
}

func TestNativeXLSXV2HappyTreeFixtureRoundTripsCanonicalGoEncoding(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := DecodeNativeWorkbookV2(data)
	if err != nil {
		t.Fatalf("happy-tree v2 fixture does not satisfy Go: %v", err)
	}
	encoded, err := EncodeNativeWorkbookV2(workbook)
	if err != nil {
		t.Fatal(err)
	}
	if os.Getenv("UPDATE_XLSX_NATIVE_V2_FIXTURES") == "1" {
		if err := os.WriteFile("testdata/native-xlsx-v2/valid/excel-authored-happy-tree.json", append(encoded, '\n'), 0o644); err != nil {
			t.Fatal(err)
		}
		return
	}
	if !bytes.Equal(encoded, bytes.TrimSpace(data)) {
		t.Fatalf("happy-tree v2 fixture is not canonical Go encoding")
	}
}

func TestExtractNativeWorkbookV2EmitsCanonicalCapabilities(t *testing.T) {
	original, err := os.ReadFile(filepath.Join("testdata", "excel-authored", "happy-tree.xlsx"))
	if err != nil {
		t.Fatal(err)
	}
	v1, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	v2, err := ExtractNativeWorkbookV2(original)
	if err != nil {
		t.Fatal(err)
	}
	if v2.Protocol != NativeXLSXV2Protocol || v2.Version != 2 || v2.SchemaSHA256 != NativeXLSXV2SchemaSHA256 {
		t.Fatalf("v2 identity mismatch: protocol=%q version=%d schema=%q", v2.Protocol, v2.Version, v2.SchemaSHA256)
	}
	want := canonicalNativeWorkbookV2Capabilities()
	if len(v2.Capabilities) != len(want) {
		t.Fatalf("v2 capabilities = %#v", v2.Capabilities)
	}
	for index, capability := range v2.Capabilities {
		if capability.Name != want[index].Name || capability.Level != want[index].Level {
			t.Fatalf("capability %d = %#v, want %#v", index, capability, want[index])
		}
	}
	if v1.DocumentID != v2.DocumentID || v1.Revision != v2.Revision || v1.Source.PackageSHA256 != v2.Source.PackageSHA256 {
		t.Fatal("v2 extract lost v1 package identity")
	}
	v1Sheets, _ := json.Marshal(v1.Sheets)
	v2Sheets, _ := json.Marshal(v2.Sheets)
	if !bytes.Equal(v1Sheets, v2Sheets) {
		t.Fatal("v2 extract drifted from lockstep v1 sheet projection")
	}
}

func TestNativeXLSXV2RejectsV1CapabilityInventory(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v2/valid/lexical-render.json")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := DecodeNativeWorkbookV2(data)
	if err != nil {
		t.Fatal(err)
	}
	workbook.Capabilities = []NativeWorkbookCapabilityV2{
		{Name: "native-ooxml-parse", Level: "read-only"},
		{Name: "native-v1-mutations", Level: "partial"},
		{Name: "unsupported-content", Level: "preserve-exact"},
	}
	if issues := ValidateNativeWorkbookV2(workbook); len(issues) == 0 {
		t.Fatal("v1 3-tuple capabilities were accepted as v2")
	}
}

func TestNativeXLSXV2SchemaBindingsMatchGoWireTypes(t *testing.T) {
	types := map[string]reflect.Type{
		"NativeWorkbookV2":                reflect.TypeOf(NativeWorkbookV2{}),
		"NativeWorkbookSourceV2":          reflect.TypeOf(NativeWorkbookSourceV2{}),
		"NativeWorkbookSheetV2":           reflect.TypeOf(NativeWorkbookSheetV2{}),
		"NativeWorkbookCellV2":            reflect.TypeOf(NativeWorkbookCellV2{}),
		"NativeWorkbookValueV2":           reflect.TypeOf(NativeWorkbookValueV2{}),
		"NativeWorkbookRichRunV2":         reflect.TypeOf(NativeWorkbookRichRunV2{}),
		"NativeWorkbookFormulaV2":         reflect.TypeOf(NativeWorkbookFormulaV2{}),
		"NativeWorkbookMergedRangeV2":     reflect.TypeOf(NativeWorkbookMergedRangeV2{}),
		"NativeWorkbookRowDimensionV2":    reflect.TypeOf(NativeWorkbookRowDimensionV2{}),
		"NativeWorkbookSheetFormatV2":     reflect.TypeOf(NativeWorkbookSheetFormatV2{}),
		"NativeWorkbookNormalStyleV2":     reflect.TypeOf(NativeWorkbookNormalStyleV2{}),
		"NativeWorkbookColumnDimensionV2": reflect.TypeOf(NativeWorkbookColumnDimensionV2{}),
		"NativeWorkbookStyleV2":           reflect.TypeOf(NativeWorkbookStyleV2{}),
		"NativeWorkbookEffectiveStyleV2":  reflect.TypeOf(NativeWorkbookEffectiveStyleV2{}),
		"NativeWorkbookFillV2":            reflect.TypeOf(NativeWorkbookFillV2{}),
		"NativeWorkbookBorderV2":          reflect.TypeOf(NativeWorkbookBorderV2{}),
		"NativeWorkbookBorderSideV2":      reflect.TypeOf(NativeWorkbookBorderSideV2{}),
		"NativeWorkbookCapabilityV2":      reflect.TypeOf(NativeWorkbookCapabilityV2{}),
		"NativeWorkbookPassthroughPartV2": reflect.TypeOf(NativeWorkbookPassthroughPartV2{}),
		"NativeWorkbookUnsupportedV2":     reflect.TypeOf(NativeWorkbookUnsupportedV2{}),
	}
	if len(types) != len(nativeXLSXV2BindingShapes) {
		t.Fatalf("Go/schema object binding count differs: Go=%d schema=%d", len(types), len(nativeXLSXV2BindingShapes))
	}
	for name, shape := range nativeXLSXV2BindingShapes {
		typeOf, ok := types[name]
		if !ok {
			t.Fatalf("schema binding %q has no Go wire type", name)
		}
		properties, required, propertyTypes := nativeGoJSONShape(typeOf)
		if !reflect.DeepEqual(properties, shape.Properties) {
			t.Errorf("%s properties differ: Go=%v schema=%v", name, properties, shape.Properties)
		}
		if !reflect.DeepEqual(required, shape.Required) {
			t.Errorf("%s required fields differ: Go=%v schema=%v", name, required, shape.Required)
		}
		if !reflect.DeepEqual(propertyTypes, shape.Types) {
			t.Errorf("%s property types differ: Go=%v schema=%v", name, propertyTypes, shape.Types)
		}
	}
}

func TestNativeXLSXV2SchemaLimitsMatchGoContract(t *testing.T) {
	if nativeXLSXV2SchemaProtocol != NativeXLSXV2Protocol || nativeXLSXV2SchemaVersion != NativeXLSXV2Version {
		t.Fatalf("schema protocol/version %q/%d differ from Go %q/%d", nativeXLSXV2SchemaProtocol, nativeXLSXV2SchemaVersion, NativeXLSXV2Protocol, NativeXLSXV2Version)
	}
	if NativeXLSXV2SchemaSHA256 == "" || NativeXLSXV2SchemaSHA256[:7] != "sha256:" {
		t.Fatalf("v2 schema digest is not a sha256 identity: %q", NativeXLSXV2SchemaSHA256)
	}
	checks := map[string][2]int{
		"JSON bytes":    {nativeXLSXV2SchemaMaxJsonBytes, NativeXLSXMaxJSONBytes},
		"JSON depth":    {nativeXLSXV2SchemaMaxJsonDepth, nativeWorkbookMaxJSONDepthV2},
		"JSON tokens":   {nativeXLSXV2SchemaMaxJsonTokens, nativeWorkbookMaxJSONTokensV2},
		"sheets":        {nativeXLSXV2SchemaMaxSheets, NativeXLSXMaxSheets},
		"rows":          {nativeXLSXV2SchemaMaxRows, NativeXLSXMaxRows},
		"columns":       {nativeXLSXV2SchemaMaxColumns, NativeXLSXMaxColumns},
		"cells":         {nativeXLSXV2SchemaMaxCells, NativeXLSXMaxCells},
		"merged ranges": {nativeXLSXV2SchemaMaxMergedRanges, NativeXLSXMaxMergedRanges},
		"inventory":     {nativeXLSXV2SchemaMaxInventory, NativeXLSXMaxInventory},
		"text bytes":    {nativeXLSXV2SchemaMaxTextBytes, NativeXLSXMaxTextLength},
	}
	for name, values := range checks {
		if values[0] != values[1] {
			t.Errorf("%s schema limit %d differs from Go limit %d", name, values[0], values[1])
		}
	}
}
