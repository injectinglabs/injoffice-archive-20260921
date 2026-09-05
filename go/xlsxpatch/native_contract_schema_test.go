package xlsxpatch

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"os"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"testing"
)

func TestNativeXLSXStyleUnsupportedDiagnosticsAreBijective(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v1/valid/lexical-render.json")
	if err != nil {
		t.Fatal(err)
	}
	decode := func() *NativeWorkbookV1 {
		workbook, decodeErr := DecodeNativeWorkbookV1(data)
		if decodeErr != nil {
			t.Fatal(decodeErr)
		}
		return workbook
	}

	missing := decode()
	missing.Styles[1].Effective.Projection = "partial"
	missing.Styles[1].Effective.Unsupported = []string{"font-color"}
	missing.Styles[1].RawProjectionSHA256, err = nativeRawStyleProjectionDigest(missing.Styles[1].Effective)
	if err != nil {
		t.Fatal(err)
	}
	if issues := ValidateNativeWorkbookV1(missing); !nativeIssuesContainPath(issues, "/styles/1/effective/unsupported") {
		t.Fatalf("missing style-scoped diagnostic was accepted: %#v", issues)
	}

	extra := decode()
	part := "XL/Styles.xml"
	location := strings.Join([]string{"STYLE_BORDER", "styles", "style:1", part, "", ""}, "\x00")
	digest := sha256.Sum256([]byte(location))
	extra.Unsupported = append(extra.Unsupported, NativeWorkbookUnsupportedV1{
		ID: "unsupported:" + hex.EncodeToString(digest[:]), Code: "STYLE_BORDER", Capability: "styles", ScopeID: "style:1",
		PartName: &part, Preservation: "preserve-exact", Message: "contradictory style authority",
	})
	if issues := ValidateNativeWorkbookV1(extra); !nativeIssuesContainPath(issues, "/unsupported") {
		t.Fatalf("contradictory style-scoped diagnostic was accepted: %#v", issues)
	}

	nonCanonical := decode()
	part = "XL/Styles.xml"
	location = strings.Join([]string{"STYLE_BORDER", "styles", "style:+1", part, "", ""}, "\x00")
	digest = sha256.Sum256([]byte(location))
	unsupportedIndex := len(nonCanonical.Unsupported)
	nonCanonical.Unsupported = append(nonCanonical.Unsupported, NativeWorkbookUnsupportedV1{
		ID: "unsupported:" + hex.EncodeToString(digest[:]), Code: "STYLE_BORDER", Capability: "styles", ScopeID: "style:+1",
		PartName: &part, Preservation: "preserve-exact", Message: "non-canonical style source",
	})
	if issues := ValidateNativeWorkbookV1(nonCanonical); !nativeIssuesContainPath(issues, "/unsupported/"+strconv.Itoa(unsupportedIndex)+"/scope_id") {
		t.Fatalf("non-canonical style scope was accepted: %#v", issues)
	}
}

type nativeUnicodeLengthVectors struct {
	Astral                  string `json:"astral"`
	AcceptedSheetASCIICount int    `json:"accepted_sheet_ascii_count"`
	RejectedSheetASCIICount int    `json:"rejected_sheet_ascii_count"`
	MetadataScalarLimit     int    `json:"metadata_scalar_limit"`
}

func TestNativeXLSXSharedFixtureRoundTripsCanonicalGoEncoding(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v1/valid/lexical-render.json")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := DecodeNativeWorkbookV1(data)
	if err != nil {
		t.Fatalf("shared Go/TS fixture does not satisfy Go: %v", err)
	}
	encoded, err := EncodeNativeWorkbookV1(workbook)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(encoded, bytes.TrimSpace(data)) {
		t.Fatalf("shared fixture is not canonical Go encoding\n got: %s\nwant: %s", encoded, bytes.TrimSpace(data))
	}
}

func TestNativeXLSXUnicodeLengthParity(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v1/unicode-length-vectors.json")
	if err != nil {
		t.Fatal(err)
	}
	var vectors nativeUnicodeLengthVectors
	if err := json.Unmarshal(data, &vectors); err != nil {
		t.Fatal(err)
	}
	fixtureData, err := os.ReadFile("testdata/native-xlsx-v1/valid/lexical-render.json")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := DecodeNativeWorkbookV1(fixtureData)
	if err != nil {
		t.Fatal(err)
	}
	workbook.Sheets[0].Name = strings.Repeat("A", vectors.AcceptedSheetASCIICount) + vectors.Astral
	detail := strings.Repeat("A", vectors.MetadataScalarLimit-1) + vectors.Astral
	workbook.Capabilities[0].Detail = &detail
	if issues := ValidateNativeWorkbookV1(workbook); len(issues) != 0 {
		t.Fatalf("accepted astral length vector was rejected: %#v", issues)
	}

	workbook.Sheets[0].Name = strings.Repeat("A", vectors.RejectedSheetASCIICount) + vectors.Astral
	if issues := ValidateNativeWorkbookV1(workbook); !nativeIssuesContainPath(issues, "/sheets/0/name") {
		t.Fatalf("OOXML UTF-16 overflow was not rejected: %#v", issues)
	}
	workbook.Sheets[0].Name = strings.Repeat("A", vectors.AcceptedSheetASCIICount) + vectors.Astral
	detail = strings.Repeat("A", vectors.MetadataScalarLimit) + vectors.Astral
	workbook.Capabilities[0].Detail = &detail
	if issues := ValidateNativeWorkbookV1(workbook); !nativeIssuesContainPath(issues, "/capabilities/0/detail") {
		t.Fatalf("JSON Schema scalar overflow was not rejected: %#v", issues)
	}
}

func TestNativeXLSXRejectsDuplicateAndBoundaryApostropheSheetNames(t *testing.T) {
	data, err := os.ReadFile("testdata/native-xlsx-v1/valid/lexical-render.json")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := DecodeNativeWorkbookV1(data)
	if err != nil {
		t.Fatal(err)
	}
	for _, pair := range [][2]string{{"Lexical", "lEXICAL"}, {"Σ", "ς"}, {"İ", "i"}} {
		workbook.Sheets[0].Name = pair[0]
		workbook.Sheets = append(workbook.Sheets, NativeWorkbookSheetV1{
			ID: "8", Name: pair[1], Order: 1, State: "visible", PartName: "Worksheets/Sheet2.xml",
			Rows: []NativeWorkbookRowDimensionV1{}, Columns: []NativeWorkbookColumnDimensionV1{}, Cells: []NativeWorkbookCellV1{}, MergedRanges: []NativeWorkbookMergedRangeV1{}, Editable: true,
		})
		if issues := ValidateNativeWorkbookV1(workbook); !nativeIssuesContainPath(issues, "/sheets/1/name") {
			t.Fatalf("case-equivalent sheet names %q/%q were not rejected: %#v", pair[0], pair[1], issues)
		}
		workbook.Sheets = workbook.Sheets[:1]
	}
	workbook.Sheets[0].Name = "Lexical"
	for _, name := range []string{"'Lexical", "Lexical'"} {
		workbook.Sheets[0].Name = name
		if issues := ValidateNativeWorkbookV1(workbook); !nativeIssuesContainPath(issues, "/sheets/0/name") {
			t.Fatalf("boundary apostrophe sheet name %q was not rejected: %#v", name, issues)
		}
	}
}

func nativeIssuesContainPath(issues []NativeWorkbookValidationIssue, path string) bool {
	for _, issue := range issues {
		if issue.Path == path {
			return true
		}
	}
	return false
}

func TestNativeXLSXExtractionCanonicalizesFormulaGroupRanges(t *testing.T) {
	parts := nativeWorkbookFixture(false)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `ref="H1:H2"`, `ref="$H$1:$H$2"`, 1)
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	master := nativeCellPointerForTest(t, workbook, "7", "H1")
	if master.Formula == nil || master.Formula.Ref == nil || *master.Formula.Ref != "H1:H2" {
		t.Fatalf("canonical formula range = %#v", master.Formula)
	}
	item := nativeUnsupportedPointerForTest(t, workbook, "FORMULA_GROUP_RANGE")
	if item.RangeRef == nil || *item.RangeRef != "H1:H2" {
		t.Fatalf("canonical unsupported range = %#v", item.RangeRef)
	}
	if _, err := EncodeNativeWorkbookV1(workbook); err != nil {
		t.Fatalf("canonicalized extraction does not validate: %v", err)
	}
}

func TestNativeXLSXSchemaBindingsMatchGoWireTypes(t *testing.T) {
	types := map[string]reflect.Type{
		"NativeWorkbookV1":                reflect.TypeOf(NativeWorkbookV1{}),
		"NativeWorkbookSourceV1":          reflect.TypeOf(NativeWorkbookSourceV1{}),
		"NativeWorkbookSheetV1":           reflect.TypeOf(NativeWorkbookSheetV1{}),
		"NativeWorkbookCellV1":            reflect.TypeOf(NativeWorkbookCellV1{}),
		"NativeWorkbookValueV1":           reflect.TypeOf(NativeWorkbookValueV1{}),
		"NativeWorkbookFormulaV1":         reflect.TypeOf(NativeWorkbookFormulaV1{}),
		"NativeWorkbookMergedRangeV1":     reflect.TypeOf(NativeWorkbookMergedRangeV1{}),
		"NativeWorkbookRowDimensionV1":    reflect.TypeOf(NativeWorkbookRowDimensionV1{}),
		"NativeWorkbookSheetFormatV1":     reflect.TypeOf(NativeWorkbookSheetFormatV1{}),
		"NativeWorkbookNormalStyleV1":     reflect.TypeOf(NativeWorkbookNormalStyleV1{}),
		"NativeWorkbookColumnDimensionV1": reflect.TypeOf(NativeWorkbookColumnDimensionV1{}),
		"NativeWorkbookStyleV1":           reflect.TypeOf(NativeWorkbookStyleV1{}),
		"NativeWorkbookEffectiveStyleV1":  reflect.TypeOf(NativeWorkbookEffectiveStyleV1{}),
		"NativeWorkbookFillV1":            reflect.TypeOf(NativeWorkbookFillV1{}),
		"NativeWorkbookBorderV1":          reflect.TypeOf(NativeWorkbookBorderV1{}),
		"NativeWorkbookBorderSideV1":      reflect.TypeOf(NativeWorkbookBorderSideV1{}),
		"NativeWorkbookCapabilityV1":      reflect.TypeOf(NativeWorkbookCapabilityV1{}),
		"NativeWorkbookPassthroughPartV1": reflect.TypeOf(NativeWorkbookPassthroughPartV1{}),
		"NativeWorkbookUnsupportedV1":     reflect.TypeOf(NativeWorkbookUnsupportedV1{}),
	}
	if len(types) != len(nativeXLSXBindingShapes) {
		t.Fatalf("Go/schema object binding count differs: Go=%d schema=%d", len(types), len(nativeXLSXBindingShapes))
	}
	for name, shape := range nativeXLSXBindingShapes {
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

func TestNativeXLSXSchemaLimitsMatchGoContract(t *testing.T) {
	if nativeXLSXSchemaProtocol != NativeXLSXProtocol || nativeXLSXSchemaVersion != NativeXLSXVersion {
		t.Fatalf("schema protocol/version %q/%d differ from Go %q/%d", nativeXLSXSchemaProtocol, nativeXLSXSchemaVersion, NativeXLSXProtocol, NativeXLSXVersion)
	}
	checks := map[string][2]int{
		"JSON bytes":    {nativeXLSXSchemaMaxJsonBytes, NativeXLSXMaxJSONBytes},
		"JSON depth":    {nativeXLSXSchemaMaxJsonDepth, nativeWorkbookMaxJSONDepth},
		"JSON tokens":   {nativeXLSXSchemaMaxJsonTokens, nativeWorkbookMaxJSONTokens},
		"sheets":        {nativeXLSXSchemaMaxSheets, NativeXLSXMaxSheets},
		"rows":          {nativeXLSXSchemaMaxRows, NativeXLSXMaxRows},
		"columns":       {nativeXLSXSchemaMaxColumns, NativeXLSXMaxColumns},
		"cells":         {nativeXLSXSchemaMaxCells, NativeXLSXMaxCells},
		"merged ranges": {nativeXLSXSchemaMaxMergedRanges, NativeXLSXMaxMergedRanges},
		"styles":        {nativeXLSXSchemaMaxStyles, maxStyleTableRecords},
		"inventory":     {nativeXLSXSchemaMaxInventory, NativeXLSXMaxInventory},
		"text bytes":    {nativeXLSXSchemaMaxTextBytes, NativeXLSXMaxTextLength},
		"part bytes":    {nativeXLSXSchemaMaxPartBytes, NativeXLSXMaxPartBytes},
	}
	for name, values := range checks {
		if values[0] != values[1] {
			t.Errorf("%s schema limit %d differs from Go limit %d", name, values[0], values[1])
		}
	}
}

func nativeGoJSONShape(typeOf reflect.Type) ([]string, []string, map[string]string) {
	properties, required := make([]string, 0, typeOf.NumField()), make([]string, 0, typeOf.NumField())
	propertyTypes := make(map[string]string, typeOf.NumField())
	for index := 0; index < typeOf.NumField(); index++ {
		tag := typeOf.Field(index).Tag.Get("json")
		parts := strings.Split(tag, ",")
		if parts[0] == "" || parts[0] == "-" {
			continue
		}
		properties = append(properties, parts[0])
		propertyTypes[parts[0]] = nativeGoWireType(typeOf.Field(index).Type)
		if len(parts) == 1 || parts[1] != "omitempty" {
			required = append(required, parts[0])
		}
	}
	sort.Strings(properties)
	sort.Strings(required)
	return properties, required, propertyTypes
}

func nativeGoWireType(typeOf reflect.Type) string {
	for typeOf.Kind() == reflect.Pointer {
		typeOf = typeOf.Elem()
	}
	if typeOf.Kind() == reflect.Slice {
		return "[]" + nativeGoWireType(typeOf.Elem())
	}
	if typeOf.Kind() == reflect.Struct {
		return typeOf.Name()
	}
	switch typeOf.Kind() {
	case reflect.String:
		return "string"
	case reflect.Bool:
		return "boolean"
	case reflect.Int, reflect.Int8, reflect.Int16, reflect.Int32, reflect.Int64,
		reflect.Uint, reflect.Uint8, reflect.Uint16, reflect.Uint32, reflect.Uint64:
		return "integer"
	case reflect.Float32, reflect.Float64:
		return "number"
	default:
		return typeOf.String()
	}
}
