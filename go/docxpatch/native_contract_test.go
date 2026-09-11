package docxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"os"
	"reflect"
	"sort"
	"strconv"
	"strings"
	"testing"
)

func nativeFixture(t *testing.T) []byte {
	t.Helper()
	data, err := os.ReadFile("../../testdata/docx-native/document-v1.json")
	if err != nil {
		t.Fatal(err)
	}
	return data
}

type nativeInvalidFixture struct {
	Cases []struct {
		Name    string `json:"name"`
		Pointer string `json:"pointer"`
		Value   any    `json:"value"`
		Code    string `json:"code"`
		Path    string `json:"path"`
	} `json:"cases"`
}

func setNativePointer(root any, pointer string, value any) error {
	segments := strings.Split(strings.TrimPrefix(pointer, "/"), "/")
	current := root
	for index, escaped := range segments {
		segment := strings.ReplaceAll(strings.ReplaceAll(escaped, "~1", "/"), "~0", "~")
		last := index == len(segments)-1
		switch owner := current.(type) {
		case map[string]any:
			if last {
				owner[segment] = value
				return nil
			}
			current = owner[segment]
		case []any:
			position, err := strconv.Atoi(segment)
			if err != nil || position < 0 || position >= len(owner) {
				return fmt.Errorf("invalid array segment %q", segment)
			}
			if last {
				owner[position] = value
				return nil
			}
			current = owner[position]
		default:
			return fmt.Errorf("pointer %q does not resolve", pointer)
		}
	}
	return fmt.Errorf("pointer %q is empty", pointer)
}

func TestValidateNativeDocumentV1ChargesCollectionsOnce(t *testing.T) {
	doc, err := DecodeNativeDocumentV1(nativeFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	doc.Unsupported = make([]NativeUnsupportedCapabilityV1, 1000)
	for i := range doc.Unsupported {
		doc.Unsupported[i] = NativeUnsupportedCapabilityV1{ID: fmt.Sprintf("unsupported:test:%d", i), Code: "PRESERVED", Capability: "metadata", ScopeID: doc.Body.ID, Preservation: "refuse-mutation", Message: "Unmodeled metadata is preserved."}
	}
	if _, err := EncodeNativeDocumentV1(doc); err != nil {
		t.Fatalf("bounded collection was charged repeatedly: %v", err)
	}
	doc.Unsupported[999].Preservation = "invalid"
	issues := ValidateNativeDocumentV1(doc)
	if len(issues) != 1 || issues[0].Path != "/unsupported/999/preservation" {
		t.Fatalf("validator did not reach final entry: %#v", issues)
	}
}

func TestNativeValidatorCollectionBudgetStillEnforced(t *testing.T) {
	v := &nativeValidator{}
	for i := 0; i < NativeDOCXMaxNodes/NativeDOCXMaxCollectionItems; i++ {
		if got := v.collection(NativeDOCXMaxCollectionItems, "/items", NativeDOCXMaxCollectionItems); got != NativeDOCXMaxCollectionItems {
			t.Fatalf("premature limit: %d", got)
		}
	}
	if got := v.collection(1, "/overflow", NativeDOCXMaxCollectionItems); got != 0 || len(v.issues) != 1 || v.issues[0].Code != "LIMIT_EXCEEDED" {
		t.Fatalf("budget no longer enforced: %#v", v)
	}
	v = &nativeValidator{}
	if got := v.collection(NativeDOCXMaxCollectionItems+1, "/too-many", NativeDOCXMaxCollectionItems); got != NativeDOCXMaxCollectionItems || len(v.issues) != 1 {
		t.Fatal("collection bound no longer enforced")
	}
}

func TestDecodeNativeDocumentV1CrossLanguageFixture(t *testing.T) {
	doc, err := DecodeNativeDocumentV1(nativeFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	if doc.Protocol != NativeDOCXProtocol || doc.Version == nil || *doc.Version != 1 {
		t.Fatalf("unexpected envelope: %#v", doc)
	}
	if got := len(doc.Body.Blocks); got != 2 {
		t.Fatalf("body blocks = %d, want 2", got)
	}
	if got := doc.Sections[0].HeaderRefs[0].StoryID; got != "story:header:default" {
		t.Fatalf("header story = %q", got)
	}
	if got := doc.CommentStories[0].Blocks[0].Kind; got != "paragraph" {
		t.Fatalf("comment body block = %q", got)
	}
}

func TestValidateNativeDocumentV1ClassifiesMediaMIMEWithASCIICaseFolding(t *testing.T) {
	doc, err := DecodeNativeDocumentV1(nativeFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	for index := range doc.PassthroughParts {
		if doc.PassthroughParts[index].PartName == "word/media/image1.png" {
			doc.PassthroughParts[index].ContentType = "IMAGE/PNG"
		}
	}
	if issues := ValidateNativeDocumentV1(doc); len(issues) != 0 {
		t.Fatalf("ASCII MIME case variant must retain media-part identity: %#v", issues)
	}
}

func TestValidateNativeDocumentV1BoundsEveryTwipBeforeMilliPointConversion(t *testing.T) {
	doc, err := DecodeNativeDocumentV1(nativeFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	table := doc.Body.Blocks[1].Table
	if table == nil {
		t.Fatal("cross-language fixture is missing its table")
	}
	maximum := nativeMaxTwipsForMilliPoints
	table.WidthTwips = nativeInt64(maximum)
	table.GridWidthsTwips = []int64{maximum}
	table.CellMargins = &NativeTableCellMarginsV1{}
	table.CellMargins.LeftTwips = maximum
	doc.Sections[0].Page.WidthTwips = nativeInt64(maximum)
	if issues := ValidateNativeDocumentV1(doc); len(issues) != 0 {
		t.Fatalf("exact maximum x50-convertible twips were rejected: %#v", issues)
	}

	for name, mutate := range map[string]func(){
		"table width":       func() { table.WidthTwips = nativeInt64(maximum + 1) },
		"table grid":        func() { table.GridWidthsTwips[0] = maximum + 1 },
		"table cell margin": func() { table.CellMargins.LeftTwips = maximum + 1 },
		"page width":        func() { doc.Sections[0].Page.WidthTwips = nativeInt64(maximum + 1) },
	} {
		t.Run(name, func(t *testing.T) {
			fresh, err := DecodeNativeDocumentV1(nativeFixture(t))
			if err != nil {
				t.Fatal(err)
			}
			doc = fresh
			table = doc.Body.Blocks[1].Table
			table.GridWidthsTwips = []int64{1}
			table.CellMargins = &NativeTableCellMarginsV1{}
			mutate()
			if issues := ValidateNativeDocumentV1(doc); len(issues) == 0 {
				t.Fatal("twips beyond the exact x50 boundary were accepted")
			}
		})
	}
}

func TestEncodeNativeDocumentV1IsDeterministic(t *testing.T) {
	doc, err := DecodeNativeDocumentV1(nativeFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	one, err := EncodeNativeDocumentV1(doc)
	if err != nil {
		t.Fatal(err)
	}
	two, err := EncodeNativeDocumentV1(doc)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(one, two) {
		t.Fatal("canonical encoding changed between calls")
	}
	if !bytes.HasPrefix(one, []byte(`{"body":`)) {
		t.Fatalf("object keys are not canonical: %.40s", one)
	}
	if bytes.Index(one, []byte(`"paragraph:intro"`)) > bytes.Index(one, []byte(`"table:summary"`)) {
		t.Fatal("source-order block array was reordered")
	}
}

func TestDecodeNativeDocumentV1RejectsUnknownFields(t *testing.T) {
	var value map[string]any
	if err := json.Unmarshal(nativeFixture(t), &value); err != nil {
		t.Fatal(err)
	}
	value["html"] = "not canonical"
	data, err := json.Marshal(value)
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativeDocumentV1(data)
	if err == nil || !strings.Contains(err.Error(), `unknown field "html"`) {
		t.Fatalf("expected strict unknown-field error, got %v", err)
	}
}

func TestDecodeNativeDocumentV1RejectsUnsafeModel(t *testing.T) {
	var value map[string]any
	if err := json.Unmarshal(nativeFixture(t), &value); err != nil {
		t.Fatal(err)
	}
	body := value["body"].(map[string]any)
	blocks := body["blocks"].([]any)
	paragraph := blocks[0].(map[string]any)["paragraph"].(map[string]any)
	runs := paragraph["runs"].([]any)
	runs[0].(map[string]any)["control"] = "tab"
	sections := value["sections"].([]any)
	sections[0].(map[string]any)["starts_at_block_id"] = "paragraph:missing"
	table := blocks[1].(map[string]any)["table"].(map[string]any)
	delete(table["edit_policy"].(map[string]any), "refusal")
	data, err := json.Marshal(value)
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativeDocumentV1(data)
	validation, ok := err.(*NativeValidationError)
	if !ok {
		t.Fatalf("expected NativeValidationError, got %T: %v", err, err)
	}
	codes := map[string]bool{}
	for _, issue := range validation.Issues {
		codes[issue.Code] = true
	}
	for _, code := range []string{"INVALID_UNION", "BROKEN_REFERENCE", "REQUIRED"} {
		if !codes[code] {
			t.Errorf("missing issue code %s in %#v", code, validation.Issues)
		}
	}
}

func TestNativeDOCXGoBindingsMatchCanonicalSchema(t *testing.T) {
	schemaBytes, err := os.ReadFile("../../packages/docs/schema/native-docx-v1.schema.json")
	if err != nil {
		t.Fatal(err)
	}
	var schema struct {
		Definitions map[string]struct {
			Properties map[string]any `json:"properties"`
		} `json:"$defs"`
	}
	if err := json.Unmarshal(schemaBytes, &schema); err != nil {
		t.Fatal(err)
	}
	bindings := map[string]reflect.Type{
		"SourcePackageV1":         reflect.TypeOf(NativeSourcePackageV1{}),
		"SourceAnchorV1":          reflect.TypeOf(NativeSourceAnchorV1{}),
		"RefusalV1":               reflect.TypeOf(NativeRefusalV1{}),
		"EditPolicyV1":            reflect.TypeOf(NativeEditPolicyV1{}),
		"CapabilityV1":            reflect.TypeOf(NativeCapabilityV1{}),
		"PassthroughPartV1":       reflect.TypeOf(NativePassthroughPartV1{}),
		"RunPropertiesV1":         reflect.TypeOf(NativeRunPropertiesV1{}),
		"DrawingV1":               reflect.TypeOf(NativeDrawingV1{}),
		"ReferenceV1":             reflect.TypeOf(NativeReferenceV1{}),
		"RunV1":                   reflect.TypeOf(NativeRunV1{}),
		"NumberingReferenceV1":    reflect.TypeOf(NativeNumberingReferenceV1{}),
		"ParagraphPropertiesV1":   reflect.TypeOf(NativeParagraphPropertiesV1{}),
		"ParagraphV1":             reflect.TypeOf(NativeParagraphV1{}),
		"TableBorderV1":           reflect.TypeOf(NativeTableBorderV1{}),
		"TableBordersV1":          reflect.TypeOf(NativeTableBordersV1{}),
		"TableCellMarginsV1":      reflect.TypeOf(NativeTableCellMarginsV1{}),
		"TableCellV1":             reflect.TypeOf(NativeTableCellV1{}),
		"TableRowV1":              reflect.TypeOf(NativeTableRowV1{}),
		"TableV1":                 reflect.TypeOf(NativeTableV1{}),
		"BlockV1":                 reflect.TypeOf(NativeBlockV1{}),
		"StoryV1":                 reflect.TypeOf(NativeStoryV1{}),
		"HeaderFooterReferenceV1": reflect.TypeOf(NativeHeaderFooterReferenceV1{}),
		"PageMarginsV1":           reflect.TypeOf(NativePageMarginsV1{}),
		"ColumnV1":                reflect.TypeOf(NativeColumnV1{}),
		"PageGeometryV1":          reflect.TypeOf(NativePageGeometryV1{}),
		"SectionV1":               reflect.TypeOf(NativeSectionV1{}),
		"CommentV1":               reflect.TypeOf(NativeCommentV1{}),
		"UnsupportedCapabilityV1": reflect.TypeOf(NativeUnsupportedCapabilityV1{}),
		"DocumentV1":              reflect.TypeOf(NativeDocumentV1{}),
	}
	for name, binding := range bindings {
		definition, ok := schema.Definitions[name]
		if !ok {
			t.Errorf("canonical schema is missing %s", name)
			continue
		}
		var want []string
		for field := range definition.Properties {
			want = append(want, field)
		}
		var got []string
		for i := 0; i < binding.NumField(); i++ {
			name := strings.Split(binding.Field(i).Tag.Get("json"), ",")[0]
			if name != "" && name != "-" {
				got = append(got, name)
			}
		}
		sort.Strings(want)
		sort.Strings(got)
		if !reflect.DeepEqual(got, want) {
			t.Errorf("%s fields = %v, schema = %v", name, got, want)
		}
	}
}

func TestNativeDOCXSchemaPublishesStructuralSafety(t *testing.T) {
	schemaBytes, err := os.ReadFile("../../packages/docs/schema/native-docx-v1.schema.json")
	if err != nil {
		t.Fatal(err)
	}
	var schema map[string]any
	if err := json.Unmarshal(schemaBytes, &schema); err != nil {
		t.Fatal(err)
	}
	definitions := schema["$defs"].(map[string]any)
	document := definitions["DocumentV1"].(map[string]any)
	documentProperties := document["properties"].(map[string]any)
	commentStories := documentProperties["comment_stories"].(map[string]any)
	if got := int(commentStories["maxItems"].(float64)); got != NativeDOCXMaxCollectionItems {
		t.Fatalf("comment_stories maxItems = %d", got)
	}
	run := definitions["RunV1"].(map[string]any)
	runProperties := run["properties"].(map[string]any)
	if got := int(runProperties["text"].(map[string]any)["maxLength"].(float64)); got != NativeDOCXMaxTextLength {
		t.Fatalf("run text maxLength = %d", got)
	}
	for index, branch := range run["oneOf"].([]any) {
		if _, ok := branch.(map[string]any)["not"]; !ok {
			t.Fatalf("RunV1 oneOf branch %d is not exclusive", index)
		}
	}
	partPattern := definitions["PartName"].(map[string]any)["pattern"].(string)
	for _, token := range []string{"2E", "2F", "5C"} {
		if !strings.Contains(partPattern, token) {
			t.Fatalf("PartName pattern does not reject encoded %s", token)
		}
	}
	if !strings.Contains(partPattern, `\.(?:/|$)`) {
		t.Fatal("PartName pattern does not reject trailing-dot segments")
	}
}

func TestDecodeNativeDocumentV1SharedInvalidVectors(t *testing.T) {
	data, err := os.ReadFile("../../testdata/docx-native/invalid-v1.json")
	if err != nil {
		t.Fatal(err)
	}
	var fixture nativeInvalidFixture
	if err := json.Unmarshal(data, &fixture); err != nil {
		t.Fatal(err)
	}
	for _, testCase := range fixture.Cases {
		t.Run(testCase.Name, func(t *testing.T) {
			var candidate any
			if err := json.Unmarshal(nativeFixture(t), &candidate); err != nil {
				t.Fatal(err)
			}
			if err := setNativePointer(candidate, testCase.Pointer, testCase.Value); err != nil {
				t.Fatal(err)
			}
			payload, err := json.Marshal(candidate)
			if err != nil {
				t.Fatal(err)
			}
			_, err = DecodeNativeDocumentV1(payload)
			validation, ok := err.(*NativeValidationError)
			if !ok {
				t.Fatalf("expected NativeValidationError, got %T: %v", err, err)
			}
			matched := false
			for _, issue := range validation.Issues {
				if issue.Code == testCase.Code && issue.Path == testCase.Path {
					matched = true
					break
				}
			}
			if !matched {
				t.Fatalf("missing %s at %s in %#v", testCase.Code, testCase.Path, validation.Issues)
			}
		})
	}
}

func TestDecodeNativeDocumentV1BoundsWorkAndIssues(t *testing.T) {
	tooLarge := bytes.Repeat([]byte{' '}, NativeDOCXMaxJSONBytes+1)
	if _, err := DecodeNativeDocumentV1(tooLarge); err == nil || !strings.Contains(err.Error(), "exceeds") {
		t.Fatalf("expected payload limit, got %v", err)
	}

	var candidate map[string]any
	if err := json.Unmarshal(nativeFixture(t), &candidate); err != nil {
		t.Fatal(err)
	}
	body := candidate["body"].(map[string]any)
	paragraph := body["blocks"].([]any)[0].(map[string]any)["paragraph"].(map[string]any)
	paragraph["edit_policy"].(map[string]any)["allowed_operations"] = []any{
		"text.replace", "properties.patch", "block.insert_after", "block.delete", "drawing.replace", "text.replace",
	}
	payload, err := json.Marshal(candidate)
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativeDocumentV1(payload)
	validation, ok := err.(*NativeValidationError)
	if !ok {
		t.Fatalf("expected NativeValidationError, got %T: %v", err, err)
	}
	if !hasNativeIssue(validation.Issues, "LIMIT_EXCEEDED", "/body/blocks/0/paragraph/edit_policy/allowed_operations") {
		t.Fatalf("missing operation bound in %#v", validation.Issues)
	}

	paragraph["runs"].([]any)[0].(map[string]any)["text"] = strings.Repeat("x", NativeDOCXMaxTextLength+1)
	payload, err = json.Marshal(candidate)
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativeDocumentV1(payload)
	validation, ok = err.(*NativeValidationError)
	if !ok || !hasNativeIssue(validation.Issues, "LIMIT_EXCEEDED", "/body/blocks/0/paragraph/runs/0/text") {
		t.Fatalf("expected text bound, got %T: %v", err, err)
	}

	nulls := map[string]any{}
	for index := 0; index < 200; index++ {
		nulls[fmt.Sprintf("null_%d", index)] = nil
	}
	payload, err = json.Marshal(nulls)
	if err != nil {
		t.Fatal(err)
	}
	_, err = DecodeNativeDocumentV1(payload)
	validation, ok = err.(*NativeValidationError)
	if !ok {
		t.Fatalf("expected NativeValidationError, got %T: %v", err, err)
	}
	if len(validation.Issues) != NativeDOCXMaxIssues {
		t.Fatalf("issue count = %d, want %d", len(validation.Issues), NativeDOCXMaxIssues)
	}
}

func TestDecodeNativeDocumentV1RejectsNegativeZero(t *testing.T) {
	raw := bytes.Replace(nativeFixture(t), []byte(`"start_byte": 100,`), []byte(`"start_byte": -0,`), 1)
	_, err := DecodeNativeDocumentV1(raw)
	validation, ok := err.(*NativeValidationError)
	if !ok || !hasNativeIssue(validation.Issues, "INVALID_VALUE", "/body/anchor/start_byte") {
		t.Fatalf("expected negative-zero refusal, got %T: %v", err, err)
	}
}

func hasNativeIssue(issues []NativeValidationIssue, code, path string) bool {
	for _, issue := range issues {
		if issue.Code == code && issue.Path == path {
			return true
		}
	}
	return false
}
