package docxpatch

import (
	"bytes"
	"encoding/json"
	"reflect"
	"strings"
	"testing"
)

func TestPartialNestedTableContainment(t *testing.T) {
	nested := `<w:tbl><w:tr><w:tc><w:p><w:r><w:t>INNER_SECRET</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`
	for _, middle := range []string{nested, `<w:sdt><w:sdtContent>` + nested + `</w:sdtContent></w:sdt>`, strings.Repeat(nested, 66)} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Before</w:t></w:r></w:p>`+middle+`<w:p><w:r><w:t>After</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`))))
		before := append([]byte(nil), source...)
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		p := doc.Body.Blocks[0].Table.Rows[0].Cells[0].Paragraphs[0]
		mutation := []NativeDOCXTextMutationV1{{TargetKind: "paragraph", TargetID: p.ID, ExpectedXMLSHA256: p.Anchor.XMLSHA256, Text: "Changed"}}
		beforeMutation, beforeErr := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
		encoded, err := InspectNativePartialSourceV1(source)
		if err != nil {
			t.Fatal(err)
		}
		var envelope struct {
			Nested *NativePartialNestedTablesV1 `json:"nested_table_omissions"`
		}
		if err = json.Unmarshal(encoded, &envelope); err != nil {
			t.Fatal(err)
		}
		table := doc.Body.Blocks[0].Table
		if table.EditPolicy.Mode != "read-only" || len(table.Rows[0].Cells[0].Paragraphs) != 2 || !bytes.Equal(source, before) {
			t.Fatal("strict ownership or source changed")
		}
		afterMutation, afterErr := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
		if (beforeErr == nil) != (afterErr == nil) || !reflect.DeepEqual(beforeMutation, afterMutation) {
			t.Fatal("inspection changed existing scoped mutation authority")
		}
		if strings.HasPrefix(middle, "<w:sdt>") {
			if envelope.Nested != nil {
				t.Fatal("unknown wrapper qualified")
			}
			continue
		}
		expected := 1
		extra := 0
		if len(middle) > len(nested) {
			expected = 64
			extra = 2
		}
		if envelope.Nested == nil || len(envelope.Nested.Items) != expected || envelope.Nested.OmittedCount != extra {
			t.Fatal("wrong bounded omission inventory")
		}
		fact := envelope.Nested.Items[0]
		if fact.PackageSHA256 != doc.Source.PackageSHA256 || fact.TableID != table.ID || fact.CellID != table.Rows[0].Cells[0].ID || fact.Anchor.Path != table.Rows[0].Cells[0].Anchor.Path+"/w:tbl[1]" {
			t.Fatal("unjoined fact")
		}
		nestedDiagnostics := 0
		for _, d := range doc.Unsupported {
			if d.Code == "NESTED_TABLE_OR_CELL_MARKUP" {
				nestedDiagnostics++
			}
		}
		if nestedDiagnostics != expected+extra {
			t.Fatal("source diagnostics removed")
		}
	}
}

func TestPartialNestedTableUnknownContainerNotQualified(t *testing.T) {
	nested := `<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`
	for _, cell := range []string{`<w:tc unknown="1"><w:p/>` + nested + `</w:tc>`, `<w:tc><w:tcPr/><w:tcPr/><w:p/>` + nested + `</w:tc>`, `<w:tc><w:p/><w:sdt>` + nested + `</w:sdt></w:tc>`} {
		data := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:tbl><w:tr>`+cell+`</w:tr></w:tbl>`))))
		doc, err := ExtractNativeDocumentV1(data)
		if err != nil {
			t.Fatal(err)
		}
		result, err := inspectNativePartialNestedTables(data, doc)
		if err != nil {
			t.Fatal(err)
		}
		if result != nil {
			t.Fatal("unknown containment qualified")
		}
	}
}
