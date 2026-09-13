package docxpatch

import (
	"bytes"
	"encoding/json"
	"reflect"
	"strings"
	"testing"
)

func partialTableContextSource(t *testing.T, styleExtra, look string) []byte {
	t.Helper()
	reference := "Base"
	wrongBaseType := styleExtra == "wrong-base-type"
	duplicateStyle := styleExtra == "duplicate-style"
	if wrongBaseType || duplicateStyle {
		styleExtra = ""
	}
	if styleExtra == "missing-base" {
		reference = "Missing"
		styleExtra = ""
	}
	if styleExtra == "cyclic-base" {
		reference = "Grid"
		styleExtra = ""
	}
	styles := `<w:styles xmlns:w="` + testW + `"><w:style w:type="table" w:styleId="Base"><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="Grid"><w:basedOn w:val="Base"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr>` + styleExtra + `</w:style></w:styles>`
	if wrongBaseType {
		styles = strings.Replace(styles, `w:type="table" w:styleId="Base"`, `w:type="paragraph" w:styleId="Base"`, 1)
	}
	if duplicateStyle {
		styles = strings.Replace(styles, `</w:styles>`, `<w:style w:type="table" w:styleId="Base"><w:rPr><w:vanish/></w:rPr></w:style></w:styles>`, 1)
	}
	styles = strings.Replace(styles, `<w:basedOn w:val="Base"/>`, `<w:basedOn w:val="`+reference+`"/>`, 1)
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = nativeMutationMain(`<w:tbl><w:tblPr><w:tblStyle w:val="Grid"/>` + look + `</w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p><w:r><w:t>Outer</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>INNER_SECRET</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`)
	return buildNativeDOCX(t, nativeEntries(parts))
}
func TestPartialTableTextContextPreservesSourceAndMutationAuthority(t *testing.T) {
	source := partialTableContextSource(t, "", `<w:tblLook w:val="04A0" w:firstRow="1" w:firstColumn="1" w:noVBand="1"/>`)
	before := append([]byte(nil), source...)
	doc, e := ExtractNativeDocumentV1(source)
	if e != nil {
		t.Fatal(e)
	}
	p := doc.Body.Blocks[0].Table.Rows[0].Cells[0].Paragraphs[0]
	mutation := []NativeDOCXTextMutationV1{{TargetKind: "paragraph", TargetID: p.ID, ExpectedXMLSHA256: p.Anchor.XMLSHA256, Text: "Changed"}}
	beforeMutation, beforeErr := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	encoded, e := InspectNativePartialSourceV1(source)
	if e != nil {
		t.Fatal(e)
	}
	var out struct {
		Contexts []NativePartialTableTextContextV1 `json:"table_text_contexts"`
		Nested   *NativePartialNestedTablesV1      `json:"nested_table_omissions"`
		Document NativeDocumentV1                  `json:"document"`
	}
	if e = json.Unmarshal(encoded, &out); e != nil {
		t.Fatal(e)
	}
	if len(out.Contexts) != 1 || out.Nested == nil || len(out.Nested.Items) != 1 {
		t.Fatalf("missing context or nested omission: %+v", out.Contexts)
	}
	context := out.Contexts[0]
	if context.PackageSHA256 != doc.Source.PackageSHA256 || context.TableID != doc.Body.Blocks[0].Table.ID || len(context.StyleChain) != 2 || context.StyleChain[0].StyleID != "Base" || context.StyleChain[1].StyleID != "Grid" || len(context.ResolvedDiagnostics) != 1 {
		t.Fatalf("wrong context: %+v", context)
	}
	if !bytes.Equal(source, before) || !reflect.DeepEqual(out.Document.Unsupported, doc.Unsupported) {
		t.Fatal("source diagnostics changed")
	}
	afterMutation, afterErr := ApplyNativeTextMutationsV1(source, doc.Source.PackageSHA256, mutation)
	if (beforeErr == nil) != (afterErr == nil) || !reflect.DeepEqual(beforeMutation, afterMutation) {
		t.Fatal("mutation authority changed")
	}
	metadata, _ := json.Marshal(context)
	if strings.Contains(string(metadata), "INNER_SECRET") {
		t.Fatal("nested text exposed")
	}
}
func TestPartialTableTextContextRejectsVisibilityConditionalAndUnknownStyles(t *testing.T) {
	for _, extra := range []string{"missing-base", "cyclic-base", "wrong-base-type", "duplicate-style", `<w:rPr><w:vanish/></w:rPr>`, `<w:tblStylePr w:type="firstRow"><w:rPr><w:vanish/></w:rPr></w:tblStylePr>`, `<w:trPr/>`, `<w:tcPr/>`, `<w:unknown/>`, `<w:pPr/>`, `<w:basedOn w:val="Missing"/>`, `<w:basedOn w:val="Grid"/>`} {
		source := partialTableContextSource(t, extra, `<w:tblLook w:val="04A0"/>`)
		encoded, e := InspectNativePartialSourceV1(source)
		if e != nil {
			continue
		}
		var out struct {
			Contexts []NativePartialTableTextContextV1 `json:"table_text_contexts"`
		}
		if e = json.Unmarshal(encoded, &out); e != nil {
			t.Fatal(e)
		}
		if len(out.Contexts) != 0 {
			t.Fatalf("style qualified: %s", extra)
		}
	}
	for _, look := range []string{`<w:tblLook w:val="04A0" w:firstRow="0"/>`, `<w:tblLook w:val="FFFF"/>`, `<w:tblLook w:val="04A0" unknown="1"/>`, `<w:tblLook w:val="04A0">TEXT</w:tblLook>`} {
		source := partialTableContextSource(t, "", look)
		encoded, e := InspectNativePartialSourceV1(source)
		if e != nil {
			t.Fatal(e)
		}
		var out struct {
			Contexts []NativePartialTableTextContextV1 `json:"table_text_contexts"`
		}
		json.Unmarshal(encoded, &out)
		if len(out.Contexts) != 0 {
			t.Fatalf("look qualified: %s", look)
		}
	}
}
func TestPartialTableGeometryOnlyStyleMarkupIsExact(t *testing.T) {
	for _, body := range []string{`<w:pPr mystery="x"/>`, `<w:pPr>TEXT</w:pPr>`, `<w:pPr><w:spacing w:after="0" unknown="1"/></w:pPr>`, `<w:tblPr><w:tblBorders><w:top w:val="single" mystery="x"/></w:tblBorders></w:tblPr>`, `<w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa">TEXT</w:left></w:tblCellMar></w:tblPr>`, `<w:tblPr><w:tblInd w:w="-1" w:type="dxa"/></w:tblPr>`} {
		n, e := parseNativeXML("word/styles.xml", []byte(`<w:style xmlns:w="`+testW+`" w:type="table" w:styleId="Grid">`+body+`</w:style>`))
		if e != nil {
			t.Fatal(e)
		}
		if nativePartialTextTableStyle(n, testW) {
			t.Fatalf("unknown style markup qualified: %s", body)
		}
	}
}
