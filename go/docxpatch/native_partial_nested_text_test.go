package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
)

func nestedTextSource(t *testing.T, change func(map[string]string)) []byte {
	styles := `<w:styles xmlns:w="` + testW + `"><w:style w:type="table" w:styleId="Grid"><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="4" w:color="auto"/></w:tblBorders></w:tblPr></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	props := `<w:tblPr><w:tblStyle w:val="Grid"/><w:tblW w:w="1417" w:type="dxa"/><w:tblLook w:val="04A0" w:firstRow="1" w:firstColumn="1" w:noVBand="1"/></w:tblPr><w:tblGrid><w:gridCol w:w="1417"/></w:tblGrid>`
	parts["word/document.xml"] = nativeMutationMain(`<w:tbl>` + props + `<w:tr><w:tc><w:p><w:r><w:t>Outer</w:t></w:r></w:p><w:tbl>` + props + `<w:tr><w:tc><w:p><w:r><w:t>Inner visible</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`)
	parts["word/document.xml"] = strings.Replace(parts["word/document.xml"], `</w:body>`, `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body>`, 1)
	if change != nil {
		change(parts)
	}
	return buildNativeDOCX(t, nativeEntries(parts))
}
func TestPartialNestedTextSameSourceResolution(t *testing.T) {
	source := nestedTextSource(t, nil)
	doc, e := ExtractNativeDocumentV1(source)
	if e != nil {
		t.Fatal(e)
	}
	encoded, e := InspectNativePartialSourceV1(source)
	if e != nil {
		t.Fatal(e)
	}
	var out struct {
		Items    []NativePartialNestedTextV1 `json:"nested_text"`
		Document NativeDocumentV1            `json:"document"`
	}
	if e = json.Unmarshal(encoded, &out); e != nil {
		t.Fatal(e)
	}
	if len(out.Items) != 1 || out.Items[0].Paragraphs[0].Runs[0].Text != "Inner visible" || len(out.Items[0].ResolvedLayout.Runs) != 1 {
		t.Fatalf("nested text missing: %s", encoded)
	}
	a, _ := json.Marshal(doc)
	b, _ := json.Marshal(out.Document)
	if string(a) != string(b) {
		t.Fatal("original model changed")
	}
	if strings.Contains(string(a), "Inner visible") {
		t.Fatal("nested content entered strict document")
	}
}
func TestPartialNestedTextRejectsUnsupportedOrHiddenSources(t *testing.T) {
	for name, change := range map[string]func(map[string]string){
		"hidden-default": func(p map[string]string) {
			p["word/styles.xml"] = strings.Replace(p["word/styles.xml"], `<w:style `, `<w:docDefaults><w:rPrDefault><w:rPr><w:vanish/></w:rPr></w:rPrDefault></w:docDefaults><w:style `, 1)
		},
		"hidden-style": func(p map[string]string) {
			p["word/styles.xml"] = strings.Replace(p["word/styles.xml"], `</w:style>`, `<w:rPr><w:vanish/></w:rPr></w:style>`, 1)
		},
		"conditional-style": func(p map[string]string) {
			p["word/styles.xml"] = strings.Replace(p["word/styles.xml"], `</w:style>`, `<w:tblStylePr w:type="firstRow"/></w:style>`, 1)
		},
		"missing-style": func(p map[string]string) {
			p["word/styles.xml"] = strings.ReplaceAll(p["word/styles.xml"], `w:styleId="Grid"`, `w:styleId="Missing"`)
		},
		"revision": func(p map[string]string) {
			p["word/document.xml"] = strings.Replace(p["word/document.xml"], `<w:r><w:t>Inner visible</w:t></w:r>`, `<w:ins w:id="1"><w:r><w:t>Inner visible</w:t></w:r></w:ins>`, 1)
		},
		"direct-hidden": func(p map[string]string) {
			p["word/document.xml"] = strings.Replace(p["word/document.xml"], `<w:r><w:t>Inner visible`, `<w:r><w:rPr><w:vanish/></w:rPr><w:t>Inner visible`, 1)
		},
		"paragraph-property": func(p map[string]string) {
			p["word/document.xml"] = strings.Replace(p["word/document.xml"], `<w:p><w:r><w:t>Inner visible`, `<w:p><w:pPr><w:pStyle w:val="Missing"/></w:pPr><w:r><w:t>Inner visible`, 1)
		},
		"deeper-table": func(p map[string]string) {
			p["word/document.xml"] = strings.Replace(p["word/document.xml"], `<w:p><w:r><w:t>Inner visible</w:t></w:r></w:p>`, `<w:tbl><w:tr><w:tc><w:p/></w:tc></w:tr></w:tbl>`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) {
			encoded, e := InspectNativePartialSourceV1(nestedTextSource(t, change))
			if e != nil {
				t.Fatal(e)
			}
			var out map[string]json.RawMessage
			json.Unmarshal(encoded, &out)
			if _, ok := out["nested_text"]; ok {
				t.Fatal("unsupported nested text admitted")
			}
		})
	}
}
