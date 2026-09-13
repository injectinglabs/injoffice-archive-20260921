package docxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func nativeTextboxFixture(kind, content string) string {
	if kind == "vml" {
		return `<w:pict xmlns:v="` + nativeTextboxVML + `"><v:shape id="box1" style="position:absolute;width:100pt;height:50pt"><v:textbox><w:txbxContent>` + content + `</w:txbxContent></v:textbox></v:shape></w:pict>`
	}
	return `<w:drawing xmlns:wp="` + wordDrawingTransitional + `" xmlns:a="` + drawingMLTransitional + `" xmlns:wps="` + nativeTextboxWPS + `"><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="box1"/><a:graphic><a:graphicData uri="` + nativeTextboxWPS + `"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr/><wps:txbx><w:txbxContent>` + content + `</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`
}
func TestPartialTextboxesSourceInventory(t *testing.T) {
	for _, kind := range []string{"vml", "drawingml"} {
		t.Run(kind, func(t *testing.T) {
			drawing := nativeTextboxFixture(kind, `<w:p><w:r><w:t xml:space="preserve">Hello &amp; source </w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p>`)
			source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r><w:t>before</w:t></w:r><w:r>`+drawing+`</w:r></w:p>`))))
			before := append([]byte(nil), source...)
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			encoded, err := InspectNativePartialSourceV1(source)
			if err != nil {
				t.Fatal(err)
			}
			var envelope struct {
				Textboxes *NativePartialTextboxesV1 `json:"textbox_inventory"`
			}
			if err = json.Unmarshal(encoded, &envelope); err != nil {
				t.Fatal(err)
			}
			if envelope.Textboxes == nil || len(envelope.Textboxes.Items) != 1 {
				t.Fatalf("missing textbox: %s", encoded)
			}
			item := envelope.Textboxes.Items[0]
			if item.Status != "supported" || len(item.Paragraphs) != 2 || item.Paragraphs[0] != "Hello & source " {
				t.Fatalf("unexpected item: %#v", item)
			}
			if item.PackageSHA256 != doc.Source.PackageSHA256 || item.ParagraphID != doc.Body.Blocks[0].Paragraph.ID || !bytes.Equal(before, source) {
				t.Fatal("source changed or identity lost")
			}
			if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" {
				t.Fatal("mutation authority expanded")
			}
		})
	}
}
func TestPartialTextboxUnsafeContent(t *testing.T) {
	for _, content := range []string{
		`<w:p><w:r><w:rPr><w:vanish/></w:rPr><w:t>SECRET</w:t></w:r></w:p>`,
		`<w:p><w:del><w:r><w:delText>SECRET</w:delText></w:r></w:del></w:p>`,
		`<w:p><w:r><w:instrText>SECRET</w:instrText></w:r></w:p>`,
		`<w:p><w:r><w:rPr><w:webHidden/></w:rPr><w:t>SECRET</w:t></w:r></w:p>`,
		`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>SECRET</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
		`<w:p><w:r><w:t>` + strings.Repeat("X", 4097) + `</w:t></w:r></w:p>`,
	} {
		t.Run(content[:20], func(t *testing.T) {
			source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+nativeTextboxFixture("vml", content)+`</w:r></w:p>`))))
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			out, err := inspectNativePartialTextboxes(source, doc)
			if err != nil {
				t.Fatal(err)
			}
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || len(out.Items[0].Paragraphs) != 0 {
				t.Fatalf("unsafe content admitted: %#v", out)
			}
		})
	}
}
func TestPartialTextboxInheritanceAndAmbiguity(t *testing.T) {
	drawing := nativeTextboxFixture("vml", `<w:p><w:r><w:t>SECRET</w:t></w:r></w:p>`)
	for _, run := range []string{`<w:r><w:rPr><w:vanish/></w:rPr>` + drawing + `</w:r>`, `<w:r>` + strings.Replace(drawing, `position:absolute`, `visibility:hidden`, 1) + `</w:r>`, `<w:r>` + strings.Replace(drawing, `<v:textbox>`, `<v:textbox><w:txbxContent><w:p/></w:txbxContent>`, 1) + `</w:r>`} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p>`+run+`</w:p>`))))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := inspectNativePartialTextboxes(source, doc)
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || out.Items[0].Status != "omitted" {
			t.Fatalf("unsafe shape admitted: %#v", out)
		}
	}
}
func TestPartialTextboxLimit(t *testing.T) {
	drawing := nativeTextboxFixture("vml", `<w:p><w:r><w:t>source</w:t></w:r></w:p>`)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p>`+strings.Repeat(`<w:r>`+drawing+`</w:r>`, 66)+`</w:p>`))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	out, err := inspectNativePartialTextboxes(source, doc)
	if err != nil {
		t.Fatal(err)
	}
	if out == nil || len(out.Items) != 64 || out.OmittedCount != 2 {
		t.Fatalf("bad budget: %#v", out)
	}
}

func TestPartialTextboxInheritedVisibility(t *testing.T) {
	for _, property := range []string{`<w:vanish/>`, `<w:webHidden/>`, `<w:unknownVisibility/>`} {
		parts := nativeMutationParts(nativeMutationMain(`<w:p><w:r>` + nativeTextboxFixture("vml", `<w:p><w:r><w:t>SECRET</w:t></w:r></w:p>`) + `</w:r></w:p>`))
		parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="` + `http://schemas.openxmlformats.org/package/2006/relationships` + `"><Relationship Id="rStyles" Type="` + relBaseTransitional + `styles" Target="styles.xml"/></Relationships>`
		parts["word/styles.xml"] = `<w:styles xmlns:w="` + wordMLTransitional + `"><w:docDefaults><w:rPrDefault><w:rPr>` + property + `</w:rPr></w:rPrDefault></w:docDefaults></w:styles>`
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
		source := buildNativeDOCX(t, nativeEntries(parts))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := inspectNativePartialTextboxes(source, doc)
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || out.Items[0].Status != "omitted" || len(out.Items[0].Paragraphs) != 0 {
			t.Fatalf("inherited hidden/unqualified text admitted: %#v", out)
		}
	}
}
func TestPartialTextboxAmbiguousAndSpoofedDrawing(t *testing.T) {
	ordinary := nativeTextboxFixture("drawingml", `<w:p><w:r><w:t>SECRET</w:t></w:r></w:p>`)
	for _, drawing := range []string{
		strings.Replace(ordinary, `<wps:txbx>`, `<wps:linkedTxbx id="1"/><wps:txbx>`, 1),
		strings.Replace(ordinary, `xmlns:wps="`+nativeTextboxWPS+`"`, `xmlns:wps="urn:spoof"`, 1),
		strings.Replace(ordinary, `<wps:cNvSpPr txBox="1"/>`, `<wps:cNvPr hidden="1"/>`, 1),
	} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+drawing+`</w:r></w:p>`))))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := inspectNativePartialTextboxes(source, doc)
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || out.Items[0].Status != "omitted" {
			t.Fatalf("unsafe drawing admitted: %#v", out)
		}
	}
	mc := `<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps">` + ordinary + `</mc:Choice><mc:Fallback>` + nativeTextboxFixture("vml", `<w:p><w:r><w:t>SECRET</w:t></w:r></w:p>`) + `</mc:Fallback></mc:AlternateContent>`
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+mc+`</w:r></w:p>`))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	out, err := inspectNativePartialTextboxes(source, doc)
	if err != nil {
		t.Fatal(err)
	}
	if out != nil {
		t.Fatal("compatibility branches were selected")
	}
}

func TestPartialTextboxUnknownAndRevisionContexts(t *testing.T) {
	drawing := nativeTextboxFixture("vml", `<w:p><w:r><w:t>SECRET</w:t></w:r></w:p>`)
	base := nativeMutationMain(`<w:p><w:r>` + drawing + `</w:r></w:p>`)
	for _, main := range []string{
		strings.Replace(base, `<w:document `, `<w:document unknown="1" `, 1),
		strings.Replace(base, `<w:body>`, `<w:body unknown="1">`, 1),
		strings.Replace(base, `<w:r>`, `<w:r><w:rPr xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:MustUnderstand="x"/>`, 1),
		strings.Replace(base, `<w:p>`, `<w:p><w:moveFromRangeStart w:id="1"/>`, 1),
		strings.Replace(base, `<w:body>`, `<w:body><w:moveFromRangeStart w:id="1"/>`, 1),
	} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(main)))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := inspectNativePartialTextboxes(source, doc)
		if err != nil {
			t.Fatal(err)
		}
		if out != nil {
			for _, item := range out.Items {
				if item.Status == "supported" {
					t.Fatal("unknown global or revision context leaked text")
				}
			}
		}
	}
	ordinary := nativeTextboxFixture("drawingml", `<w:p><w:r><w:t>SECRET</w:t></w:r></w:p>`)
	for _, drawing := range []string{
		strings.Replace(ordinary, `<wps:spPr/>`, `<wps:spPr><a:extLst><a:ext uri="unknown"/></a:extLst></wps:spPr>`, 1),
		strings.Replace(ordinary, `<wps:bodyPr/>`, `<wps:bodyPr unknown="1"/>`, 1),
	} {
		source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+drawing+`</w:r></w:p>`))))
		doc, err := ExtractNativeDocumentV1(source)
		if err != nil {
			t.Fatal(err)
		}
		out, err := inspectNativePartialTextboxes(source, doc)
		if err != nil {
			t.Fatal(err)
		}
		if out == nil || out.Items[0].Status != "omitted" {
			t.Fatal("unknown drawing properties leaked text")
		}
	}
}

func TestPartialTextboxStylesSourceGrammar(t *testing.T) {
	const mc = `xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"`
	hiddenChoice := `<mc:AlternateContent ` + mc + `><mc:Choice Requires="w"><w:rPr><w:vanish/></w:rPr></mc:Choice><mc:Fallback/></mc:AlternateContent>`
	for _, test := range []struct {
		name, style string
		supported   bool
	}{
		{"empty styles", ``, true},
		{"ordinary defaults", `<w:docDefaults><w:rPrDefault><w:rPr><w:b/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>`, true},
		{"ordinary style inheritance", `<w:style w:type="paragraph" w:styleId="Base"><w:name w:val="Base"/><w:rPr><w:b/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:basedOn w:val="Base"/></w:style>`, true},
		{"hidden inherited style", `<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:vanish/></w:rPr></w:style>`, false},
		{"default MC choice", `<w:docDefaults><w:rPrDefault>` + hiddenChoice + `</w:rPrDefault></w:docDefaults>`, false},
		{"defaults MC choice", `<w:docDefaults>` + hiddenChoice + `</w:docDefaults>`, false},
		{"paragraph default MC choice", `<w:docDefaults><w:pPrDefault>` + hiddenChoice + `</w:pPrDefault></w:docDefaults>`, false},
		{"default properties unknown attr", `<w:docDefaults><w:rPrDefault unknown="1"><w:rPr/></w:rPrDefault></w:docDefaults>`, false},
		{"default run properties unknown attr", `<w:docDefaults><w:rPrDefault><w:rPr unknown="1"/></w:rPrDefault></w:docDefaults>`, false},
		{"default properties mixed text", `<w:docDefaults><w:rPrDefault>unparsed<w:rPr/></w:rPrDefault></w:docDefaults>`, false},
		{"default indirect style reference", `<w:docDefaults><w:rPrDefault><w:rPr><w:rStyle w:val="Hidden"/></w:rPr></w:rPrDefault></w:docDefaults>`, false},
		{"style MC choice", `<w:style w:type="paragraph" w:styleId="Normal" w:default="1">` + hiddenChoice + `</w:style>`, false},
		{"style unknown attr", `<w:style w:type="paragraph" w:styleId="Normal" unknown="1"/>`, false},
		{"style basedOn unknown attr", `<w:style w:type="paragraph" w:styleId="Normal"><w:basedOn w:val="Base" unknown="1"/></w:style>`, false},
		{"style basedOn child", `<w:style w:type="paragraph" w:styleId="Normal"><w:basedOn w:val="Base">` + hiddenChoice + `</w:basedOn></w:style>`, false},
		{"style indirect run reference", `<w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:rStyle w:val="Hidden"/></w:rPr></w:style>`, false},
		{"paragraph properties unknown attr", `<w:style w:type="paragraph" w:styleId="Normal"><w:pPr unknown="1"/></w:style>`, false},
		{"root unknown attr", `ROOT_ATTRIBUTE`, false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := nativeMutationParts(nativeMutationMain(`<w:p><w:r>` + nativeTextboxFixture("vml", `<w:p><w:r><w:t>Textbox source</w:t></w:r></w:p>`) + `</w:r></w:p>`))
			parts["word/_rels/document.xml.rels"] = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyles" Type="` + relBaseTransitional + `styles" Target="styles.xml"/></Relationships>`
			parts["word/styles.xml"] = `<w:styles xmlns:w="` + wordMLTransitional + `">` + test.style + `</w:styles>`
			if test.style == "ROOT_ATTRIBUTE" {
				parts["word/styles.xml"] = `<w:styles xmlns:w="` + wordMLTransitional + `" unknown="1"/>`
			}
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`, 1)
			source := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(source)
			if err != nil {
				t.Fatal(err)
			}
			out, err := inspectNativePartialTextboxes(source, doc)
			if err != nil {
				t.Fatal(err)
			}
			if out == nil || len(out.Items) != 1 {
				t.Fatal("source inventory lost")
			}
			item := out.Items[0]
			if (item.Status == "supported") != test.supported {
				t.Fatalf("unexpected qualification: %#v", item)
			}
			if !test.supported && len(item.Paragraphs) != 0 {
				t.Fatal("unqualified style context leaked text")
			}
		})
	}
}

func TestPartialTextboxUnjoinedDrawingDiagnostic(t *testing.T) {
	drawing := nativeTextboxFixture("drawingml", `<w:p><w:r><w:t>Source text</w:t></w:r></w:p>`)
	drawing = strings.Replace(drawing, `cx="914400"`, `cx="invalid"`, 1)
	source := buildNativeDOCX(t, nativeEntries(nativeMutationParts(nativeMutationMain(`<w:p><w:r>`+drawing+`</w:r></w:p>`))))
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if !hasUnsupportedCode(doc, "INVALID_DRAWING_EXTENT") {
		t.Fatal("original drawing refusal lost")
	}
	out, err := inspectNativePartialTextboxes(source, doc)
	if err != nil {
		t.Fatal(err)
	}
	if out != nil {
		t.Fatal("property-leaf diagnostic produced an unjoinable textbox sidecar")
	}
	if _, err := InspectNativePartialSourceV1(source); err != nil {
		t.Fatalf("ordinary omission inventory failed: %v", err)
	}
}
