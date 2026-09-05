package docxpatch

import (
	"archive/zip"
	"bytes"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

func buildNoteTestDocx(t *testing.T) []byte {
	t.Helper()
	// Minimal InjOffice-authored Transitional WordprocessingML package used as
	// structural-only note-writer input. It is not Microsoft Office authoring
	// or visual-fidelity evidence.
	return buildNativeDOCX(t, nativeEntries(map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Types xmlns="` + opcContentTypesNS + `">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
			`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
			`</Types>`,
		"_rels/.rels": `<Relationships xmlns="` + opcRelationshipsNS + `">` +
			`<Relationship Id="rId1" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/>` +
			`</Relationships>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `">` +
			`<Relationship Id="rId1" Type="` + relBaseTransitional + `styles" Target="styles.xml"/>` +
			`</Relationships>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` +
			`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Title paragraph</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>First body paragraph.</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>Second body paragraph.</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>Third body paragraph.</w:t></w:r></w:p>` +
			`<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>` +
			`</w:body></w:document>`,
		"word/styles.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:styles xmlns:w="` + wordMLTransitional + `">` +
			`<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Fixture Sans" w:hAnsi="Fixture Sans"/><w:sz w:val="24"/></w:rPr></w:rPrDefault>` +
			`<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>` +
			`<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
			`<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>` +
			`<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="200" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="28"/></w:rPr></w:style>` +
			`<w:style w:type="paragraph" w:styleId="Heading3"><w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:qFormat/><w:pPr><w:keepNext/><w:spacing w:before="160" w:after="120"/></w:pPr><w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>` +
			`</w:styles>`,
	}))
}

func TestInsertNote_Footnote_NoExistingPart(t *testing.T) {
	src := buildNoteTestDocx(t)
	out, id, err := InsertNote(src, Footnote, 0, "See Smith et al., 2024.")
	if err != nil {
		t.Fatal(err)
	}
	if id != 1 {
		t.Fatalf("first real footnote id should be 1, got %d", id)
	}

	notes, ok := zipPart(t, out, "word/footnotes.xml")
	if !ok {
		t.Fatal("word/footnotes.xml missing")
	}
	if !strings.Contains(notes, `w:id="1"`) || !strings.Contains(notes, "See Smith et al., 2024.") {
		t.Fatalf("footnote body missing/wrong: %s", notes)
	}
	if !strings.Contains(notes, `w:type="separator" w:id="-1"`) || !strings.Contains(notes, `w:type="continuationSeparator" w:id="0"`) {
		t.Fatalf("missing separator/continuationSeparator boilerplate: %s", notes)
	}

	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, `<w:footnoteReference w:id="1"/>`) || !strings.Contains(doc, `w:rStyle w:val="FootnoteReference"`) {
		t.Fatalf("document.xml missing footnote reference: %s", doc)
	}
	// Reference appended AFTER the paragraph's existing text.
	firstIdx := strings.Index(doc, "Title paragraph")
	refIdx := strings.Index(doc, "<w:footnoteReference")
	secondIdx := strings.Index(doc, "First body paragraph")
	if !(firstIdx < refIdx && refIdx < secondIdx) {
		t.Fatalf("reference not positioned inside paragraph 0: %s", doc)
	}

	rels, _ := zipPart(t, out, docRelsPart)
	if !strings.Contains(rels, `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"`) ||
		!strings.Contains(rels, `Target="footnotes.xml"`) {
		t.Fatalf("relationship missing/wrong: %s", rels)
	}

	ct, _ := zipPart(t, out, contentTypes)
	if !strings.Contains(ct, `PartName="/word/footnotes.xml"`) {
		t.Fatalf("content types missing footnotes override: %s", ct)
	}

	// The source fixture styles remain intact while the note style is appended.
	origStyles, _ := zipPart(t, src, "word/styles.xml")
	outStyles, _ := zipPart(t, out, "word/styles.xml")
	if !strings.Contains(outStyles, origStyles[:strings.LastIndex(origStyles, "</w:styles>")]) || !strings.Contains(outStyles, `w:styleId="FootnoteReference"`) || !strings.Contains(outStyles, `<w:vertAlign w:val="superscript"/>`) {
		t.Fatal("source fixture styles or note-reference baseline semantics changed")
	}
}

func TestInsertNote_SyntheticFixturePreservesSuperscriptButNativeResolutionRefusesIt(t *testing.T) {
	for _, kind := range []NoteKind{Footnote, Endnote} {
		t.Run(string(kind), func(t *testing.T) {
			out, _, err := InsertNote(buildNoteTestDocx(t), kind, 0, "Synthetic fixture note.")
			if err != nil {
				t.Fatal(err)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(out)
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, diagnostic := range resolved.Diagnostics {
				found = found || diagnostic.Code == "VERTICAL_ALIGNMENT_UNSUPPORTED"
			}
			if !found {
				t.Fatalf("%s superscript must remain non-authoritative for native layout: %#v", kind, resolved.Diagnostics)
			}
		})
	}
}

func TestInsertNote_Endnote_NoExistingPart(t *testing.T) {
	src := buildNoteTestDocx(t)
	out, id, err := InsertNote(src, Endnote, 1, "Further reading.")
	if err != nil {
		t.Fatal(err)
	}
	if id != 1 {
		t.Fatalf("first real endnote id should be 1, got %d", id)
	}
	notes, ok := zipPart(t, out, "word/endnotes.xml")
	if !ok || !strings.Contains(notes, "Further reading.") {
		t.Fatalf("endnote body missing/wrong (ok=%v): %s", ok, notes)
	}
	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, `<w:endnoteReference w:id="1"/>`) || !strings.Contains(doc, `w:rStyle w:val="EndnoteReference"`) {
		t.Fatalf("document.xml missing endnote reference: %s", doc)
	}
}

func TestInsertNote_WritesTransitionalAndStrictDialectsDirectly(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name := "transitional"
		wordNS, relBase := wordMLTransitional, relBaseTransitional
		if strict {
			name, wordNS, relBase = "strict", wordMLStrict, relBaseStrict
		}
		t.Run(name, func(t *testing.T) {
			parts := map[string]string{
				"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/Odd/Main.XML" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
				"_rels/.rels":         `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBase + `officeDocument" Target="Odd/Main.XML"/></Relationships>`,
				"Odd/Main.XML":        `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body></w:document>`,
			}
			out, id, err := InsertNote(buildNativeDOCX(t, nativeEntries(parts)), Footnote, 0, "strict-safe note")
			if err != nil {
				t.Fatal(err)
			}
			if id != 1 {
				t.Fatalf("id=%d", id)
			}
			pkg, err := openNativeDOCXPackage(out)
			if err != nil {
				t.Fatal(err)
			}
			main, gotStrict, err := pkg.officeDocumentPart()
			if err != nil || main != "Odd/Main.XML" || gotStrict != strict {
				t.Fatalf("dialect/main changed: main=%q strict=%v err=%v", main, gotStrict, err)
			}
			rel, err := nativeWriterSingletonRelationship(pkg, main, relBase+"footnotes", relBaseStrict+"footnotes", relBaseTransitional+"footnotes")
			if err != nil || rel == nil {
				t.Fatalf("note relationship: rel=%#v err=%v", rel, err)
			}
			root, _, err := inspectNativeNotePart(pkg.files[rel.PartName], noteShape{part: rel.PartName, rootEl: "w:footnotes", noteEl: "w:footnote"}, wordNS)
			if err != nil || root.Name.Space != wordNS {
				t.Fatalf("note dialect changed: root=%#v err=%v", root, err)
			}
		})
	}
}

func TestInsertNote_AtomicallyRefusesNonExactWriterInputs(t *testing.T) {
	source, _, err := InsertNote(buildNoteTestDocx(t), Footnote, 0, "first")
	if err != nil {
		t.Fatal(err)
	}
	for name, mutate := range map[string]func(map[string][]byte){
		"explicit TargetMode": func(parts map[string][]byte) {
			parts[docRelsPart] = []byte(strings.Replace(string(parts[docRelsPart]), `Target="footnotes.xml"`, `Target="footnotes.xml" TargetMode="Internal"`, 1))
		},
		"noncanonical Target": func(parts map[string][]byte) {
			parts[docRelsPart] = []byte(strings.Replace(string(parts[docRelsPart]), `Target="footnotes.xml"`, `Target="./footnotes.xml"`, 1))
		},
		"noncanonical note id": func(parts map[string][]byte) {
			parts["word/footnotes.xml"] = []byte(strings.Replace(string(parts["word/footnotes.xml"]), `w:id="1"`, `w:id="01"`, 1))
		},
		"conflicting reference style": func(parts map[string][]byte) {
			parts["word/styles.xml"] = []byte(strings.Replace(string(parts["word/styles.xml"]), `<w:vertAlign w:val="superscript"/>`, `<w:vertAlign w:val="subscript"/>`, 1))
		},
	} {
		t.Run(name, func(t *testing.T) {
			parts := map[string][]byte{}
			for _, partName := range []string{docRelsPart, "word/footnotes.xml", "word/styles.xml"} {
				value, ok := zipPart(t, source, partName)
				if !ok {
					t.Fatalf("missing %s", partName)
				}
				parts[partName] = []byte(value)
			}
			mutate(parts)
			candidate, patchErr := ApplyPatch(source, Patch{Replace: parts, Add: map[string][]byte{}, Delete: map[string]bool{}})
			if patchErr != nil {
				t.Fatal(patchErr)
			}
			if out, _, insertErr := InsertNote(candidate, Footnote, 1, "second"); insertErr == nil || out != nil {
				t.Fatalf("non-exact input produced output: err=%v bytes=%d", insertErr, len(out))
			}
		})
	}
	for _, invalid := range []string{"bad\rtext", "bad\x00text"} {
		if out, _, insertErr := InsertNote(source, Footnote, 1, invalid); insertErr == nil || out != nil {
			t.Fatalf("invalid XML 1.0 text produced output: err=%v bytes=%d", insertErr, len(out))
		}
	}
}

func TestInsertNote_SecondNoteExtendsExistingPartWithFreshID(t *testing.T) {
	src := buildNoteTestDocx(t)
	mid, id1, err := InsertNote(src, Footnote, 0, "First citation.")
	if err != nil {
		t.Fatal(err)
	}
	out, id2, err := InsertNote(mid, Footnote, 1, "Second citation.")
	if err != nil {
		t.Fatal(err)
	}
	if id1 != 1 || id2 != 2 {
		t.Fatalf("expected sequential ids 1, 2 — got %d, %d", id1, id2)
	}
	notes, _ := zipPart(t, out, "word/footnotes.xml")
	if !strings.Contains(notes, "First citation.") || !strings.Contains(notes, "Second citation.") {
		t.Fatalf("both footnote bodies must be present: %s", notes)
	}
	// Only ONE separator/continuationSeparator pair — the second insert must
	// extend the existing part, not recreate it.
	if strings.Count(notes, `w:type="separator"`) != 1 {
		t.Fatalf("separator boilerplate must appear exactly once: %s", notes)
	}
	rels, _ := zipPart(t, out, docRelsPart)
	if strings.Count(rels, "relationships/footnotes") != 1 {
		t.Fatalf("only one footnotes relationship should exist across both inserts: %s", rels)
	}
}

func TestInsertNote_RejectsBadInput(t *testing.T) {
	src := buildNoteTestDocx(t)
	if _, _, err := InsertNote(src, Footnote, 0, ""); err == nil {
		t.Fatal("empty note text must be rejected")
	}
	if _, _, err := InsertNote(src, Footnote, 0, "   "); err == nil {
		t.Fatal("whitespace-only note text must be rejected")
	}
	if _, _, err := InsertNote(src, Footnote, 99, "text"); err == nil {
		t.Fatal("out-of-range paragraph must be rejected")
	}
	if _, _, err := InsertNote(src, NoteKind("comment"), 0, "text"); err == nil {
		t.Fatal("unknown note kind must be rejected")
	}
}

func TestInsertNote_AppendsToParagraphWithExistingDrawing(t *testing.T) {
	// Appending a note reference never destroys existing non-text content —
	// unlike Apply's "set", which must refuse a HasNonText paragraph,
	// InsertNote only ever adds a sibling run at the end.
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
			`<w:p><w:r><w:drawing>img</w:drawing></w:r><w:r><w:t>caption</w:t></w:r></w:p>` +
			`</w:body></w:document>`,
	}
	for name, content := range parts {
		w, _ := zw.Create(name)
		w.Write([]byte(content)) //nolint:errcheck
	}
	zw.Close() //nolint:errcheck
	src := buf.Bytes()

	out, _, err := InsertNote(src, Footnote, 0, "note on the image caption")
	if err != nil {
		t.Fatal(err)
	}
	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, "<w:drawing>img</w:drawing>") {
		t.Fatalf("existing drawing destroyed: %s", doc)
	}
	if !strings.Contains(doc, "<w:footnoteReference") {
		t.Fatalf("footnote reference not appended: %s", doc)
	}
}

// TestInsertNote_OpensWithPythonDocxOxml independently validates the
// produced .docx via python-docx's LOW-LEVEL docx.oxml/lxml access —
// python-docx 1.2.0 has no high-level footnote object model, but its
// package/relationship/part machinery (opc layer) is real, independent
// code that still parses the produced file's rels graph and lets us pull
// the raw footnotes.xml part and inspect it with lxml. Same
// never-trust-your-own-writer discipline as the D11-images python-docx
// test and the project's openpyxl validation elsewhere.
func TestInsertNote_OpensWithPythonDocxOxml(t *testing.T) {
	py, err := exec.LookPath("python3")
	if err != nil {
		t.Skip("python3 not available")
	}
	if err := exec.Command(py, "-c", "import docx, lxml.etree").Run(); err != nil {
		t.Skip("python-docx / lxml not installed")
	}

	out, _, err := InsertNote(buildNoteTestDocx(t), Footnote, 0, "See Smith et al., 2024.")
	if err != nil {
		t.Fatal(err)
	}
	dir := t.TempDir()
	docxPath := filepath.Join(dir, "out.docx")
	if err := os.WriteFile(docxPath, out, 0o644); err != nil {
		t.Fatal(err)
	}

	script := `
import sys
import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

d = docx.Document(sys.argv[1])

# The document (package) OPENS AT ALL via python-docx's real OPC package
# loader (validates zip integrity, [Content_Types].xml, and the rels graph
# independently of our own writer) -- paragraphs are readable.
paras = [p.text for p in d.paragraphs]
assert paras[0] == "Title paragraph", paras
assert paras[1] == "First body paragraph.", paras

# Find the footnotes relationship off the document part and pull the raw
# part bytes -- this is python-docx's real opc.part.Part machinery, not
# our own reader.
footnotes_rel = None
for rel in d.part.rels.values():
    if rel.reltype == RT.FOOTNOTES:
        footnotes_rel = rel
        break
assert footnotes_rel is not None, "no footnotes relationship found by python-docx"
footnotes_xml = footnotes_rel.target_part.blob

ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
root = etree.fromstring(footnotes_xml)
notes = root.findall("w:footnote", ns)
ids = {n.get("{%s}id" % ns["w"]) for n in notes}
assert {"-1", "0", "1"} <= ids, ids

note1 = [n for n in notes if n.get("{%s}id" % ns["w"]) == "1"][0]
text = "".join(t.text or "" for t in note1.iter("{%s}t" % ns["w"]))
assert "Smith et al., 2024" in text, text

# The reference in document.xml really points at id 1.
doc_xml = d.element.xml
doc_root = etree.fromstring(doc_xml.encode("utf-8"))
refs = doc_root.findall(".//w:footnoteReference", ns)
assert len(refs) == 1, refs
assert refs[0].get("{%s}id" % ns["w"]) == "1", refs[0].attrib

print("OK")
`
	cmd := exec.Command(py, "-c", script, docxPath)
	outBytes, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("python-docx/lxml validation failed: %v\n%s", err, outBytes)
	}
	if !strings.Contains(string(outBytes), "OK") {
		t.Fatalf("unexpected validation output: %s", outBytes)
	}
}
