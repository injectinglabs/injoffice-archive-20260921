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

func buildWatermarkTestDocx(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
			`<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>` +
			`</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
			`</Relationships>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
			`<w:p><w:r><w:t>Hello watermark.</w:t></w:r></w:p>` +
			`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
			`</w:body></w:document>`,
		"word/styles.xml": `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`,
	}
	for name, content := range parts {
		w, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := w.Write([]byte(content)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func TestInsertWatermark_NoExistingHeader(t *testing.T) {
	src := buildWatermarkTestDocx(t)
	out, err := InsertWatermark(src, WatermarkSpec{Text: "CONFIDENTIAL"})
	if err != nil {
		t.Fatal(err)
	}

	headerXML, ok := zipPart(t, out, "word/header1.xml")
	if !ok {
		t.Fatal("word/header1.xml missing")
	}
	if !strings.Contains(headerXML, "CONFIDENTIAL") {
		t.Fatalf("watermark text missing: %s", headerXML)
	}
	if !strings.Contains(headerXML, `type="#_x0000_t136"`) {
		t.Fatalf("watermark shapetype reference missing: %s", headerXML)
	}
	if !strings.Contains(headerXML, `rotation:315`) {
		t.Fatalf("expected default diagonal rotation: %s", headerXML)
	}
	if !strings.Contains(headerXML, `fillcolor="#808080"`) {
		t.Fatalf("expected default gray fill: %s", headerXML)
	}

	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, `<w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>`) {
		t.Fatalf("sectPr missing headerReference: %s", doc)
	}
	// pgSz (an existing sectPr child) survives.
	if !strings.Contains(doc, `<w:pgSz w:w="12240" w:h="15840"/>`) {
		t.Fatalf("existing sectPr content lost: %s", doc)
	}
	// Body text untouched.
	if !strings.Contains(doc, "Hello watermark.") {
		t.Fatalf("body content lost: %s", doc)
	}

	rels, _ := zipPart(t, out, docRelsPart)
	if !strings.Contains(rels, `Type="`+relTypeHeader+`"`) || !strings.Contains(rels, `Target="header1.xml"`) {
		t.Fatalf("header relationship missing/wrong: %s", rels)
	}

	ct, _ := zipPart(t, out, contentTypes)
	if !strings.Contains(ct, `PartName="/word/header1.xml"`) {
		t.Fatalf("content types missing header override: %s", ct)
	}

	origStyles, _ := zipPart(t, src, "word/styles.xml")
	outStyles, _ := zipPart(t, out, "word/styles.xml")
	if origStyles != outStyles {
		t.Fatal("word/styles.xml must be byte-identical")
	}
}

func TestInsertWatermark_AppendsToExistingHeader(t *testing.T) {
	src := buildWatermarkTestDocx(t)
	// First insert creates the header; a second call must APPEND to it
	// (extend, not replace or duplicate the relationship/content-type).
	mid, err := InsertWatermark(src, WatermarkSpec{Text: "DRAFT"})
	if err != nil {
		t.Fatal(err)
	}
	// Simulate "existing header had real content already": read it, verify
	// re-running InsertWatermark on the SAME doc (already has a default
	// header) appends rather than creating a second header part.
	out, err := InsertWatermark(mid, WatermarkSpec{Text: "SECOND PASS"})
	if err != nil {
		t.Fatal(err)
	}
	headerXML, ok := zipPart(t, out, "word/header1.xml")
	if !ok {
		t.Fatal("word/header1.xml missing")
	}
	if !strings.Contains(headerXML, "DRAFT") || !strings.Contains(headerXML, "SECOND PASS") {
		t.Fatalf("expected BOTH watermark paragraphs present (append, not replace): %s", headerXML)
	}
	if _, ok := zipPart(t, out, "word/header2.xml"); ok {
		t.Fatal("must not create a second header part when a default header already exists")
	}
	rels, _ := zipPart(t, out, docRelsPart)
	if strings.Count(rels, "relationships/header") != 1 {
		t.Fatalf("only one header relationship should exist: %s", rels)
	}
	doc, _ := zipPart(t, out, docPart)
	if strings.Count(doc, "<w:headerReference") != 1 {
		t.Fatalf("only one headerReference should exist: %s", doc)
	}
}

func TestInsertWatermark_HorizontalOption(t *testing.T) {
	src := buildWatermarkTestDocx(t)
	out, err := InsertWatermark(src, WatermarkSpec{Text: "FLAT", Horizontal: true})
	if err != nil {
		t.Fatal(err)
	}
	headerXML, _ := zipPart(t, out, "word/header1.xml")
	if !strings.Contains(headerXML, "rotation:0") {
		t.Fatalf("expected horizontal (rotation:0): %s", headerXML)
	}
}

func TestInsertWatermark_CustomColorAndFont(t *testing.T) {
	src := buildWatermarkTestDocx(t)
	out, err := InsertWatermark(src, WatermarkSpec{Text: "SAMPLE", ColorHex: "FF0000", FontFamily: "Georgia"})
	if err != nil {
		t.Fatal(err)
	}
	headerXML, _ := zipPart(t, out, "word/header1.xml")
	if !strings.Contains(headerXML, `fillcolor="#FF0000"`) {
		t.Fatalf("custom color not applied: %s", headerXML)
	}
	if !strings.Contains(headerXML, "Georgia") {
		t.Fatalf("custom font not applied: %s", headerXML)
	}
}

func TestInsertWatermark_SelfClosingSectPr(t *testing.T) {
	// A minimal <w:sectPr/> (no children yet) — matches the fixture shape
	// used by imagewrite_test.go/notewrite_test.go/chartwrite_test.go.
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="x"/></Types>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
			`<w:p><w:r><w:t>text</w:t></w:r></w:p>` +
			`<w:sectPr/>` +
			`</w:body></w:document>`,
	}
	for name, content := range parts {
		w, _ := zw.Create(name)
		w.Write([]byte(content)) //nolint:errcheck
	}
	zw.Close() //nolint:errcheck

	out, err := InsertWatermark(buf.Bytes(), WatermarkSpec{Text: "X"})
	if err != nil {
		t.Fatal(err)
	}
	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, `<w:sectPr><w:headerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></w:sectPr>`) {
		t.Fatalf("self-closing sectPr not expanded correctly: %s", doc)
	}
}

func TestInsertWatermark_RejectsBadInput(t *testing.T) {
	src := buildWatermarkTestDocx(t)
	if _, err := InsertWatermark(src, WatermarkSpec{Text: ""}); err == nil {
		t.Fatal("empty text must be rejected")
	}
	if _, err := InsertWatermark(src, WatermarkSpec{Text: "  "}); err == nil {
		t.Fatal("whitespace-only text must be rejected")
	}
	if _, err := InsertWatermark(src, WatermarkSpec{Text: "X", ColorHex: "notacolor"}); err == nil {
		t.Fatal("invalid color must be rejected")
	}
}

func TestInsertWatermark_RejectsMultiSection(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="x"/></Types>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
			`<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="1"/></w:sectPr></w:pPr><w:r><w:t>section 1</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>section 2</w:t></w:r></w:p>` +
			`<w:sectPr/>` +
			`</w:body></w:document>`,
	}
	for name, content := range parts {
		w, _ := zw.Create(name)
		w.Write([]byte(content)) //nolint:errcheck
	}
	zw.Close() //nolint:errcheck

	if _, err := InsertWatermark(buf.Bytes(), WatermarkSpec{Text: "X"}); err == nil {
		t.Fatal("multi-section document must be rejected, not silently mishandled")
	}
}

// TestInsertWatermark_OpensWithPythonDocxOxml independently validates via
// python-docx: section.header is a REAL, first-class python-docx object
// (unlike theme/chart, which have no high-level API at all) — this proves
// the header relationship + section wiring resolves correctly through
// python-docx's OWN OPC/section machinery, independent of anything this
// package wrote. lxml then confirms the VML shape structure and watermark
// text on the header part's raw XML (python-docx doesn't parse VML
// itself). Visual rendering (rotation/opacity/position) is NOT verified —
// that needs a real Word/LibreOffice render; LibreOffice headless isn't
// installed on this machine (checked, not assumed) — flagged, not faked.
func TestInsertWatermark_OpensWithPythonDocxOxml(t *testing.T) {
	py, err := exec.LookPath("python3")
	if err != nil {
		t.Skip("python3 not available")
	}
	if err := exec.Command(py, "-c", "import docx, lxml.etree").Run(); err != nil {
		t.Skip("python-docx / lxml not installed")
	}

	out, err := InsertWatermark(buildWatermarkTestDocx(t), WatermarkSpec{Text: "CONFIDENTIAL DRAFT"})
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
from lxml import etree

d = docx.Document(sys.argv[1])
assert d.paragraphs[0].text == "Hello watermark.", d.paragraphs[0].text

# section.header is python-docx's own real object model -- this resolves
# the headerReference -> relationship -> part chain independently.
section = d.sections[0]
header = section.header
assert not header.is_linked_to_previous, "a real (non-inherited) header must be wired"
header_xml = header.part.blob

ns = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "v": "urn:schemas-microsoft-com:vml",
}
root = etree.fromstring(header_xml)
shape = root.find(".//v:shape", ns)
assert shape is not None, "no VML shape in header"
assert shape.get("type") == "#_x0000_t136", shape.get("type")
textpath = shape.find("v:textpath", ns)
assert textpath is not None
assert textpath.get("string") == "CONFIDENTIAL DRAFT", textpath.get("string")

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
