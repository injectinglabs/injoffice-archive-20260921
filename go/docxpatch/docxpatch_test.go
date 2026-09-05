package docxpatch

import (
	"archive/zip"
	"bytes"
	"io"
	"strings"
	"testing"
)

const testDoc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
	`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Quarterly </w:t></w:r><w:r><w:t>Report</w:t></w:r></w:p>` +
	`<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Revenue grew </w:t></w:r><w:r><w:t>fast &amp; steady</w:t></w:r></w:p>` +
	`<w:p/>` +
	`<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Cell one</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Cell two</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
	`<w:p><w:r><w:drawing>img</w:drawing></w:r><w:r><w:t>caption</w:t></w:r></w:p>` +
	`<w:p><w:r><w:t>The end.</w:t></w:r></w:p>` +
	`</w:body></w:document>`

func buildDocx(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml":   `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
		"word/document.xml":     testDoc,
		"word/styles.xml":       `<w:styles xmlns:w="x"><w:style w:styleId="Heading1"/></w:styles>`,
		"word/media/image1.png": "PNGBYTES",
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

func part(t *testing.T, docx []byte, name string) string {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name == name {
			rc, _ := f.Open()
			b, _ := io.ReadAll(rc)
			rc.Close()
			return string(b)
		}
	}
	t.Fatalf("part %s missing", name)
	return ""
}

func TestExtract(t *testing.T) {
	paras, err := Extract(buildDocx(t))
	if err != nil {
		t.Fatal(err)
	}
	if len(paras) != 7 {
		t.Fatalf("want 7 paragraphs, got %d: %+v", len(paras), paras)
	}
	if paras[0].Text != "Quarterly Report" || paras[0].Style != "Heading1" {
		t.Errorf("p0: %+v", paras[0])
	}
	if paras[1].Text != "Revenue grew fast & steady" {
		t.Errorf("p1 text: %q", paras[1].Text)
	}
	if paras[2].Text != "" || paras[2].InTable {
		t.Errorf("p2 (empty): %+v", paras[2])
	}
	if !paras[3].InTable || paras[3].Text != "Cell one" || !paras[4].InTable || paras[4].Text != "Cell two" {
		t.Errorf("table paras: %+v %+v", paras[3], paras[4])
	}
	if !paras[5].HasNonText {
		t.Errorf("drawing para should be non-text: %+v", paras[5])
	}
	if paras[6].Text != "The end." || paras[6].InTable {
		t.Errorf("p6: %+v", paras[6])
	}
}

func TestApplySetPreservesEverythingElse(t *testing.T) {
	src := buildDocx(t)
	out, err := Apply(src, []Edit{{Op: "set", Index: 1, Text: "Revenue fell <sharply> & sadly.\nSecond line"}})
	if err != nil {
		t.Fatal(err)
	}
	paras, _ := Extract(out)
	if paras[1].Text != "Revenue fell <sharply> & sadly.\nSecond line" {
		t.Fatalf("edited text: %q", paras[1].Text)
	}
	// Untouched parts byte-identical.
	if part(t, out, "word/styles.xml") != part(t, src, "word/styles.xml") ||
		part(t, out, "word/media/image1.png") != part(t, src, "word/media/image1.png") {
		t.Fatal("untouched parts must be verbatim")
	}
	doc := part(t, out, "word/document.xml")
	// Untouched paragraphs verbatim.
	for _, verbatim := range []string{
		`<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>Quarterly </w:t></w:r><w:r><w:t>Report</w:t></w:r></w:p>`,
		`<w:p><w:r><w:drawing>img</w:drawing></w:r><w:r><w:t>caption</w:t></w:r></w:p>`,
	} {
		if !strings.Contains(doc, verbatim) {
			t.Errorf("untouched paragraph changed:\n%s", doc)
		}
	}
	// The edited paragraph keeps its first run's rPr (italic).
	if !strings.Contains(doc, `<w:rPr><w:i/></w:rPr><w:t xml:space="preserve">Revenue fell &lt;sharply&gt; &amp; sadly.</w:t>`) {
		t.Errorf("rPr/escaping lost:\n%s", doc)
	}
	if !strings.Contains(doc, `<w:br/><w:t xml:space="preserve">Second line</w:t>`) {
		t.Errorf("line break lost:\n%s", doc)
	}
}

func TestApplyRefusesNonText(t *testing.T) {
	if _, err := Apply(buildDocx(t), []Edit{{Op: "set", Index: 5, Text: "x"}}); err == nil {
		t.Fatal("editing a drawing paragraph must fail")
	}
}

func TestApplyInsertDeleteAndOrder(t *testing.T) {
	out, err := Apply(buildDocx(t), []Edit{
		{Op: "insert_after", Index: 6, Text: "Appendix"},
		{Op: "delete", Index: 2},
		{Op: "set", Index: 0, Text: "Annual Report"},
	})
	if err != nil {
		t.Fatal(err)
	}
	paras, _ := Extract(out)
	texts := make([]string, len(paras))
	for i, p := range paras {
		texts[i] = p.Text
	}
	want := []string{"Annual Report", "Revenue grew fast & steady", "Cell one", "Cell two", "caption", "The end.", "Appendix"}
	if len(paras) != len(want) {
		t.Fatalf("want %d paras, got %v", len(want), texts)
	}
	for i := range want {
		if texts[i] != want[i] {
			t.Fatalf("para %d = %q, want %q (all: %v)", i, texts[i], want[i], texts)
		}
	}
	// Inserted paragraph inherits the anchor's (plain) formatting; heading
	// edit keeps the Heading1 style + bold rPr.
	if paras[0].Style != "Heading1" {
		t.Errorf("style lost on set: %+v", paras[0])
	}
}

func TestInsertBeforeStart(t *testing.T) {
	out, err := Apply(buildDocx(t), []Edit{{Op: "insert_after", Index: -1, Text: "Preface"}})
	if err != nil {
		t.Fatal(err)
	}
	paras, _ := Extract(out)
	if paras[0].Text != "Preface" || paras[0].Style != "Heading1" {
		t.Fatalf("prepend: %+v", paras[0])
	}
}

func TestReplaceText(t *testing.T) {
	out, changed, skipped, err := ReplaceText(buildDocx(t), "Revenue grew fast", "Revenue DOUBLED")
	if err != nil || changed != 1 || skipped != 0 {
		t.Fatalf("replace: %d %d %v", changed, skipped, err)
	}
	paras, _ := Extract(out)
	if paras[1].Text != "Revenue DOUBLED & steady" {
		t.Fatalf("cross-run replace failed: %q", paras[1].Text)
	}
	// Occurrence inside a non-text paragraph is skipped, not destroyed.
	_, changed, skipped, err = ReplaceText(buildDocx(t), "caption", "CAPTION")
	if err != nil || changed != 0 || skipped != 1 {
		t.Fatalf("non-text skip: %d %d %v", changed, skipped, err)
	}
}

func TestEmptyParagraphSet(t *testing.T) {
	out, err := Apply(buildDocx(t), []Edit{{Op: "set", Index: 2, Text: "no longer empty"}})
	if err != nil {
		t.Fatal(err)
	}
	paras, _ := Extract(out)
	if paras[2].Text != "no longer empty" {
		t.Fatalf("self-closed set: %+v", paras[2])
	}
}
