package docxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
)

// A tiny valid 1x1 PNG (transparent pixel) — real bytes, not a placeholder
// string, since InsertImageAfter writes them straight into a real zip part
// and the python-docx validation test below re-parses the file.
var onePxPNG = mustDecodePNG()

func mustDecodePNG() []byte {
	b, err := base64.StdEncoding.DecodeString(
		"iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
	)
	if err != nil {
		panic(err)
	}
	return b
}

func buildImageTestDocx(t *testing.T) []byte {
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
			`<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>` +
			`<w:p><w:r><w:t>Second paragraph.</w:t></w:r></w:p>` +
			`<w:sectPr/>` +
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

func zipPart(t *testing.T, docx []byte, name string) (string, bool) {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(docx), int64(len(docx)))
	if err != nil {
		t.Fatal(err)
	}
	for _, f := range zr.File {
		if f.Name == name {
			rc, _ := f.Open()
			var buf bytes.Buffer
			buf.ReadFrom(rc) //nolint:errcheck
			rc.Close()
			return buf.String(), true
		}
	}
	return "", false
}

func TestInsertImageAfter_NoExistingRels(t *testing.T) {
	src := buildImageTestDocx(t)
	out, err := InsertImageAfter(src, 0, onePxPNG, "png", PixelsToEMU(200), PixelsToEMU(100))
	if err != nil {
		t.Fatal(err)
	}

	// New media part added with the exact bytes given.
	media, ok := zipPart(t, out, "word/media/image1.png")
	if !ok || media != string(onePxPNG) {
		t.Fatalf("media part missing or mismatched (ok=%v)", ok)
	}

	// New relationship created (no rels part existed before).
	rels, ok := zipPart(t, out, docRelsPart)
	if !ok || !strings.Contains(rels, `Type="`+relTypeImage+`"`) || !strings.Contains(rels, `Target="media/image1.png"`) {
		t.Fatalf("relationship missing/wrong: ok=%v rels=%s", ok, rels)
	}

	// Content types gained a Default for png.
	ct, _ := zipPart(t, out, contentTypes)
	if !strings.Contains(ct, `Extension="png"`) {
		t.Fatalf("content types missing png default: %s", ct)
	}

	// document.xml has a new paragraph with a drawing referencing the rel,
	// placed right after paragraph 0, and paragraph 0's own XML is untouched.
	doc, _ := zipPart(t, out, docPart)
	if !strings.Contains(doc, `r:embed="rId1"`) {
		t.Fatalf("drawing does not reference rId1: %s", doc)
	}
	if !strings.Contains(doc, `<w:p><w:r><w:t>First paragraph.</w:t></w:r></w:p>`) {
		t.Fatalf("original paragraph 0 changed: %s", doc)
	}
	firstIdx := strings.Index(doc, "First paragraph.")
	drawIdx := strings.Index(doc, "<w:drawing>")
	secondIdx := strings.Index(doc, "Second paragraph.")
	if !(firstIdx < drawIdx && drawIdx < secondIdx) {
		t.Fatalf("image paragraph not positioned between paragraph 0 and 1: %s", doc)
	}

	// Untouched parts byte-identical.
	origStyles, _ := zipPart(t, src, "word/styles.xml")
	outStyles, _ := zipPart(t, out, "word/styles.xml")
	if origStyles != outStyles {
		t.Fatal("word/styles.xml must be byte-identical")
	}
	origRootRels, _ := zipPart(t, src, "_rels/.rels")
	outRootRels, _ := zipPart(t, out, "_rels/.rels")
	if origRootRels != outRootRels {
		t.Fatal("_rels/.rels must be byte-identical")
	}
}

func TestInsertImageAfter_ExtendsExistingRelsAndReusesContentType(t *testing.T) {
	src := buildImageTestDocx(t)
	// Insert two images: the second must get a fresh media filename + rel
	// id, and must NOT duplicate the png Default entry.
	mid, err := InsertImageAfter(src, 0, onePxPNG, "png", PixelsToEMU(100), PixelsToEMU(100))
	if err != nil {
		t.Fatal(err)
	}
	out, err := InsertImageAfter(mid, 1, onePxPNG, "png", PixelsToEMU(50), PixelsToEMU(50))
	if err != nil {
		t.Fatal(err)
	}

	if _, ok := zipPart(t, out, "word/media/image1.png"); !ok {
		t.Fatal("image1.png missing")
	}
	if _, ok := zipPart(t, out, "word/media/image2.png"); !ok {
		t.Fatal("image2.png missing (media naming must not collide)")
	}
	rels, _ := zipPart(t, out, docRelsPart)
	if !strings.Contains(rels, `Id="rId1"`) || !strings.Contains(rels, `Id="rId2"`) {
		t.Fatalf("expected two distinct relationship ids: %s", rels)
	}
	ct, _ := zipPart(t, out, contentTypes)
	if strings.Count(ct, `Extension="png"`) != 1 {
		t.Fatalf("png Default entry must appear exactly once: %s", ct)
	}
}

func TestInsertImageAfter_RejectsBadInput(t *testing.T) {
	src := buildImageTestDocx(t)
	cases := []struct {
		name string
		fn   func() error
	}{
		{"unsupported extension", func() error {
			_, err := InsertImageAfter(src, 0, onePxPNG, "webp", 100, 100)
			return err
		}},
		{"empty image bytes", func() error {
			_, err := InsertImageAfter(src, 0, nil, "png", 100, 100)
			return err
		}},
		{"zero dimensions", func() error {
			_, err := InsertImageAfter(src, 0, onePxPNG, "png", 0, 100)
			return err
		}},
		{"out-of-range paragraph", func() error {
			_, err := InsertImageAfter(src, 99, onePxPNG, "png", 100, 100)
			return err
		}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			if err := c.fn(); err == nil {
				t.Fatal("expected an error, got nil")
			}
		})
	}
}

func TestInsertImageAfter_PrependBeforeFirstParagraph(t *testing.T) {
	src := buildImageTestDocx(t)
	out, err := InsertImageAfter(src, -1, onePxPNG, "png", PixelsToEMU(10), PixelsToEMU(10))
	if err != nil {
		t.Fatal(err)
	}
	doc, _ := zipPart(t, out, docPart)
	drawIdx := strings.Index(doc, "<w:drawing>")
	firstIdx := strings.Index(doc, "First paragraph.")
	if !(drawIdx >= 0 && drawIdx < firstIdx) {
		t.Fatalf("image paragraph must precede paragraph 0: %s", doc)
	}
}

// TestInsertImageAfter_OpensWithPythonDocx independently validates the
// produced file with python-docx (not our own reader) — the same
// never-trust-your-own-roundtrip discipline this project uses openpyxl for
// on the xlsx side. Skips (does not fail) when python3/python-docx aren't
// available in this environment, since CI installs them explicitly for this
// job (see .github/workflows/test.yml) but a bare `go test` elsewhere might
// not have them.
func TestInsertImageAfter_OpensWithPythonDocx(t *testing.T) {
	py, err := exec.LookPath("python3")
	if err != nil {
		t.Skip("python3 not available")
	}
	if err := exec.Command(py, "-c", "import docx").Run(); err != nil {
		t.Skip("python-docx not installed")
	}

	out, err := InsertImageAfter(buildImageTestDocx(t), 0, onePxPNG, "png", PixelsToEMU(200), PixelsToEMU(100))
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
d = docx.Document(sys.argv[1])
paras = [p.text for p in d.paragraphs]
assert paras[0] == "First paragraph.", paras
assert paras[-1] == "Second paragraph.", paras
# The inline shape must exist and be a real picture with the right size.
assert len(d.inline_shapes) == 1, d.inline_shapes
shape = d.inline_shapes[0]
assert shape.type == docx.enum.shape.WD_INLINE_SHAPE.PICTURE, shape.type
# EMU round-trip: 200px/100px at 96dpi.
assert shape.width == 200 * 914400 // 96, shape.width
assert shape.height == 100 * 914400 // 96, shape.height
print("OK")
`
	cmd := exec.Command(py, "-c", script, docxPath)
	outBytes, err := cmd.CombinedOutput()
	if err != nil {
		t.Fatalf("python-docx validation failed: %v\n%s", err, outBytes)
	}
	if !strings.Contains(string(outBytes), "OK") {
		t.Fatalf("unexpected python-docx output: %s", outBytes)
	}
}
