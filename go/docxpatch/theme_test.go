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

// A realistic theme1.xml: dk1/lt1 as sysClr (Word's actual default —
// customizing them is exactly the case that exercises the sysClr->srgbClr
// conversion in ApplyTheme), the rest srgbClr, real major/minor fonts, and
// a substantial fmtScheme block (fill/line/effect styles) that must
// survive every ApplyTheme call byte-identical — this is the part of a
// real theme this patcher deliberately never touches.
const testThemeXML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
<a:themeElements>
<a:clrScheme name="Office">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="44546A"/></a:dk2>
<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Office">
<a:majorFont>
<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
</a:majorFont>
<a:minorFont>
<a:latin typeface="Calibri" panose="020F0502020204030204"/>
<a:ea typeface=""/>
<a:cs typeface=""/>
</a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Office">
<a:fillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/></a:schemeClr></a:gs></a:gsLst></a:gradFill>
<a:noFill/>
</a:fillStyleLst>
<a:lnStyleLst>
<a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
</a:lnStyleLst>
<a:effectStyleLst>
<a:effectStyle><a:effectLst/></a:effectStyle>
</a:effectStyleLst>
<a:bgFillStyleLst>
<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
</a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`

func buildThemeTestDocx(t *testing.T) []byte {
	t.Helper()
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
			`<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/>` +
			`<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>` +
			`<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>` +
			`</Types>`,
		"_rels/.rels": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
			`</Relationships>`,
		docRelsPart: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
			`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
			`</Relationships>`,
		"word/document.xml": `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>` +
			`<w:p><w:r><w:t>Hello theme.</w:t></w:r></w:p>` +
			`<w:sectPr/>` +
			`</w:body></w:document>`,
		"word/theme/theme1.xml": testThemeXML,
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

func TestExtractTheme(t *testing.T) {
	theme, err := ExtractTheme(buildThemeTestDocx(t))
	if err != nil {
		t.Fatal(err)
	}
	if theme.Name != "Office Theme" {
		t.Errorf("name: %q", theme.Name)
	}
	want := map[ThemeColorSlot]string{
		ThemeDark1: "000000", ThemeLight1: "FFFFFF",
		ThemeDark2: "44546A", ThemeLight2: "E7E6E6",
		ThemeAccent1: "4472C4", ThemeAccent2: "ED7D31", ThemeAccent3: "A5A5A5",
		ThemeAccent4: "FFC000", ThemeAccent5: "5B9BD5", ThemeAccent6: "70AD47",
		ThemeHyperlink: "0563C1", ThemeFollowedHyperlink: "954F72",
	}
	for slot, hex := range want {
		if theme.Colors[slot] != hex {
			t.Errorf("slot %s: got %q, want %q", slot, theme.Colors[slot], hex)
		}
	}
	if theme.MajorLatin != "Calibri Light" {
		t.Errorf("major font: %q", theme.MajorLatin)
	}
	if theme.MinorLatin != "Calibri" {
		t.Errorf("minor font: %q", theme.MinorLatin)
	}
}

func TestExtractTheme_NoThemePart(t *testing.T) {
	// A docx with no theme relationship at all.
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="x"/></Types>`,
		"word/document.xml":   `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`,
	}
	for name, content := range parts {
		w, _ := zw.Create(name)
		w.Write([]byte(content)) //nolint:errcheck
	}
	zw.Close() //nolint:errcheck
	if _, err := ExtractTheme(buf.Bytes()); err == nil {
		t.Fatal("expected an error for a docx with no theme part")
	}
}

func TestApplyTheme_Colors(t *testing.T) {
	src := buildThemeTestDocx(t)
	out, err := ApplyTheme(src, ThemePatch{Colors: map[ThemeColorSlot]string{
		ThemeAccent1: "#FF0000",
		ThemeDark1:   "112233", // was sysClr — must convert to srgbClr
	}})
	if err != nil {
		t.Fatal(err)
	}
	theme, err := ExtractTheme(out)
	if err != nil {
		t.Fatal(err)
	}
	if theme.Colors[ThemeAccent1] != "FF0000" {
		t.Errorf("accent1 not applied: %q", theme.Colors[ThemeAccent1])
	}
	if theme.Colors[ThemeDark1] != "112233" {
		t.Errorf("dk1 not applied: %q", theme.Colors[ThemeDark1])
	}
	// Untouched slots survive exactly.
	if theme.Colors[ThemeAccent2] != "ED7D31" {
		t.Errorf("accent2 changed unexpectedly: %q", theme.Colors[ThemeAccent2])
	}
	if theme.Colors[ThemeLight1] != "FFFFFF" {
		t.Errorf("lt1 changed unexpectedly: %q", theme.Colors[ThemeLight1])
	}

	// dk1 is now a plain srgbClr, not a sysClr.
	themeXML, _ := zipPart(t, out, "word/theme/theme1.xml")
	if !strings.Contains(themeXML, `<a:dk1><a:srgbClr val="112233"/></a:dk1>`) {
		t.Errorf("dk1 not converted to srgbClr: %s", themeXML)
	}

	// fmtScheme and everything else byte-identical.
	if !strings.Contains(themeXML, "<a:fmtScheme") || !strings.Contains(themeXML, "gradFill") {
		t.Errorf("fmtScheme lost: %s", themeXML)
	}
	origDoc, _ := zipPart(t, src, "word/document.xml")
	outDoc, _ := zipPart(t, out, "word/document.xml")
	if origDoc != outDoc {
		t.Fatal("word/document.xml must be byte-identical")
	}
}

func TestApplyTheme_Fonts(t *testing.T) {
	src := buildThemeTestDocx(t)
	out, err := ApplyTheme(src, ThemePatch{MajorLatinFont: "Georgia", MinorLatinFont: "Verdana"})
	if err != nil {
		t.Fatal(err)
	}
	theme, err := ExtractTheme(out)
	if err != nil {
		t.Fatal(err)
	}
	if theme.MajorLatin != "Georgia" {
		t.Errorf("major font not applied: %q", theme.MajorLatin)
	}
	if theme.MinorLatin != "Verdana" {
		t.Errorf("minor font not applied: %q", theme.MinorLatin)
	}
	// Colors untouched.
	if theme.Colors[ThemeAccent1] != "4472C4" {
		t.Errorf("colors changed unexpectedly: %q", theme.Colors[ThemeAccent1])
	}
}

func TestApplyTheme_ColorsAndFontsTogether(t *testing.T) {
	src := buildThemeTestDocx(t)
	out, err := ApplyTheme(src, ThemePatch{
		Colors:         map[ThemeColorSlot]string{ThemeAccent3: "00FF00"},
		MajorLatinFont: "Georgia",
	})
	if err != nil {
		t.Fatal(err)
	}
	theme, err := ExtractTheme(out)
	if err != nil {
		t.Fatal(err)
	}
	if theme.Colors[ThemeAccent3] != "00FF00" || theme.MajorLatin != "Georgia" {
		t.Fatalf("combined patch failed: %+v", theme)
	}
	if theme.MinorLatin != "Calibri" {
		t.Errorf("minor font should be untouched: %q", theme.MinorLatin)
	}
}

func TestApplyTheme_RejectsBadInput(t *testing.T) {
	src := buildThemeTestDocx(t)
	cases := []struct {
		name string
		fn   func() error
	}{
		{"empty patch", func() error {
			_, err := ApplyTheme(src, ThemePatch{})
			return err
		}},
		{"invalid hex color", func() error {
			_, err := ApplyTheme(src, ThemePatch{Colors: map[ThemeColorSlot]string{ThemeAccent1: "notacolor"}})
			return err
		}},
		{"unknown slot", func() error {
			_, err := ApplyTheme(src, ThemePatch{Colors: map[ThemeColorSlot]string{"bogus": "FF0000"}})
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

func TestApplyTheme_NoThemePart(t *testing.T) {
	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	parts := map[string]string{
		"[Content_Types].xml": `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="x"/></Types>`,
		"word/document.xml":   `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`,
	}
	for name, content := range parts {
		w, _ := zw.Create(name)
		w.Write([]byte(content)) //nolint:errcheck
	}
	zw.Close() //nolint:errcheck
	if _, err := ApplyTheme(buf.Bytes(), ThemePatch{Colors: map[ThemeColorSlot]string{ThemeAccent1: "FF0000"}}); err == nil {
		t.Fatal("expected an error for a docx with no theme part")
	}
}

// TestApplyTheme_OpensWithPythonDocxOxml independently validates via
// python-docx's low-level docx.oxml/lxml access — python-docx has NO theme
// API at all (confirmed: no `theme` attribute anywhere on Document or
// DocumentPart), the same gap footnotes/charts had. Opens the file with
// python-docx's real OPC package/relationship machinery (proves zip/rels/
// content-types integrity independently of this package's own writer),
// resolves the theme relationship, then lxml parses the raw theme part.
func TestApplyTheme_OpensWithPythonDocxOxml(t *testing.T) {
	py, err := exec.LookPath("python3")
	if err != nil {
		t.Skip("python3 not available")
	}
	if err := exec.Command(py, "-c", "import docx, lxml.etree").Run(); err != nil {
		t.Skip("python-docx / lxml not installed")
	}

	out, err := ApplyTheme(buildThemeTestDocx(t), ThemePatch{
		Colors:         map[ThemeColorSlot]string{ThemeAccent1: "FF0000", ThemeDark1: "112233"},
		MajorLatinFont: "Georgia",
	})
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
assert d.paragraphs[0].text == "Hello theme.", d.paragraphs[0].text

theme_rel = None
for rel in d.part.rels.values():
    if rel.reltype.endswith("/theme"):
        theme_rel = rel
        break
assert theme_rel is not None, "no theme relationship found by python-docx"
theme_xml = theme_rel.target_part.blob

ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
root = etree.fromstring(theme_xml)
accent1 = root.find(".//a:clrScheme/a:accent1/a:srgbClr", ns)
assert accent1.get("val") == "FF0000", accent1.get("val")
dk1 = root.find(".//a:clrScheme/a:dk1/a:srgbClr", ns)
assert dk1 is not None and dk1.get("val") == "112233", etree.tostring(root.find(".//a:dk1", ns))
major_latin = root.find(".//a:fontScheme/a:majorFont/a:latin", ns)
assert major_latin.get("typeface") == "Georgia", major_latin.get("typeface")
# fmtScheme survived.
assert root.find(".//a:fmtScheme", ns) is not None

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
