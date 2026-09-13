package xlsxpatch

import (
	"strings"
	"testing"
)

const richSchemeTheme = `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"><a:themeElements><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Example"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`
const richSchemeRuns = `<r><rPr><rFont val="Arial"/><scheme val="major"/></rPr><t>Major scheme</t></r><r><t> </t></r><r><rPr><rFont val="Times New Roman"/><scheme val="minor"/></rPr><t>Minor scheme</t></r>`

func richSchemeFixture() map[string]string {
	p := richFixture(false, richSchemeRuns)
	p["Charts/chart1.xml"] = previewChartFixture()
	p["[Content_Types].xml"] = strings.Replace(p["[Content_Types].xml"], `</Types>`, `<Override PartName="/Themes/office.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
	p["Book/_rels/Workbook.xml.rels"] = strings.Replace(p["Book/_rels/Workbook.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relTypeThemeTransitional+`" Target="../Themes/office.xml"/></Relationships>`, 1)
	p["Themes/office.xml"] = richSchemeTheme
	return p
}
func TestNativeRichSchemeSourceJoin(t *testing.T) {
	p := richSchemeFixture()
	b := buildZip(t, p)
	objects, err := InspectNativeWorkbookObjectsV1(b)
	if err != nil {
		t.Fatal(err)
	}
	if objects.RichText == nil || len(objects.RichText.Cells) != 1 {
		t.Fatalf("missing supplement: %#v", objects.RichText)
	}
	c := objects.RichText.Cells[0]
	if c.Status != "available" {
		t.Fatalf("omitted: %#v", c)
	}
	for _, i := range []int{0, 2} {
		r := c.Runs[i]
		want, declared, scheme := "Cambria", "Arial", "major"
		if i == 2 {
			want, declared, scheme = "Calibri", "Times New Roman", "minor"
		}
		if r.FontName == nil || *r.FontName != want || r.DeclaredFontName != declared || r.FontScheme != scheme || r.ThemePart != "Themes/office.xml" || r.ThemeSHA256 != nativeWorkbookDigest([]byte(richSchemeTheme)) {
			t.Fatalf("source font evidence: %#v", r)
		}
	}
	if c.Runs[1].FontScheme != "" || c.Runs[1].FontName != nil || objects.PackageSHA256 != nativeWorkbookDigest(b) {
		t.Fatal("scheme leaked or source identity lost")
	}
}
func TestNativeRichSchemeAtomicRefusals(t *testing.T) {
	for name, bad := range map[string]string{
		"duplicate-latin":     strings.Replace(richSchemeTheme, `<a:latin typeface="Cambria"/>`, `<a:latin typeface="Cambria"/><a:latin typeface="Arial"/>`, 1),
		"duplicate-scheme":    strings.Replace(richSchemeTheme, `</a:themeElements>`, `<a:fontScheme name="other"/></a:themeElements>`, 1),
		"foreign-latin":       strings.Replace(richSchemeTheme, `<a:latin typeface="Cambria"/>`, `<latin xmlns="urn:foreign" typeface="Cambria"/>`, 1),
		"unknown-theme":       strings.Replace(richSchemeTheme, `</a:theme>`, `<a:extLst/></a:theme>`, 1),
		"nonempty-east-asian": strings.Replace(richSchemeTheme, `<a:ea typeface=""/>`, `<a:ea typeface="Other"/>`, 1),
		"font-metadata":       strings.Replace(richSchemeTheme, `typeface="Cambria"`, `typeface="Cambria" charset="0"`, 1),
	} {
		t.Run(name, func(t *testing.T) {
			p := richSchemeFixture()
			p["Themes/office.xml"] = bad
			assertRichSchemeOmitted(t, p)
		})
	}
	for name, bad := range map[string]string{
		"non-ascii":          strings.Replace(richSchemeRuns, "Minor scheme", "世界", 1),
		"absent-declared":    strings.Replace(richSchemeRuns, `<rFont val="Arial"/>`, "", 1),
		"none":               strings.Replace(richSchemeRuns, `val="major"`, `val="none"`, 1),
		"duplicate-property": strings.Replace(richSchemeRuns, `<scheme val="major"/>`, `<scheme val="major"/><scheme val="minor"/>`, 1),
	} {
		t.Run(name, func(t *testing.T) {
			p := richSchemeFixture()
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], richSchemeRuns, bad, 1)
			assertRichSchemeOmitted(t, p)
		})
	}
}
func assertRichSchemeOmitted(t *testing.T, p map[string]string) {
	t.Helper()
	b := buildZip(t, p)
	w, err := ExtractNativeWorkbookV2(b)
	if err != nil {
		return
	}
	pkg, err := openNativeWorkbookPackage(b)
	if err != nil {
		t.Fatal(err)
	}
	r := previewNativeRichText(pkg, w)
	if len(r.Cells) != 1 || r.Cells[0].Status != "omitted" || len(r.Cells[0].Runs) != 0 || r.Cells[0].Text == "" {
		t.Fatalf("partial style or missing text: %#v", r)
	}
}

func TestNativeRichSchemeMissingAndStrictTheme(t *testing.T) {
	for _, strict := range []bool{false, true} {
		p := richFixture(strict, richSchemeRuns)
		assertRichSchemeOmitted(t, p)
	}
	for name, edit := range map[string]func(map[string]string){
		"missing-part": func(p map[string]string) { delete(p, "Themes/office.xml") },
		"external": func(p map[string]string) {
			p["Book/_rels/Workbook.xml.rels"] = strings.Replace(p["Book/_rels/Workbook.xml.rels"], `Target="../Themes/office.xml"`, `Target="https://example.invalid/theme.xml" TargetMode="External"`, 1)
		},
		"wrong-content-type": func(p map[string]string) {
			p["[Content_Types].xml"] = strings.Replace(p["[Content_Types].xml"], themePartContentType, "application/xml", 1)
		},
		"duplicate-relationship": func(p map[string]string) {
			p["Book/_rels/Workbook.xml.rels"] = strings.Replace(p["Book/_rels/Workbook.xml.rels"], `</Relationships>`, `<Relationship Id="secondTheme" Type="`+relTypeThemeTransitional+`" Target="../Themes/office.xml"/></Relationships>`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) { p := richSchemeFixture(); edit(p); assertRichSchemeOmitted(t, p) })
	}
}
