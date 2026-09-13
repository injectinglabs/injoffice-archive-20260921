package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func richFixture(strict bool, runs string) map[string]string {
	p := nativeMutationFixture(strict)
	ns := spreadsheetMLTransitional
	if strict {
		ns = spreadsheetMLStrict
	}
	p["Sheets/s1.xml"] = `<worksheet xmlns="` + ns + `"><sheetData><row r="1"><c r="A1" t="inlineStr"><is>` + runs + `</is></c></row></sheetData></worksheet>`
	return p
}

const richPositiveRuns = `<r><rPr><b/><color rgb="FF0000FF"/><rFont val="Calibri"/><sz val="11"/><family val="2"/><vertAlign val="baseline"/></rPr><t xml:space="preserve"> bold &amp; </t></r><r><rPr><i val="false"/><color rgb="FFFF0000"/></rPr><t>normal</t></r><r><t> inherited</t></r>`

func richPreviewFixture(t *testing.T, p map[string]string) NativeRichTextPreviewV1 {
	t.Helper()
	b := buildZip(t, p)
	before := bytes.Clone(b)
	w, e := ExtractNativeWorkbookV2(b)
	if e != nil {
		t.Fatal(e)
	}
	pkg, e := openNativeWorkbookPackage(b)
	if e != nil {
		t.Fatal(e)
	}
	r := previewNativeRichText(pkg, w)
	if !bytes.Equal(b, before) {
		t.Fatal("source changed")
	}
	return r
}
func TestNativeRichTextDirectAndInherited(t *testing.T) {
	for _, strict := range []bool{false, true} {
		r := richPreviewFixture(t, richFixture(strict, richPositiveRuns))
		if len(r.Cells) != 1 {
			t.Fatalf("entries %#v", r)
		}
		c := r.Cells[0]
		if c.Status != "available" || len(c.Runs) != 3 {
			t.Fatalf("not available %#v", c)
		}
		if c.Text != " bold & normal inherited" || c.Runs[0].Bold == nil || !*c.Runs[0].Bold || c.Runs[1].Bold != nil || c.Runs[1].Italic == nil || *c.Runs[1].Italic || c.Runs[2].Properties != "cell-inherited" {
			t.Fatal("run source properties or text lost")
		}
		if c.Runs[0].FontColor == nil || *c.Runs[0].FontColor != "#0000FF" || len(c.Runs[0].Omitted) != 2 || len(c.Warnings) != 1 {
			t.Fatal("direct color or omissions lost")
		}
	}
}
func TestNativeRichTextRawRunRefusals(t *testing.T) {
	for name, patch := range map[string]string{
		"underline": `<u val="single"/>`, "superscript": `<vertAlign val="superscript"/>`, "scheme": `<scheme val="major"/>`, "theme": `<color theme="1"/>`, "translucent": `<color rgb="80123456"/>`, "unknown": `<unknown/>`, "foreign": `<b xmlns="urn:foreign"/>`, "duplicate": `<b/><b/>`, "effect": `<strike/>`, "emptyboolean": `<b val=""/>`, "propertytext": `<b>text</b>`,
	} {
		t.Run(name, func(t *testing.T) {
			raw := `<is xmlns="` + spreadsheetMLTransitional + `"><r><rPr>` + patch + `</rPr><t>text</t></r></is>`
			root, e := parsePreviewXML([]byte(raw))
			if e != nil {
				t.Fatal(e)
			}
			if runs, why := nativeRichRuns(root, spreadsheetMLTransitional); why == "" || runs != nil {
				t.Fatal("unqualified run accepted")
			}
		})
	}
}
func TestNativeRichTextWholeCellOmission(t *testing.T) {
	for name, edit := range map[string]func(map[string]string){
		"metadata": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `r="A1"`, `r="A1" cm="1"`, 1)
		},
		"tableparts": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `</worksheet>`, `<tableParts count="1"/></worksheet>`, 1)
		},
		"wrapped": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `r="A1"`, `r="A1" s="1"`, 1)
		},
		"group": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `</row>`, `<c r="B1"><f t="array" ref="B1:B2">1</f><v>1</v></c></row>`, 1)
		},
		"underline": func(p map[string]string) {
			p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `<b/>`, `<b/><u val="single"/>`, 1)
		},
	} {
		t.Run(name, func(t *testing.T) {
			p := richFixture(false, richPositiveRuns)
			edit(p)
			r := richPreviewFixture(t, p)
			if len(r.Cells) != 1 || r.Cells[0].Status != "omitted" || len(r.Cells[0].Runs) != 0 || r.Cells[0].Text != " bold & normal inherited" {
				t.Fatalf("partial style or lost text %#v", r)
			}
		})
	}
}
func TestNativeRichTextBounds(t *testing.T) {
	for _, runs := range []string{strings.Repeat(`<r><t>x</t></r>`, 65), `<r><t>` + strings.Repeat("x", 2049) + `</t></r>`} {
		root, e := parsePreviewXML([]byte(`<is xmlns="` + spreadsheetMLTransitional + `">` + runs + `</is>`))
		if e != nil {
			t.Fatal(e)
		}
		if _, why := nativeRichRuns(root, spreadsheetMLTransitional); why == "" {
			t.Fatal("budget accepted")
		}
	}
}

func TestNativeRichTextEnclosingBudget(t *testing.T) {
	base := map[string]any{"existing": make([]int, 199998)}
	if !nativeRichTextEnvelopeFits(base) {
		t.Fatal("valid older envelope refused")
	}
	base["rich_text"] = NativeRichTextPreviewV1{Cells: []NativeRichTextCellV1{}, Warnings: []string{}}
	if nativeRichTextEnvelopeFits(base) {
		t.Fatal("supplement overflow accepted")
	}
	delete(base, "rich_text")
	if !nativeRichTextEnvelopeFits(base) {
		t.Fatal("older fields changed")
	}
}

func TestNativeRichTextSharedOwnershipAndNoOps(t *testing.T) {
	p := richFixture(false, richPositiveRuns)
	p["Sheets/s2.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData/></worksheet>`
	p["Meta/Strings.xml"] = `<sst xmlns="` + spreadsheetMLTransitional + `" count="2" uniqueCount="1"><si><r><rPr><b val="0"/><i/><strike val="false"/><u val="none"/><color rgb="ff112233"/></rPr><t>_x0041_ &amp; </t></r><r><t>text</t></r></si></sst>`
	p["Sheets/s1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>0</v></c></row></sheetData></worksheet>`
	r := richPreviewFixture(t, p)
	if len(r.Cells) != 2 {
		t.Fatal("missing shared cells")
	}
	for _, c := range r.Cells {
		if c.Status != "available" || c.SourcePart != "Meta/Strings.xml" || c.SharedIndex != "0" || c.Text != "A & text" || c.Runs[0].Bold == nil || *c.Runs[0].Bold || c.Runs[0].Italic == nil || !*c.Runs[0].Italic || len(c.Runs[0].Omitted) != 2 {
			t.Fatalf("shared evidence mismatch %#v", c)
		}
	}
	r.Cells[0].Runs[0].Text = "changed"
	if r.Cells[1].Runs[0].Text != "A & " {
		t.Fatal("per-cell evidence aliases shared runs")
	}
}
func TestNativeRichTextNoMisplacedCellOwnership(t *testing.T) {
	p := richFixture(false, richPositiveRuns)
	source := buildZip(t, p)
	w, e := ExtractNativeWorkbookV2(source)
	if e != nil {
		t.Fatal(e)
	}
	pkg, e := openNativeWorkbookPackage(source)
	if e != nil {
		t.Fatal(e)
	}
	// Exercise the supplemental raw guard independently of strict extraction.
	raw := string(pkg.files["Sheets/s1.xml"])
	raw = strings.Replace(raw, "<sheetData>", "<wrapper><sheetData>", 1)
	raw = strings.Replace(raw, "</sheetData>", "</sheetData></wrapper>", 1)
	pkg.files["Sheets/s1.xml"] = []byte(raw)
	got := previewNativeRichText(pkg, w)
	if got.Cells[0].Status != "omitted" {
		t.Fatal("non-root-owned sheetData accepted")
	}
}

func TestNativeRichTextExplicitEmptyTableContainer(t *testing.T) {
	p := richFixture(false, richPositiveRuns)
	p["Sheets/s1.xml"] = strings.Replace(p["Sheets/s1.xml"], `</worksheet>`, `<tableParts count="0"/></worksheet>`, 1)
	if r := richPreviewFixture(t, p); r.Cells[0].Status != "available" {
		t.Fatal("explicit empty table container blocked runs")
	}
}

func TestNativeRichTextPublicObjectsSupplement(t *testing.T) {
	p := richFixture(false, richPositiveRuns)
	p["Charts/chart1.xml"] = previewChartFixture()
	b := buildZip(t, p)
	before := bytes.Clone(b)
	objects, e := InspectNativeWorkbookObjectsV1(b)
	if e != nil {
		t.Fatal(e)
	}
	if objects.RichText == nil || len(objects.RichText.Cells) != 1 || objects.RichText.Cells[0].Status != "available" {
		t.Fatal("public supplement missing")
	}
	if objects.PackageSHA256 != nativeWorkbookDigest(b) || !bytes.Equal(before, b) {
		t.Fatal("public source changed")
	}
}

func TestNativeRichTextEnclosingByteBudget(t *testing.T) {
	// {"existing":""} contributes15 bytes to the serialized envelope.
	base := map[string]any{"existing": strings.Repeat("x", 8*1024*1024-15)}
	if !nativeRichTextEnvelopeFits(base) {
		t.Fatal("existing at-limit envelope refused")
	}
	base["rich_text"] = NativeRichTextPreviewV1{Cells: []NativeRichTextCellV1{}, Warnings: []string{}}
	if nativeRichTextEnvelopeFits(base) {
		t.Fatal("new supplement exceeded output byte ceiling")
	}
	delete(base, "rich_text")
	if !nativeRichTextEnvelopeFits(base) {
		t.Fatal("older envelope changed")
	}
}
