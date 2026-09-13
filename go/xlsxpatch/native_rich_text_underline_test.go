package xlsxpatch

import (
	"testing"
)

func TestNativeRichUnderlineSourceDeclarations(t *testing.T) {
	for _, strict := range []bool{false, true} {
		r := richPreviewFixture(t, richFixture(strict, `<r><rPr><u val="single"/></rPr><t>explicit</t></r><r><rPr><u/></rPr><t>default</t></r><r><rPr><u val="none"/></rPr><t>none</t></r><r><rPr/><t>absent</t></r><r><t>inherited</t></r>`))
		if len(r.Cells) != 1 || r.Cells[0].Status != "available" {
			t.Fatalf("missing source runs: %#v", r)
		}
		runs := r.Cells[0].Runs
		for i, want := range []struct{ value, origin string }{{"single", "explicit-val"}, {"single", "default-val"}, {"none", "explicit-val"}, {"", ""}, {"", ""}} {
			if runs[i].Underline != want.value || runs[i].UnderlineOrigin != want.origin {
				t.Fatalf("run %d: %#v", i, runs[i])
			}
		}
		if runs[3].Properties != "direct" || runs[4].Properties != "cell-inherited" {
			t.Fatal("missing property ownership")
		}
	}
}

func TestNativeRichUnderlineAtomicRefusals(t *testing.T) {
	for _, strict := range []bool{false, true} {
		ns := spreadsheetMLTransitional
		if strict {
			ns = spreadsheetMLStrict
		}
		for _, bad := range []string{`<u val="double"/>`, `<u val="singleAccounting"/>`, `<u val="doubleAccounting"/>`, `<u val=""/>`, `<u val="Single"/>`, `<u val=" single"/>`, `<u val="bad"/>`, `<u/><u/>`, `<u xmlns="urn:foreign"/>`, `<u xmlns:f="urn:f" f:val="single"/>`, `<u> </u>`, `<u><b/></u>`, `<u val="single" val="none"/>`} {
			raw := `<is xmlns="` + ns + `"><r><rPr><u/></rPr><t>good</t></r><r><rPr>` + bad + `</rPr><t>bad</t></r></is>`
			root, err := parsePreviewXML([]byte(raw))
			if err != nil {
				continue
			} // malformed duplicate attributes can refuse before projection
			if runs, why := nativeRichRuns(root, ns); why == "" || runs != nil {
				t.Fatalf("accepted partial styling: %s", bad)
			}
		}
		r := richPreviewFixture(t, richFixture(strict, `<r><rPr><u/></rPr><t>good</t></r><r><rPr><u val="double"/></rPr><t>bad</t></r>`))
		if len(r.Cells) != 1 || r.Cells[0].Status != "omitted" || r.Cells[0].Runs != nil || r.Cells[0].Text != "goodbad" {
			t.Fatalf("lost plaintext or partial styling: %#v", r)
		}
	}
}

func TestNativeRichUnderlinePublicSupplement(t *testing.T) {
	p := richFixture(false, `<r><rPr><u/></rPr><t>underlined</t></r><r><t>plain</t></r>`)
	p["Charts/chart1.xml"] = previewChartFixture()
	b := buildZip(t, p)
	objects, err := InspectNativeWorkbookObjectsV1(b)
	if err != nil {
		t.Fatal(err)
	}
	if objects.RichText == nil || len(objects.RichText.Cells) != 1 || len(objects.RichText.Cells[0].Runs) != 2 || objects.RichText.Cells[0].Runs[0].UnderlineOrigin != "default-val" || objects.RichText.Cells[0].Runs[1].Underline != "" || objects.PackageSHA256 != nativeWorkbookDigest(b) {
		t.Fatalf("public underline/source lost: %#v", objects.RichText)
	}
}
