package docxpatch

import (
	"encoding/json"
	"strings"
	"testing"
	"unicode/utf8"
)

const testMathStyles = `<w:styles xmlns:w="` + testW + `"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Caption"><w:name w:val="caption"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:color w:val="4F81BD"/><w:sz w:val="18"/></w:rPr></w:style></w:styles>`

func nativeApproximateEquationSource(t *testing.T, body string) []byte {
	entries := resolvedStylesTestParts(testMathStyles)
	entries["word/document.xml"] = nativeMutationMain(body)
	return buildNativeDOCX(t, nativeEntries(entries))
}

func nativeMathParagraph(pStyle, math string) string {
	pPr := ""
	if pStyle != "" {
		pPr = `<w:pPr><w:pStyle w:val="` + pStyle + `"/></w:pPr>`
	}
	return `<w:p>` + pPr + `<m:oMathPara xmlns:m="` + nativePartialMathNamespace(testW) + `">` + math + `</m:oMathPara></w:p>`
}

const testMathRPr = `<w:rPr><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/></w:rPr>`

func mathRun(text string) string {
	return `<m:r><m:rPr><m:sty m:val="bi"/></m:rPr>` + testMathRPr + `<m:t>` + text + `</m:t></m:r>`
}

func mathCtrl() string { return `<m:ctrlPr>` + testMathRPr + `</m:ctrlPr>` }

// The benchmark corpus equation: (x+a)^n = Σ_{k=0}^{n} (n choose k) x^k a^{n-k}.
func testBinomialTheorem() string {
	return `<m:oMath><m:sSup><m:sSupPr>` + mathCtrl() + `</m:sSupPr><m:e><m:d><m:dPr>` + mathCtrl() + `</m:dPr><m:e>` + mathRun("x+a") + `</m:e></m:d></m:e><m:sup>` + mathRun("n") + `</m:sup></m:sSup>` + mathRun("=") +
		`<m:nary><m:naryPr><m:chr m:val="∑"/><m:grow m:val="1"/>` + mathCtrl() + `</m:naryPr><m:sub>` + mathRun("k=0") + `</m:sub><m:sup>` + mathRun("n") + `</m:sup><m:e><m:d><m:dPr>` + mathCtrl() + `</m:dPr><m:e><m:f><m:fPr><m:type m:val="noBar"/>` + mathCtrl() + `</m:fPr><m:num>` + mathRun("n") + `</m:num><m:den>` + mathRun("k") + `</m:den></m:f></m:e></m:d>` +
		`<m:sSup><m:sSupPr>` + mathCtrl() + `</m:sSupPr><m:e>` + mathRun("x") + `</m:e><m:sup>` + mathRun("k") + `</m:sup></m:sSup><m:sSup><m:sSupPr>` + mathCtrl() + `</m:sSupPr><m:e>` + mathRun("a") + `</m:e><m:sup>` + mathRun("n-k") + `</m:sup></m:sSup></m:e></m:nary></m:oMath>`
}

func inspectEquations(t *testing.T, source []byte) (*NativeDocumentV1, *NativeApproximateEquationsV1) {
	t.Helper()
	before := append([]byte(nil), source...)
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		t.Fatal(err)
	}
	out, err := InspectNativeApproximateEquationsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if string(before) != string(source) {
		t.Fatal("inspection changed caller bytes")
	}
	return doc, out
}

func TestApproximateEquationsBinomialTheorem(t *testing.T) {
	source := nativeApproximateEquationSource(t, `<w:p><w:pPr><w:pStyle w:val="Caption"/></w:pPr><w:r><w:t>scientific</w:t></w:r></w:p>`+nativeMathParagraph("Caption", testBinomialTheorem()))
	doc, out := inspectEquations(t, source)
	if out == nil || out.Protocol != NativeApproximateEquationsProtocol || out.Policy != NativeApproximateEquationPolicy || out.PackageSHA256 != doc.Source.PackageSHA256 || len(out.Items) != 1 || out.OmittedCount != 0 {
		t.Fatalf("unexpected sidecar: %#v", out)
	}
	item := out.Items[0]
	p := doc.Body.Blocks[1].Paragraph
	if item.Status != "supported" || item.ParagraphID != p.ID || !item.Display || item.Justification != "centerGroup" || len(item.Lines) != 1 || item.Reason != "" || item.ID != "approximate-equation:"+strings.TrimPrefix(p.ID, "paragraph:")+":1" {
		t.Fatalf("unexpected item: %#v", item)
	}
	// Strict refusal is unchanged and joined.
	found := false
	for _, d := range doc.Unsupported {
		if d.Code == "UNMODELED_PARAGRAPH_CONTENT" && d.ScopeID == p.ID && d.Anchor != nil && d.Anchor.Path == item.Anchor.Path {
			found = d.ID == item.DiagnosticIDs[0] && d.Anchor.XMLSHA256 == item.Anchor.XMLSHA256
		}
	}
	if !found || p.EditPolicy.Mode != "read-only" || len(p.Runs) != 0 {
		t.Fatalf("strict diagnostic or read-only policy changed: %#v", p)
	}
	row := item.Lines[0]
	if row.Kind != "row" || len(row.Children) != 3 {
		t.Fatalf("unexpected top row: %#v", row)
	}
	sup, equals, nary := row.Children[0], row.Children[1], row.Children[2]
	if sup.Kind != "superscript" || len(sup.Children) != 2 || sup.Run == nil || sup.Run.FontFamily != "Cambria Math" || sup.Run.Bold != true || sup.Run.Italic != false || sup.Run.FontSizeHalfPoints != 18 || sup.Run.Color != "4F81BD" {
		t.Fatalf("unexpected superscript: %#v", sup)
	}
	delimiter := sup.Children[0].Children[0]
	if delimiter.Kind != "delimiter" || delimiter.BegChr != nil || delimiter.EndChr != nil || len(delimiter.Children) != 1 {
		t.Fatalf("unexpected delimiter: %#v", delimiter)
	}
	text := delimiter.Children[0].Children[0]
	if text.Kind != "text" || text.Text != "x+a" || text.Run == nil || !text.Run.Bold || !text.Run.Italic || text.Run.Style != "bi" || text.Run.FontSizeHalfPoints != 18 {
		t.Fatalf("unexpected math text: %#v", text)
	}
	if equals.Kind != "text" || equals.Text != "=" {
		t.Fatalf("unexpected operator run: %#v", equals)
	}
	if nary.Kind != "nary" || nary.Chr != "∑" || !nary.Grow || nary.SubHide || nary.SupHide || nary.LimitLocation != "" || len(nary.Children) != 3 {
		t.Fatalf("unexpected nary: %#v", nary)
	}
	body := nary.Children[2]
	if len(body.Children) != 3 || body.Children[0].Kind != "delimiter" {
		t.Fatalf("unexpected nary body: %#v", body)
	}
	binomial := body.Children[0].Children[0].Children[0]
	if binomial.Kind != "fraction" || binomial.Bar == nil || *binomial.Bar || len(binomial.Children) != 2 {
		t.Fatalf("unexpected binomial: %#v", binomial)
	}
	wantFonts := map[string]bool{"Cambria Math|700|normal": true, "Cambria Math|700|italic": true, "Calibri|700|normal": true, "Calibri|700|italic": true}
	for _, f := range out.FontRequests {
		delete(wantFonts, f.Family+"|"+func() string {
			if f.Weight == 700 {
				return "700"
			}
			return "400"
		}()+"|"+f.Style)
	}
	if len(wantFonts) != 0 || len(out.FontRequests) != 4 {
		t.Fatalf("unexpected font requests: %#v (missing %v)", out.FontRequests, wantFonts)
	}
	encoded, err := json.Marshal(out)
	if err != nil || !strings.Contains(string(encoded), `"bar":false`) || strings.Contains(string(encoded), `"beg_chr"`) {
		t.Fatalf("unexpected wire: %s", encoded)
	}
}

func TestApproximateEquationsConstructs(t *testing.T) {
	plain := func(text string) string { return `<m:r>` + testMathRPr + `<m:t>` + text + `</m:t></m:r>` }
	cases := []struct {
		name, math string
		check      func(t *testing.T, n NativeApproximateMathNodeV1)
	}{
		{"fraction with bar", `<m:f><m:num>` + plain("a") + `</m:num><m:den>` + plain("b") + `</m:den></m:f>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "fraction" || n.Bar == nil || !*n.Bar || len(n.Children) != 2 || n.Run != nil {
				t.Fatalf("fraction: %#v", n)
			}
		}},
		{"subscript", `<m:sSub><m:e>` + plain("x") + `</m:e><m:sub>` + plain("1") + `</m:sub></m:sSub>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "subscript" || len(n.Children) != 2 {
				t.Fatalf("subscript: %#v", n)
			}
		}},
		{"sub-superscript", `<m:sSubSup><m:e>` + plain("x") + `</m:e><m:sub>` + plain("1") + `</m:sub><m:sup>` + plain("2") + `</m:sup></m:sSubSup>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "subsuperscript" || len(n.Children) != 3 {
				t.Fatalf("subsuperscript: %#v", n)
			}
		}},
		{"nary side limits with hidden lower", `<m:nary><m:naryPr><m:chr m:val="∫"/><m:limLoc m:val="subSup"/><m:subHide m:val="1"/></m:naryPr><m:sub/><m:sup>` + plain("1") + `</m:sup><m:e>` + plain("f") + `</m:e></m:nary>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "nary" || n.Chr != "∫" || n.LimitLocation != "subSup" || !n.SubHide || n.SupHide || len(n.Children) != 3 || len(n.Children[0].Children) != 0 {
				t.Fatalf("nary: %#v", n)
			}
		}},
		{"delimiter with explicit characters", `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/><m:sepChr m:val=""/></m:dPr><m:e>` + plain("a") + `</m:e><m:e>` + plain("b") + `</m:e></m:d>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "delimiter" || n.BegChr == nil || *n.BegChr != "[" || n.EndChr == nil || *n.EndChr != "]" || n.SepChr == nil || *n.SepChr != "" || len(n.Children) != 2 {
				t.Fatalf("delimiter: %#v", n)
			}
		}},
		{"radical with hidden degree", `<m:rad><m:radPr><m:degHide m:val="on"/></m:radPr><m:deg/><m:e>` + plain("2") + `</m:e></m:rad>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "radical" || !n.DegreeHide || len(n.Children) != 2 {
				t.Fatalf("radical: %#v", n)
			}
		}},
		{"function", `<m:func><m:fName><m:r><m:rPr><m:sty m:val="p"/></m:rPr>` + testMathRPr + `<m:t>sin</m:t></m:r></m:fName><m:e>` + plain("x") + `</m:e></m:func>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			name := n.Children[0].Children[0]
			if n.Kind != "function" || len(n.Children) != 2 || name.Run == nil || name.Run.Style != "p" || name.Run.Italic || name.Run.Bold {
				t.Fatalf("function: %#v", n)
			}
		}},
		{"bar on top", `<m:bar><m:barPr><m:pos m:val="top"/></m:barPr><m:e>` + plain("x") + `</m:e></m:bar>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "bar" || n.Position != "top" || len(n.Children) != 1 {
				t.Fatalf("bar: %#v", n)
			}
		}},
		{"accent", `<m:acc><m:accPr><m:chr m:val="̃"/></m:accPr><m:e>` + plain("x") + `</m:e></m:acc>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "accent" || n.Chr != "̃" || len(n.Children) != 1 {
				t.Fatalf("accent: %#v", n)
			}
		}},
		{"lower limit", `<m:limLow><m:e><m:r><m:rPr><m:sty m:val="p"/></m:rPr>` + testMathRPr + `<m:t>lim</m:t></m:r></m:e><m:lim>` + plain("x→0") + `</m:lim></m:limLow>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "limit-lower" || len(n.Children) != 2 {
				t.Fatalf("limit: %#v", n)
			}
		}},
		{"normal text run", `<m:r><m:rPr><m:nor/><m:scr m:val="roman"/></m:rPr>` + testMathRPr + `<m:t xml:space="preserve"> if </m:t></m:r>`, func(t *testing.T, n NativeApproximateMathNodeV1) {
			if n.Kind != "text" || n.Text != " if " || n.Run == nil || !n.Run.Normal || n.Run.Style != "" || n.Run.Italic || n.Run.FontFamily != "Cambria Math" || n.Run.FontSizeHalfPoints != 22 {
				t.Fatalf("normal run: %#v", n)
			}
		}},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			_, out := inspectEquations(t, nativeApproximateEquationSource(t, nativeMathParagraph("", `<m:oMathParaPr><m:jc m:val="left"/></m:oMathParaPr><m:oMath>`+c.math+`</m:oMath>`)))
			if out == nil || len(out.Items) != 1 || out.Items[0].Status != "supported" || out.Items[0].Justification != "left" {
				t.Fatalf("unexpected sidecar: %+v", out)
			}
			row := out.Items[0].Lines[0]
			if len(row.Children) != 1 {
				t.Fatalf("unexpected row: %#v", row)
			}
			c.check(t, row.Children[0])
		})
	}
}

func TestApproximateEquationsInlineAndRefusals(t *testing.T) {
	plain := func(text string) string { return `<m:r>` + testMathRPr + `<m:t>` + text + `</m:t></m:r>` }
	ns := nativePartialMathNamespace(testW)
	inline := `<w:p><w:r><w:t>Let </w:t></w:r><m:oMath xmlns:m="` + ns + `">` + plain("x") + `</m:oMath><w:r><w:t> hold.</w:t></w:r></w:p>`
	doc, out := inspectEquations(t, nativeApproximateEquationSource(t, inline))
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "supported" || out.Items[0].Display || out.Items[0].Justification != "" || len(doc.Body.Blocks[0].Paragraph.Runs) != 2 {
		t.Fatalf("inline equation: %+v", out)
	}
	if text := out.Items[0].Lines[0].Children[0]; text.Run == nil || text.Run.FontSizeHalfPoints != 22 || text.Run.Bold || text.Run.Italic || text.Run.Style != "" {
		t.Fatalf("inline run inherits document defaults: %#v", text)
	}
	deep := plain("x")
	for i := 0; i < 20; i++ {
		deep = `<m:d><m:e>` + deep + `</m:e></m:d>`
	}
	wide := strings.Builder{}
	for i := 0; i < 600; i++ {
		wide.WriteString(plain("x"))
	}
	refusals := []struct{ name, math, reason string }{
		{"depth budget", deep, "depth or node budget"},
		{"node budget", wide.String(), "depth or node budget"},
		{"unknown element", `<m:eqArr><m:e>` + plain("x") + `</m:e></m:eqArr>`, "unsupported equation element m:eqArr"},
		{"matrix", `<m:m><m:mr><m:e>` + plain("x") + `</m:e></m:mr></m:m>`, "unsupported equation element m:m"},
		{"multi-character operator", `<m:nary><m:naryPr><m:chr m:val="∑∑"/></m:naryPr><m:sub/><m:sup/><m:e>` + plain("x") + `</m:e></m:nary>`, "exactly one character"},
		{"empty operator", `<m:acc><m:accPr><m:chr m:val=""/></m:accPr><m:e>` + plain("x") + `</m:e></m:acc>`, "empty equation operator character"},
		{"operator without value", `<m:nary><m:naryPr><m:chr/></m:naryPr><m:sub/><m:sup/><m:e>` + plain("x") + `</m:e></m:nary>`, "without a value"},
		{"unsupported fraction type", `<m:f><m:fPr><m:type m:val="skw"/></m:fPr><m:num>` + plain("a") + `</m:num><m:den>` + plain("b") + `</m:den></m:f>`, "unsupported fraction type skw"},
		{"unsupported script alphabet", `<m:r><m:rPr><m:scr m:val="fraktur"/></m:rPr>` + testMathRPr + `<m:t>x</m:t></m:r>`, "unsupported math script alphabet"},
		{"line break", `<m:r><m:rPr><m:brk/></m:rPr>` + testMathRPr + `<m:t>x</m:t></m:r>`, "unsupported math run property m:brk"},
		{"argument size", `<m:sSup><m:e>` + plain("x") + `</m:e><m:sup><m:argPr><m:argSz m:val="-1"/></m:argPr>` + plain("2") + `</m:sup></m:sSup>`, "argument property"},
		{"unknown construct property", `<m:d><m:dPr><m:unknown/></m:dPr><m:e>` + plain("x") + `</m:e></m:d>`, "unsupported equation property m:unknown"},
		{"foreign content", `<m:d><m:e>` + plain("x") + `<w:r><w:t>y</w:t></w:r></m:e></m:d>`, "foreign equation content r"},
		{"wrong argument order", `<m:f><m:den>` + plain("b") + `</m:den><m:num>` + plain("a") + `</m:num></m:f>`, "fraction requires numerator and denominator"},
		{"run without text", `<m:r>` + testMathRPr + `</m:r>`, "equation run without text"},
		{"unsupported run properties", `<m:r><w:rPr><w:shd w:val="clear" w:fill="FF0000"/></w:rPr><m:t>x</m:t></m:r>`, "unsupported run properties in equation"},
		{"hidden text", `<m:r><w:rPr><w:vanish/></w:rPr><m:t>x</m:t></m:r>`, "hidden equation text"},
	}
	for _, c := range refusals {
		t.Run(c.name, func(t *testing.T) {
			source := nativeApproximateEquationSource(t, nativeMathParagraph("", `<m:oMath>`+c.math+`</m:oMath>`))
			doc, out := inspectEquations(t, source)
			if out == nil || len(out.Items) != 1 {
				t.Fatalf("expected one omitted item: %+v", out)
			}
			// Equation content remains structural, read-only markup. Allowing
			// preserved body decoration must not widen math rendering/editing.
			if doc.Body.Blocks[0].Paragraph.EditPolicy.Mode != "read-only" {
				t.Fatal("omitted equation granted text editing authority")
			}
			item := out.Items[0]
			if item.Status != "omitted" || item.Lines != nil || !strings.Contains(item.Reason, c.reason) || len(item.DiagnosticIDs) == 0 || item.ParagraphID != doc.Body.Blocks[0].Paragraph.ID {
				t.Fatalf("unexpected omission: %#v", item)
			}
			if len(out.FontRequests) != 0 {
				t.Fatalf("omitted equations must not request fonts: %#v", out.FontRequests)
			}
		})
	}
	// Unsupported oMathPara justification and non-equation children refuse.
	for _, math := range []string{`<m:oMathParaPr><m:jc m:val="justify"/></m:oMathParaPr><m:oMath>` + plain("x") + `</m:oMath>`, `<m:oMath>` + plain("x") + `</m:oMath>` + plain("y")} {
		_, out := inspectEquations(t, nativeApproximateEquationSource(t, nativeMathParagraph("", math)))
		if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" {
			t.Fatalf("expected omission for %s: %+v", math, out)
		}
	}
	// No equations at all: nil sidecar.
	if _, out := inspectEquations(t, nativeApproximateEquationSource(t, `<w:p><w:r><w:t>plain</w:t></w:r></w:p>`)); out != nil {
		t.Fatalf("expected nil sidecar without equations: %+v", out)
	}
	if _, err := InspectNativeApproximateEquationsV1(nil); err == nil {
		t.Fatal("empty package accepted")
	}
}

func TestApproximateEquationsBudgetAndStrictNamespace(t *testing.T) {
	plain := func(text string) string { return `<m:r>` + testMathRPr + `<m:t>` + text + `</m:t></m:r>` }
	paragraphs := strings.Builder{}
	for i := 0; i < nativeApproximateEquationLimit+2; i++ {
		paragraphs.WriteString(nativeMathParagraph("", `<m:oMath>`+plain("x")+`</m:oMath>`))
	}
	_, out := inspectEquations(t, nativeApproximateEquationSource(t, paragraphs.String()))
	if out == nil || len(out.Items) != nativeApproximateEquationLimit || out.OmittedCount != 2 {
		t.Fatalf("equation budget: items=%d omitted=%d", len(out.Items), out.OmittedCount)
	}
	longFamily := strings.Repeat("F", nativeApproximateEquationFamilyLimit+1)
	_, out = inspectEquations(t, nativeApproximateEquationSource(t, nativeMathParagraph("", `<m:oMath><m:r><w:rPr><w:rFonts w:ascii="`+longFamily+`" w:hAnsi="`+longFamily+`"/></w:rPr><m:t>x</m:t></m:r></m:oMath>`)))
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || out.Items[0].Reason != "equation font family exceeds the preview bound" || len(out.FontRequests) != 0 {
		t.Fatalf("family bound: %+v", out)
	}
	longValue := strings.Repeat("é", 300)
	_, out = inspectEquations(t, nativeApproximateEquationSource(t, nativeMathParagraph("", `<m:oMath><m:r><m:rPr><m:scr m:val="`+longValue+`"/></m:rPr>`+testMathRPr+`<m:t>x</m:t></m:r></m:oMath>`)))
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || len(out.Items[0].Reason) > nativeApproximateEquationReasonLimit || !utf8.ValidString(out.Items[0].Reason) || !strings.HasPrefix(out.Items[0].Reason, "unsupported math script alphabet ") {
		t.Fatalf("reason bound: %q", out.Items[0].Reason)
	}
	long := strings.Repeat("x", nativeApproximateEquationTextLimit+1)
	_, out = inspectEquations(t, nativeApproximateEquationSource(t, nativeMathParagraph("", `<m:oMath>`+plain(long)+`</m:oMath>`)))
	if out == nil || len(out.Items) != 1 || out.Items[0].Status != "omitted" || !strings.Contains(out.Items[0].Reason, "text exceeds") {
		t.Fatalf("text budget: %+v", out)
	}
}
