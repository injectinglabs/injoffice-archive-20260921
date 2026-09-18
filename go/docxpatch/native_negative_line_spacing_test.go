package docxpatch

import (
	"strings"
	"testing"
)

// ECMA-376 17.3.1.33 types w:line as ST_SignedTwipsMeasure. A negative value is
// a measurement Word reads - it applies the absolute value as an exact line
// height and compresses the lines - not malformed markup. Recording it as
// unknown spacing structure refused the whole document; it now carries its own
// code so the approximate tier can paint the paragraph at the line spacing it
// inherits and disclose that the compression was not applied.
//
// Every other defect in the same element stays unmodeled and keeps refusing on
// both tiers: a negative w:before or w:after moves the paragraph box itself, and
// an unparseable measurement states nothing this reading can name.
func TestNativeNegativeLineSpacingIsItsOwnCode(t *testing.T) {
	for _, tc := range []struct {
		name    string
		spacing string
		code    string
	}{
		{"negative line", `<w:spacing w:line="-240"/>`, "NEGATIVE_LINE_SPACING_UNAPPLIED"},
		{"negative line with an exact rule", `<w:spacing w:line="-240" w:lineRule="exact"/>`, "NEGATIVE_LINE_SPACING_UNAPPLIED"},
		{"negative line with an at-least rule", `<w:spacing w:line="-240" w:lineRule="atLeast"/>`, "NEGATIVE_LINE_SPACING_UNAPPLIED"},
		{"negative line beside valid before/after", `<w:spacing w:before="120" w:after="240" w:line="-360"/>`, "NEGATIVE_LINE_SPACING_UNAPPLIED"},

		{"negative before", `<w:spacing w:before="-120" w:line="240"/>`, "UNMODELED_PARAGRAPH_SPACING"},
		{"negative after", `<w:spacing w:after="-120" w:line="240"/>`, "UNMODELED_PARAGRAPH_SPACING"},
		{"unparseable line", `<w:spacing w:line="tight"/>`, "UNMODELED_PARAGRAPH_SPACING"},
		{"empty line", `<w:spacing w:line=""/>`, "UNMODELED_PARAGRAPH_SPACING"},
		{"negative line with an invalid rule", `<w:spacing w:line="-240" w:lineRule="squeeze"/>`, "UNMODELED_PARAGRAPH_SPACING"},
		{"line-unit spacing", `<w:spacing w:beforeLines="100" w:line="-240"/>`, "UNMODELED_PARAGRAPH_SPACING"},
		{"unknown attribute", `<w:spacing w:line="-240" w:other="1"/>`, "UNMODELED_PARAGRAPH_SPACING"},
	} {
		t.Run(tc.name, func(t *testing.T) {
			codes := strings.Join(autoParagraphSpacingUnsupported(t, tc.spacing), ",")
			if !strings.Contains(codes, tc.code) {
				t.Fatalf("extractor recorded %v, want %s", codes, tc.code)
			}
			other := "UNMODELED_PARAGRAPH_SPACING"
			if tc.code == other {
				other = "NEGATIVE_LINE_SPACING_UNAPPLIED"
			}
			if strings.Contains(codes, other) {
				t.Fatalf("%s must replace %s, not accompany it: %v", tc.code, other, codes)
			}
		})
	}
}

// The resolver reports the same distinction, because the approximate tier
// consults its diagnostics as well as the extractor's unsupported records. A
// negative measurement is dropped either way - the paragraph keeps the spacing
// it inherits - so what changes is only whether the page can be painted and the
// drop disclosed, or the whole document refused.
func TestNativeNegativeLineSpacingResolverDiagnostic(t *testing.T) {
	for _, tc := range []struct {
		spacing string
		code    string
	}{
		{`<w:spacing w:line="-240"/>`, "NEGATIVE_LINE_SPACING_UNAPPLIED"},
		{`<w:spacing w:line="-240" w:lineRule="exact"/>`, "NEGATIVE_LINE_SPACING_UNAPPLIED"},
		{`<w:spacing w:line="tight"/>`, "INVALID_LINE_SPACING"},
		{`<w:spacing w:line="-240" w:lineRule="squeeze"/>`, "INVALID_LINE_SPACING"},
	} {
		t.Run(tc.spacing, func(t *testing.T) {
			resolved, codes := autoParagraphSpacingResolved(t, tc.spacing)
			joined := strings.Join(codes, ",")
			if !strings.Contains(joined, tc.code) {
				t.Fatalf("resolver recorded %v, want %s", codes, tc.code)
			}
			if tc.code == "NEGATIVE_LINE_SPACING_UNAPPLIED" {
				if strings.Contains(joined, "INVALID_LINE_SPACING") {
					t.Fatalf("a negative measurement must not also be reported as invalid: %v", codes)
				}
				// A rule beside a negative measurement is not a rule without a
				// measurement: the source states one and this tier dropped it.
				if strings.Contains(joined, "INCOMPLETE_LINE_SPACING") {
					t.Fatalf("the authored w:lineRule must not be reported as orphaned: %v", codes)
				}
			}
			// Whatever the code, the measurement is never applied: the paragraph
			// keeps the line spacing it inherits.
			for _, paragraph := range resolved.Paragraphs {
				if paragraph.Properties.Line != nil {
					t.Fatalf("a negative line measurement must not reach the resolved layout: %#v", paragraph.Properties)
				}
			}
		})
	}
}

// The disclosure must say the compression was not applied and which way the
// painted page is wrong, or the page is silently further apart than Word's.
func TestNativeNegativeLineSpacingDisclosesTheDeviation(t *testing.T) {
	document, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(autoParagraphSpacingParts(`<w:spacing w:line="-240"/>`))))
	if err != nil {
		t.Fatal(err)
	}
	message := ""
	for _, entry := range document.Unsupported {
		if entry.Code == "NEGATIVE_LINE_SPACING_UNAPPLIED" {
			message = entry.Message
		}
	}
	for _, want := range []string{"NOT applied", "exact line height of its absolute value", "FURTHER APART"} {
		if !strings.Contains(message, want) {
			t.Fatalf("disclosure does not say %q: %s", want, message)
		}
	}
}
