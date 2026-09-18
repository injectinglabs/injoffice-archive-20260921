package docxpatch

import (
	"strings"
	"testing"
)

// A style that states before/after spacing, so the automatic flags below are
// resolved against an inherited measurement as well as the companion
// measurement they suppress.
const autoParagraphSpacingStyles = `<w:styles xmlns:w="` + wordMLTransitional + `">` +
	`<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:rPr><w:sz w:val="24"/></w:rPr></w:style>` +
	`<w:style w:type="paragraph" w:styleId="Spaced"><w:pPr><w:spacing w:before="240" w:after="360"/></w:pPr></w:style>` +
	`</w:styles>`

func autoParagraphSpacingParts(spacing string) map[string]string {
	parts := resolvedStylesTestParts(autoParagraphSpacingStyles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:p><w:pPr><w:pStyle w:val="Spaced"/>` + spacing + `</w:pPr><w:r><w:t>test</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
	return parts
}

func autoParagraphSpacingUnsupported(t *testing.T, spacing string) []string {
	t.Helper()
	document, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(autoParagraphSpacingParts(spacing))))
	if err != nil {
		t.Fatal(err)
	}
	codes := []string{}
	for _, entry := range document.Unsupported {
		codes = append(codes, entry.Code)
	}
	return codes
}

func autoParagraphSpacingResolved(t *testing.T, spacing string) (*NativeResolvedLayoutInputV1, []string) {
	t.Helper()
	resolved, err := ResolveNativeDocumentLayoutV1(buildNativeDOCX(t, nativeEntries(autoParagraphSpacingParts(spacing))))
	if err != nil {
		t.Fatal(err)
	}
	codes := []string{}
	for _, diagnostic := range resolved.Diagnostics {
		codes = append(codes, diagnostic.Code)
	}
	return resolved, codes
}

// w:beforeAutospacing/w:afterAutospacing is an exact source shape whose value
// Word determines; it is not unknown spacing structure. The extraction
// projection owns no spacing measurement anyway, so it must record the
// automatic fact under the same code the resolved-layout projection uses
// instead of reporting the element as unmodeled and stopping every downstream
// reader on a shape it can read.
func TestNativeAutomaticParagraphSpacingIsRecordedAsAutomatic(t *testing.T) {
	codes := autoParagraphSpacingUnsupported(t, `<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>`)
	joined := strings.Join(codes, ",")
	if strings.Contains(joined, "UNMODELED_PARAGRAPH_SPACING") {
		t.Fatalf("automatic paragraph spacing is still reported as unmodeled structure: %v", codes)
	}
	if !strings.Contains(joined, "AUTO_PARAGRAPH_SPACING_PRESERVED") {
		t.Fatalf("automatic paragraph spacing lost its diagnostic: %v", codes)
	}
}

// Word ignores only the companion measurement of an automatic flag; the value
// it determines is what the two Word references in the benchmark corpus
// attest, and both agree with the measurement the paragraph would have
// inherited anyway. para-auto-spacing.docx paints its automatic paragraph
// flush with the top margin and adds nothing between it and the following
// paragraph's own 400-twip before spacing, over a Normal style that states no
// spacing. dml-groupshape-paraspacing.docx puts an automatic paragraph under
// docDefaults that state after="200": its before spacing is again nothing -
// the gap to the preceding paragraph is one line plus that paragraph's own
// 18-point after spacing - while the text block sits in its shape as if the
// inherited after spacing still trailed it. So the flag suppresses its own
// companion measurement and nothing else; guessing a value for it, in either
// direction, is not attested. This pins that resolution, which the
// AUTO_PARAGRAPH_SPACING_PRESERVED diagnostic keeps out of the exact tier.
func TestNativeAutomaticParagraphSpacingIgnoresOnlyItsCompanionMeasurement(t *testing.T) {
	resolved, codes := autoParagraphSpacingResolved(t, `<w:spacing w:before="100" w:beforeAutospacing="1" w:after="100" w:afterAutospacing="1"/>`)
	if !strings.Contains(strings.Join(codes, ","), "AUTO_PARAGRAPH_SPACING_PRESERVED") {
		t.Fatalf("automatic paragraph spacing lost its resolution diagnostic: %v", codes)
	}
	if len(resolved.Paragraphs) != 1 {
		t.Fatalf("paragraphs = %#v", resolved.Paragraphs)
	}
	properties := resolved.Paragraphs[0].Properties
	if properties.SpacingBeforeTwips == nil || *properties.SpacingBeforeTwips != 240 {
		t.Fatalf("automatic before spacing did not inherit the style measurement: %#v", properties.SpacingBeforeTwips)
	}
	if properties.SpacingAfterTwips == nil || *properties.SpacingAfterTwips != 360 {
		t.Fatalf("automatic after spacing did not inherit the style measurement: %#v", properties.SpacingAfterTwips)
	}
	// An explicitly disabled flag keeps stating nothing, so its companion
	// measurement stays the paragraph's own direct formatting.
	disabled, _ := autoParagraphSpacingResolved(t, `<w:spacing w:before="100" w:beforeAutospacing="0"/>`)
	if before := disabled.Paragraphs[0].Properties.SpacingBeforeTwips; before == nil || *before != 100 {
		t.Fatalf("disabled automatic spacing discarded its own measurement: %#v", before)
	}
}

// Everything else about w:spacing that the exact resolved-layout subset cannot
// read keeps refusing. Admitting the automatic flag must not admit a negative
// or non-numeric measurement, a line-unit measurement, an unreadable flag, or a
// line rule without a measurement.
func TestNativeUnreadableParagraphSpacingStillRefuses(t *testing.T) {
	for _, spacing := range []string{
		// `<w:spacing w:line="-240"/>` used to live here. A negative w:line is a
		// measurement Word reads, not unknown structure, so it now carries
		// NEGATIVE_LINE_SPACING_UNAPPLIED and the approximate tier paints the
		// paragraph and discloses the unapplied compression - see
		// native_negative_line_spacing_test.go, which also pins that a negative
		// w:before or w:after, and a negative w:line beside an invalid rule,
		// stay unmodeled and keep refusing.
		`<w:spacing w:before="-240" w:beforeAutospacing="1"/>`,
		`<w:spacing w:after="-240" w:line="-240"/>`,
		`<w:spacing w:before="auto" w:beforeAutospacing="1"/>`,
		`<w:spacing w:beforeLines="150" w:afterLines="175"/>`,
		`<w:spacing w:beforeAutospacing="maybe"/>`,
		`<w:spacing w:lineRule="exact"/>`,
		`<w:spacing w:line="240" w:lineRule="atMost"/>`,
	} {
		codes := autoParagraphSpacingUnsupported(t, spacing)
		if !strings.Contains(strings.Join(codes, ","), "UNMODELED_PARAGRAPH_SPACING") {
			t.Fatalf("%s no longer refuses: %v", spacing, codes)
		}
	}
	// A spacing element the subset does read keeps carrying no spacing
	// diagnostic at all.
	codes := autoParagraphSpacingUnsupported(t, `<w:spacing w:before="240" w:after="240" w:line="360" w:lineRule="auto"/>`)
	joined := strings.Join(codes, ",")
	if strings.Contains(joined, "UNMODELED_PARAGRAPH_SPACING") || strings.Contains(joined, "AUTO_PARAGRAPH_SPACING_PRESERVED") {
		t.Fatalf("exact paragraph spacing gained a diagnostic: %v", codes)
	}
}
