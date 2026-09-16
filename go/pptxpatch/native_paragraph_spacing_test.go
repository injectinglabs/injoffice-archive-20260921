package pptxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

// nativeParagraphSpacingFixture builds the inherited-text preview fixture with
// authored spacing injected into the requested cascade layers. Every layer the
// caller leaves empty stays absent, so one helper exercises each of them.
func nativeParagraphSpacingFixture(t *testing.T, strict bool, presentation, master, list, local string) []byte {
	t.Helper()
	ns := nsDrawingTransitional
	if strict {
		ns = nsDrawingStrict
	}
	return nativeShapeReferenceFixture(t, strict, nativeShapeStyleRefs, `lang="en-US" b="0" i="0" sz="1800">`, list, "Hi", nativeShapeReferenceFonts, func(parts map[string]string) {
		parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:presentation>`,
			`<p:defaultTextStyle xmlns:a="`+ns+`"><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr>`+presentation+`</a:lvl1pPr></p:defaultTextStyle></p:presentation>`, 1)
		parts["relocated/masters/master.xml"] = strings.Replace(parts["relocated/masters/master.xml"], `</p:sldMaster>`,
			`<p:txStyles xmlns:a="`+ns+`"><p:otherStyle><a:lvl1pPr>`+master+`</a:lvl1pPr></p:otherStyle></p:txStyles></p:sldMaster>`, 1)
		if local != "" {
			part := "relocated/slides/slide-a.xml"
			parts[part] = strings.Replace(parts[part], `<a:pPr algn="l" lvl="0"><a:buNone/></a:pPr>`,
				`<a:pPr algn="l" lvl="0">`+local+`<a:buNone/></a:pPr>`, 1)
		}
	})
}

func nativeParagraphSpacingElement(t *testing.T, data []byte, preview bool) NativeElement {
	t.Helper()
	options := nativeMutationExtractOptions()
	options.AllowInheritedTextPreview = preview
	deck, err := ExtractNativePPTX(data, options)
	if err != nil {
		t.Fatal(err)
	}
	if issues := ValidateNativePPTX(deck); len(issues) > 0 {
		t.Fatalf("invalid contract: %+v", issues)
	}
	return nativeFixtureAutoShapes(deck.Slides[0])[0]
}

func nativeSpacingOf(t *testing.T, element NativeElement) NativeParagraph {
	t.Helper()
	if element.Paragraphs == nil || len(*element.Paragraphs) != 1 {
		t.Fatalf("expected one projected paragraph: %+v", element.Paragraphs)
	}
	return (*element.Paragraphs)[0]
}

// The preview walks direct paragraph properties over the body list style over
// the master text style over the presentation defaults. Each layer must be
// able to supply a value and each later layer must be able to replace it.
func TestNativeParagraphSpacingResolvesThroughEachCascadeLayer(t *testing.T) {
	const (
		lnSpc80  = `<a:lnSpc><a:spcPct val="80000"/></a:lnSpc>`
		lnSpc90  = `<a:lnSpc><a:spcPct val="90000"/></a:lnSpc>`
		before4  = `<a:spcBef><a:spcPts val="400"/></a:spcBef>`
		before8  = `<a:spcBef><a:spcPts val="800"/></a:spcBef>`
		after12  = `<a:spcAft><a:spcPts val="1200"/></a:spcAft>`
		after16  = `<a:spcAft><a:spcPts val="1600"/></a:spcAft>`
		emuPerPt = int64(12700)
	)
	for _, tc := range []struct {
		name                                   string
		presentation, master, list, local      string
		wantLinePercent, wantBefore, wantAfter int64
	}{
		{"presentation defaults", lnSpc80 + before4 + after12, "", "", "", 80000, 4 * emuPerPt, 12 * emuPerPt},
		{"master text style wins", lnSpc80 + before4, lnSpc90 + before8 + after12, "", "", 90000, 8 * emuPerPt, 12 * emuPerPt},
		{"body list style wins", lnSpc80 + before4, lnSpc90 + before8, lnSpc80 + after16, "", 80000, 8 * emuPerPt, 16 * emuPerPt},
		{"direct paragraph properties win", lnSpc80 + before4, lnSpc90 + before8, lnSpc80 + after16, lnSpc90 + before4 + after12, 90000, 4 * emuPerPt, 12 * emuPerPt},
	} {
		t.Run(tc.name, func(t *testing.T) {
			for _, strict := range []bool{false, true} {
				list := tc.list
				if list != "" {
					list = `<a:lvl1pPr>` + list + `</a:lvl1pPr>`
				}
				data := nativeParagraphSpacingFixture(t, strict, tc.presentation, tc.master, list, tc.local)
				element := nativeParagraphSpacingElement(t, data, true)
				paragraph := nativeSpacingOf(t, element)
				if paragraph.LineSpacingPercent1000 == nil || *paragraph.LineSpacingPercent1000 != tc.wantLinePercent {
					t.Fatalf("line spacing came from the wrong layer: %+v", paragraph.LineSpacingPercent1000)
				}
				if paragraph.LineSpacingEmu != nil {
					t.Fatalf("a percentage line spacing also emitted an absolute pitch: %+v", paragraph.LineSpacingEmu)
				}
				if paragraph.SpaceBeforeEmu == nil || *paragraph.SpaceBeforeEmu != tc.wantBefore {
					t.Fatalf("space before came from the wrong layer: %+v", paragraph.SpaceBeforeEmu)
				}
				if paragraph.SpaceAfterEmu == nil || *paragraph.SpaceAfterEmu != tc.wantAfter {
					t.Fatalf("space after came from the wrong layer: %+v", paragraph.SpaceAfterEmu)
				}
				if codes := nativeDiagnosticCodes(element); codes[nativeParagraphSpacingCode] != 1 {
					t.Fatalf("spacing was carried without exactly one disclosure: %+v", element.Compatibility.Diagnostics)
				}
				for _, diagnostic := range element.Compatibility.Diagnostics {
					if diagnostic.Code == nativeInheritedTextOmissionsCode && (strings.Contains(diagnostic.Message, "a:lnSpc") || strings.Contains(diagnostic.Message, "a:spcBef") || strings.Contains(diagnostic.Message, "a:spcAft")) {
						t.Fatalf("carried spacing is still listed as an omission: %s", diagnostic.Message)
					}
				}
			}
		})
	}
}

// ECMA-376 21.1.2.2.7/21.1.2.2.9: a:spcPts is absolute hundredths of a point
// and a:spcPct is a percentage of the paragraph's largest authored text size.
func TestNativeParagraphSpacingResolvesPointsAndPercentages(t *testing.T) {
	// The fixture run is sz="1800", so 100% of the text size is 1800*127 EMU.
	const textSizeEMU = int64(1800 * 127)
	for _, tc := range []struct {
		name, markup string
		wantBefore   int64
	}{
		{"absolute points", `<a:spcBef><a:spcPts val="1000"/></a:spcBef>`, 1000 * 127},
		{"quarter of the text size", `<a:spcBef><a:spcPct val="25000"/></a:spcBef>`, textSizeEMU / 4},
		{"full text size", `<a:spcBef><a:spcPct val="100000"/></a:spcBef>`, textSizeEMU},
	} {
		t.Run(tc.name, func(t *testing.T) {
			element := nativeParagraphSpacingElement(t, nativeParagraphSpacingFixture(t, false, tc.markup, "", "", ""), true)
			paragraph := nativeSpacingOf(t, element)
			if paragraph.SpaceBeforeEmu == nil || *paragraph.SpaceBeforeEmu != tc.wantBefore {
				t.Fatalf("want %d EMU, got %+v", tc.wantBefore, paragraph.SpaceBeforeEmu)
			}
		})
	}
	// An absolute a:lnSpc converts to an absolute pitch, never a percentage.
	element := nativeParagraphSpacingElement(t, nativeParagraphSpacingFixture(t, false, `<a:lnSpc><a:spcPts val="1500"/></a:lnSpc>`, "", "", ""), true)
	paragraph := nativeSpacingOf(t, element)
	if paragraph.LineSpacingEmu == nil || *paragraph.LineSpacingEmu != 1500*127 || paragraph.LineSpacingPercent1000 != nil {
		t.Fatalf("absolute line spacing was not carried exactly: %+v %+v", paragraph.LineSpacingEmu, paragraph.LineSpacingPercent1000)
	}
}

// Absence already means "no authored adjustment", so identity values stay out
// of the contract and never raise a disclosure of their own.
func TestNativeParagraphSpacingOmitsIdentityValues(t *testing.T) {
	for _, markup := range []string{
		``,
		`<a:lnSpc><a:spcPct val="100000"/></a:lnSpc>`,
		`<a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPct val="0"/></a:spcAft>`,
		`<a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef>`,
	} {
		element := nativeParagraphSpacingElement(t, nativeParagraphSpacingFixture(t, false, markup, "", "", ""), true)
		if nativeParagraphsCarrySpacing(*element.Paragraphs) {
			t.Fatalf("%s emitted a spacing projection: %+v", markup, *element.Paragraphs)
		}
		if codes := nativeDiagnosticCodes(element); codes[nativeParagraphSpacingCode] != 0 {
			t.Fatalf("%s disclosed spacing it never carried: %+v", markup, element.Compatibility.Diagnostics)
		}
	}
}

// The exact tier never walks the cascade, so authored spacing in the
// presentation/master/list layers must leave its output byte for byte alone.
// The two fixtures differ only by that markup, so their exact-tier paragraphs,
// text-body layout and diagnostic codes must be indistinguishable.
func TestNativeParagraphSpacingLeavesTheExactTierByteIdentical(t *testing.T) {
	for _, strict := range []bool{false, true} {
		spaced := nativeParagraphSpacingFixture(t, strict,
			`<a:lnSpc><a:spcPct val="80000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef>`,
			`<a:spcAft><a:spcPts val="600"/></a:spcAft>`,
			`<a:lvl1pPr><a:lnSpc><a:spcPct val="90000"/></a:lnSpc></a:lvl1pPr>`, "")
		plain := nativeParagraphSpacingFixture(t, strict, "", "", "", "")
		spacedExact := nativeExactTierTextShape(t, spaced)
		plainExact := nativeExactTierTextShape(t, plain)
		if !bytes.Equal(spacedExact, plainExact) {
			t.Fatalf("authored cascade spacing changed the exact tier:\n%s\n%s", spacedExact, plainExact)
		}
		if strings.Contains(string(spacedExact), "lineSpacing") || strings.Contains(string(spacedExact), "spaceBefore") || strings.Contains(string(spacedExact), "spaceAfter") {
			t.Fatalf("the exact tier emitted an approximate spacing field: %s", spacedExact)
		}
		// The approximate tier does carry it, from the same source bytes.
		previewParagraph := nativeSpacingOf(t, nativeParagraphSpacingElement(t, spaced, true))
		if previewParagraph.LineSpacingPercent1000 == nil || *previewParagraph.LineSpacingPercent1000 != 90000 || previewParagraph.SpaceBeforeEmu == nil || previewParagraph.SpaceAfterEmu == nil {
			t.Fatalf("the approximate tier did not carry the same authored spacing: %+v", previewParagraph)
		}
	}
}

// nativeExactTierTextShape renders the exact-tier projection of the fixture's
// shape: its paragraphs, text-body layout, status and diagnostic codes. Deck
// identity is intentionally excluded because the two fixtures differ in bytes.
func nativeExactTierTextShape(t *testing.T, data []byte) []byte {
	t.Helper()
	element := nativeFixtureAutoShapes(mustExtractNativeSpacingDeck(t, data, false).Slides[0])[0]
	codes := []string{}
	for _, diagnostic := range element.Compatibility.Diagnostics {
		codes = append(codes, string(diagnostic.Severity)+" "+diagnostic.Code)
	}
	encoded, err := json.Marshal(struct {
		Paragraphs  *[]NativeParagraph
		TextBody    *NativeTextBodyLayout
		Status      NativeCompatibilityStatus
		Diagnostics []string
	}{element.Paragraphs, element.TextBody, element.Compatibility.Status, codes})
	if err != nil {
		t.Fatal(err)
	}
	return encoded
}

func mustExtractNativeSpacingDeck(t *testing.T, data []byte, preview bool) NativePPTXDeck {
	t.Helper()
	options := nativeMutationExtractOptions()
	options.AllowInheritedTextPreview = preview
	deck, err := ExtractNativePPTX(data, options)
	if err != nil {
		t.Fatal(err)
	}
	return deck
}

func TestNativeParagraphSpacingContractRules(t *testing.T) {
	extract := func(t *testing.T) NativePPTXDeck {
		t.Helper()
		return mustExtractNativeSpacingDeck(t, nativeParagraphSpacingFixture(t, false,
			`<a:lnSpc><a:spcPct val="80000"/></a:lnSpc><a:spcBef><a:spcPts val="1000"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft>`, "", "", ""), true)
	}
	target := func(deck NativePPTXDeck) *NativeParagraph {
		for index := range deck.Slides[0].Elements {
			element := &deck.Slides[0].Elements[index]
			if element.Paragraphs != nil && nativeParagraphsCarrySpacing(*element.Paragraphs) {
				return &(*element.Paragraphs)[0]
			}
		}
		t.Fatal("fixture did not carry paragraph spacing")
		return nil
	}
	element := func(deck NativePPTXDeck) *NativeElement {
		for index := range deck.Slides[0].Elements {
			if deck.Slides[0].Elements[index].Paragraphs != nil && nativeParagraphsCarrySpacing(*deck.Slides[0].Elements[index].Paragraphs) {
				return &deck.Slides[0].Elements[index]
			}
		}
		t.Fatal("fixture did not carry paragraph spacing")
		return nil
	}
	if issues := ValidateNativePPTX(extract(t)); len(issues) > 0 {
		t.Fatalf("the authored projection is invalid: %+v", issues)
	}
	for _, value := range []int64{1, 90000, 13200000} {
		deck := extract(t)
		target(deck).LineSpacingPercent1000 = int64Pointer(value)
		if issues := ValidateNativePPTX(deck); len(issues) > 0 {
			t.Fatalf("line spacing %d was rejected: %+v", value, issues)
		}
	}
	for _, value := range []int64{0, -1, 13200001} {
		deck := extract(t)
		target(deck).LineSpacingPercent1000 = int64Pointer(value)
		if len(ValidateNativePPTX(deck)) == 0 {
			t.Fatalf("line spacing %d validated", value)
		}
	}
	for _, field := range []string{"lineSpacingEmu", "spaceBeforeEmu", "spaceAfterEmu"} {
		for _, value := range []int64{0, -1, 51206401} {
			deck := extract(t)
			paragraph := target(deck)
			paragraph.LineSpacingPercent1000 = nil
			switch field {
			case "lineSpacingEmu":
				paragraph.LineSpacingEmu = int64Pointer(value)
			case "spaceBeforeEmu":
				paragraph.SpaceBeforeEmu = int64Pointer(value)
			case "spaceAfterEmu":
				paragraph.SpaceAfterEmu = int64Pointer(value)
			}
			if len(ValidateNativePPTX(deck)) == 0 {
				t.Fatalf("%s %d validated", field, value)
			}
		}
	}
	// A paragraph carries one authored line spacing, never both forms.
	deck := extract(t)
	target(deck).LineSpacingEmu = int64Pointer(200000)
	if len(ValidateNativePPTX(deck)) == 0 {
		t.Fatal("a paragraph validated with both line-spacing forms")
	}
	// The projection may not outlive its disclosure.
	deck = extract(t)
	owner := element(deck)
	owner.Compatibility.Diagnostics = nil
	owner.Compatibility.Status = NativeCompatibilityStatusEditable
	if len(ValidateNativePPTX(deck)) == 0 {
		t.Fatal("an editable element kept the authored paragraph spacing")
	}
	deck = extract(t)
	owner = element(deck)
	for index := range owner.Compatibility.Diagnostics {
		if owner.Compatibility.Diagnostics[index].Code == nativeParagraphSpacingCode {
			owner.Compatibility.Diagnostics[index].Code = nativeInheritedTextPreviewCode
		}
	}
	if len(ValidateNativePPTX(deck)) == 0 {
		t.Fatal("the inherited-text disclosure authorized a paragraph-spacing projection")
	}
	// A preview-only approximation never grants editing authority.
	deck = extract(t)
	if element(deck).Compatibility.Status == NativeCompatibilityStatusEditable {
		t.Fatal("paragraph spacing left the element editable")
	}
}

func TestNativeParagraphGapEMURefusesWhatItCannotModel(t *testing.T) {
	// A percentage with no authored text size has no basis to resolve against.
	if _, ok := nativeParagraphGapEMU(int64Pointer(25000), nil, 0); ok {
		t.Fatal("a percentage gap resolved without an authored text size")
	}
	// A zero percentage is exactly zero whatever the basis would have been.
	if gap, ok := nativeParagraphGapEMU(int64Pointer(0), nil, 0); !ok || gap != 0 {
		t.Fatalf("a zero percentage gap was not modeled: %d %v", gap, ok)
	}
	if _, ok := nativeParagraphGapEMU(nil, int64Pointer(51206401/127+1), 0); ok {
		t.Fatal("an out-of-range absolute gap resolved")
	}
	if gap, ok := nativeParagraphGapEMU(nil, int64Pointer(1000), 1800); !ok || gap != 127000 {
		t.Fatalf("absolute gap conversion is not exact: %d %v", gap, ok)
	}
}
