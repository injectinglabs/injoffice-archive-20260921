package pptxpatch

import (
	"strings"
	"testing"
)

// LibreOffice writes layouts whose placeholders repeat a cNvPr id and masters
// whose first placeholder is id 0, and both parts are read-only inheritance
// sources whose ids never anchor a mutation.
const nativeInheritedLayoutShapes = `<p:sp><p:nvSpPr><p:cNvPr id="5" name="PlaceHolder 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>` +
	`<p:sp><p:nvSpPr><p:cNvPr id="5" name="PlaceHolder 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`

const nativeInheritedMasterShapes = `<p:sp><p:nvSpPr><p:cNvPr id="0" name="PlaceHolder 1"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>` +
	`<p:sp><p:nvSpPr><p:cNvPr id="1" name="PlaceHolder 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`

func TestExtractNativePPTXReadsLayoutAndMasterShapeTreesWithRepeatedOrZeroIDs(t *testing.T) {
	t.Parallel()

	objectIDs := map[string][]string{}
	options := nativeTestExtractOptions()
	inner := options.TokenFactory
	options.TokenFactory = NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		objectIDs[request.OwnerPart] = append(objectIDs[request.OwnerPart], request.ObjectID)
		return inner.IssueNativePassthroughToken(request)
	})

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], `</p:spTree>`, nativeInheritedLayoutShapes+`</p:spTree>`, 1)
		// The master's second placeholder also repeats the root group's id.
		parts["relocated/masters/master.xml"] = strings.Replace(parts["relocated/masters/master.xml"], `</p:spTree>`, nativeInheritedMasterShapes+`</p:spTree>`, 1)
	}})

	deck, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("layout/master shape trees with repeated or zero cNvPr ids were refused: %v", err)
	}
	if len(deck.Slides) != 1 {
		t.Fatalf("expected the deck's one slide, got %d", len(deck.Slides))
	}
	for _, part := range []string{"relocated/layouts/layout.xml", "relocated/masters/master.xml"} {
		seen := map[string]bool{}
		for _, objectID := range objectIDs[part] {
			if seen[objectID] {
				t.Fatalf("%s reused the preserved object label %q", part, objectID)
			}
			seen[objectID] = true
		}
		if len(seen) < 2 {
			t.Fatalf("%s preserved %d inherited shapes, expected both", part, len(seen))
		}
	}
}

func TestExtractNativePPTXStillRefusesRepeatedOrZeroCNvPrIDsOnASlide(t *testing.T) {
	t.Parallel()

	// A slide's ids are identity: every mutation anchors on one, so the same
	// markup a layout tolerates must still fail the package here.
	for _, testCase := range []struct{ name, shapes, want string }{
		{"duplicate", `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Repeat"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`, "duplicate cNvPr id cNvPr-2"},
		{"rootDuplicate", `<p:sp><p:nvSpPr><p:cNvPr id="1" name="Root"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`, "duplicate cNvPr id cNvPr-1"},
		{"zero", `<p:sp><p:nvSpPr><p:cNvPr id="0" name="Zero"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>`, "invalid shape-tree cNvPr id"},
	} {
		payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, testCase.shapes+`</p:spTree>`, 1)
		}})
		_, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
		if err == nil || !strings.Contains(err.Error(), testCase.want) {
			t.Fatalf("slide case %q: expected %q, got %v", testCase.name, testCase.want, err)
		}
	}
}
