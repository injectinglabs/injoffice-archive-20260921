package pptxpatch

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"strings"
	"testing"
)

func TestExtractNativePPTXSecureOPCMinimalText(t *testing.T) {
	t.Parallel()

	for _, strict := range []bool{false, true} {
		strict := strict
		name := "transitional"
		if strict {
			name = "strict"
		}
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict})
			deck, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("ExtractNativePPTX: %v", err)
			}
			if deck.ContractVersion != NativePPTXContractVersion || deck.Origin != NativeOriginParsed {
				t.Fatalf("unexpected contract identity: %#v", deck)
			}
			if deck.Size.Cx == nil || *deck.Size.Cx != 12_192_000 || deck.Size.Cy == nil || *deck.Size.Cy != 6_858_000 {
				t.Fatalf("unexpected slide size: %#v", deck.Size)
			}
			if len(deck.Slides) != 1 || deck.Slides[0].Source == nil || deck.Slides[0].Source.PartName != "relocated/slides/slide-a.xml" {
				t.Fatalf("unexpected relocated slide: %#v", deck.Slides)
			}
			elements := deck.Slides[0].Elements
			if len(elements) != 1 || elements[0].Kind != NativeElementKindText || elements[0].Source == nil || elements[0].Source.ObjectID != "cNvPr-2" {
				t.Fatalf("unexpected text element: %#v", elements)
			}
			if elements[0].Transform.X == nil || *elements[0].Transform.X != 914_400 || elements[0].Transform.Cx == nil || *elements[0].Transform.Cx != 4_572_000 {
				t.Fatalf("unexpected text transform: %#v", elements[0].Transform)
			}
			if elements[0].Paragraphs == nil || len(*elements[0].Paragraphs) != 1 || len((*elements[0].Paragraphs)[0].Runs) != 2 {
				t.Fatalf("unexpected text paragraphs: %#v", elements[0].Paragraphs)
			}
			if elements[0].Compatibility.Status != NativeCompatibilityStatusEditable || len(elements[0].Passthrough) != 0 || elements[0].TextBody == nil {
				t.Fatalf("exact v1 text body was not editable: %#v", elements[0])
			}
			paragraph := (*elements[0].Paragraphs)[0]
			if paragraph.Align == nil || *paragraph.Align != NativeTextAlignCenter || paragraph.Bullet == nil || *paragraph.Bullet {
				t.Fatalf("unexpected paragraph formatting: %#v", paragraph)
			}
			first := paragraph.Runs[0]
			if first.Text == nil || *first.Text != "Hello " || first.Bold == nil || !*first.Bold || first.FontFamily == nil || *first.FontFamily != "Aptos" || first.Color == nil || *first.Color != "112233" {
				t.Fatalf("unexpected rich run: %#v", first)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("extracted deck is invalid: %#v", issues)
			}
			if _, err := MarshalNativePPTXJSON(deck); err != nil {
				t.Fatalf("canonical JSON: %v", err)
			}
		})
	}
}

func TestExtractNativePPTXAcceptsCanonicalOfficeXMLDeclaration(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/deck.xml"] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` + parts["relocated/deck.xml"]
	}})
	if _, err := ExtractNativePPTX(payload, nativeTestExtractOptions()); err != nil {
		t.Fatalf("canonical Office XML declaration was rejected: %v", err)
	}
}

func TestExtractNativePPTXStrictUsesOPCRelationshipNamespace(t *testing.T) {
	t.Parallel()

	if _, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{strict: true}), nativeTestExtractOptions()); err != nil {
		t.Fatalf("Strict Office relationship types under the OPC relationship namespace were rejected: %v", err)
	}
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: true, mutate: func(parts map[string]string) {
		parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], nsPackageRels, "http://purl.oclc.org/ooxml/package/relationships", 1)
	}})
	if _, err := ExtractNativePPTX(payload, nativeTestExtractOptions()); err == nil {
		t.Fatal("nonstandard purl package relationship namespace was accepted")
	}
}

func TestExtractNativePPTXReusesPreviousNativeIdentity(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	previous.DocumentID = "deck-reused"
	oldSlideID := previous.Slides[0].ID
	oldElementID := previous.Slides[0].Elements[0].ID
	previous.Slides[0].ID = "slide-reused"
	previous.Slides[0].Elements[0].ID = "element-reused"
	for _, compatibility := range []*NativeCompatibility{&previous.Compatibility, &previous.Slides[0].Compatibility, &previous.Slides[0].Elements[0].Compatibility} {
		for index := range compatibility.Diagnostics {
			scope := compatibility.Diagnostics[index].Scope
			if scope == nil {
				continue
			}
			if scope.SlideID != nil && *scope.SlideID == oldSlideID {
				*scope.SlideID = "slide-reused"
			}
			if scope.ElementID != nil && *scope.ElementID == oldElementID {
				*scope.ElementID = "element-reused"
			}
		}
	}

	options := nativeTestExtractOptions()
	options.Previous = &previous
	deck, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("extract with previous: %v", err)
	}
	if deck.DocumentID != "deck-reused" || deck.Slides[0].ID != "slide-reused" || deck.Slides[0].Elements[0].ID != "element-reused" {
		t.Fatalf("identities were not reused: document=%q slide=%q element=%q", deck.DocumentID, deck.Slides[0].ID, deck.Slides[0].Elements[0].ID)
	}
}

func TestExtractNativePPTXReusesSlideAndElementIdentityAcrossReorder(t *testing.T) {
	t.Parallel()

	initial, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{secondSlide: true}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	initialIDs := map[string][2]string{}
	for _, slide := range initial.Slides {
		initialIDs[slide.Source.PartName] = [2]string{slide.ID, slide.Elements[0].ID}
	}
	options := nativeTestExtractOptions()
	options.Previous = &initial
	reordered, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{secondSlide: true, reverseSlides: true}), options)
	if err != nil {
		t.Fatalf("reordered extract: %v", err)
	}
	if len(reordered.Slides) != 2 || reordered.Slides[0].Source.PartName != "relocated/slides/slide-b.xml" {
		t.Fatalf("slide relationship order was not retained: %#v", reordered.Slides)
	}
	for _, slide := range reordered.Slides {
		expected := initialIDs[slide.Source.PartName]
		if slide.ID != expected[0] || slide.Elements[0].ID != expected[1] {
			t.Fatalf("identity changed across reorder for %q: got %q/%q want %q/%q", slide.Source.PartName, slide.ID, slide.Elements[0].ID, expected[0], expected[1])
		}
	}
}

func TestExtractNativePPTXFiltersPreviousIdentitiesToMatchedSlideAnchors(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{secondSlide: true})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	previous.DocumentID = "deck-partial-anchor"
	matchedSlideID := previous.Slides[0].ID
	matchedElementID := previous.Slides[0].Elements[0].ID
	oldSlideID := previous.Slides[1].ID
	oldElementID := previous.Slides[1].Elements[0].ID
	previous.Slides[1].ID = "slide-poisoned"
	previous.Slides[1].Source.ObjectID = "sldId-999"
	previous.Slides[1].Elements[0].ID = "element-poisoned"
	for _, compatibility := range []*NativeCompatibility{
		&previous.Compatibility,
		&previous.Slides[0].Compatibility,
		&previous.Slides[0].Elements[0].Compatibility,
		&previous.Slides[1].Compatibility,
		&previous.Slides[1].Elements[0].Compatibility,
	} {
		for index := range compatibility.Diagnostics {
			scope := compatibility.Diagnostics[index].Scope
			if scope == nil {
				continue
			}
			if scope.SlideID != nil && *scope.SlideID == oldSlideID {
				*scope.SlideID = previous.Slides[1].ID
			}
			if scope.ElementID != nil && *scope.ElementID == oldElementID {
				*scope.ElementID = previous.Slides[1].Elements[0].ID
			}
		}
	}
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("partial-anchor Previous must remain contract-valid: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	current, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("extract with partial-anchor Previous: %v", err)
	}
	if current.DocumentID != previous.DocumentID || current.Slides[0].ID != matchedSlideID || current.Slides[0].Elements[0].ID != matchedElementID {
		t.Fatalf("matched document/slide identities lost durability: %#v", current)
	}
	if current.Slides[1].ID == "slide-poisoned" || current.Slides[1].Elements[0].ID == "element-poisoned" {
		t.Fatalf("unmatched slide identities poisoned the current deck: %#v", current.Slides[1])
	}
}

func TestExtractNativePPTXRejectsElementIdentityFromUnmatchedSlideSharingMatchedPart(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{secondSlide: true})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	matchedPart := previous.Slides[0].Source.PartName
	matchedObjectID := previous.Slides[0].Elements[0].Source.ObjectID
	removedElementID := previous.Slides[0].Elements[0].ID
	previous.Slides[0].Elements = []NativeElement{}
	for _, compatibility := range []*NativeCompatibility{&previous.Compatibility, &previous.Slides[0].Compatibility} {
		filtered := compatibility.Diagnostics[:0]
		for _, diagnostic := range compatibility.Diagnostics {
			if diagnostic.Scope != nil && diagnostic.Scope.ElementID != nil && *diagnostic.Scope.ElementID == removedElementID {
				continue
			}
			filtered = append(filtered, diagnostic)
		}
		compatibility.Diagnostics = filtered
	}
	poisoned := &previous.Slides[1].Elements[0]
	oldElementID := poisoned.ID
	poisoned.ID = "element-cross-slide-poison"
	poisoned.Source.PartName = matchedPart
	poisoned.Source.ObjectID = matchedObjectID
	previous.Slides[1].Source.PartName = matchedPart
	previous.Slides[1].Source.ObjectID = "sldId-999"
	for _, compatibility := range []*NativeCompatibility{&previous.Compatibility, &previous.Slides[1].Compatibility, &poisoned.Compatibility} {
		for index := range compatibility.Diagnostics {
			scope := compatibility.Diagnostics[index].Scope
			if scope == nil {
				continue
			}
			if scope.ElementID != nil && *scope.ElementID == oldElementID {
				*scope.ElementID = poisoned.ID
			}
			if scope.PartName != nil && *scope.PartName == "relocated/slides/slide-b.xml" {
				*scope.PartName = matchedPart
			}
		}
	}
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("same-part poison Previous must remain contract-valid: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	current, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("extract with same-part poison Previous: %v", err)
	}
	if current.Slides[0].Elements[0].ID == poisoned.ID {
		t.Fatalf("unmatched slide sharing a matched part poisoned element identity: %#v", current.Slides[0].Elements[0])
	}
}

func TestExtractNativePPTXRejectsOPCAliasesAndTraversal(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name    string
		options nativeExtractFixtureOptions
	}{
		{name: "case alias", options: nativeExtractFixtureOptions{extraPartName: "RELOCATED/SLIDES/SLIDE-A.XML"}},
		{name: "percent alias", options: nativeExtractFixtureOptions{extraPartName: "relocated/slides/slide-%61.xml"}},
		{name: "encoded traversal target", options: nativeExtractFixtureOptions{slideTarget: "slides/%2E%2E/slide-a.xml"}},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			if _, err := ExtractNativePPTX(nativeExtractFixture(t, test.options), nativeTestExtractOptions()); err == nil {
				t.Fatal("expected extraction to reject hostile OPC package")
			}
		})
	}
}

func TestExtractNativePPTXRoutesCaseInsensitiveOPCNamesToActualSpelling(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{
		rename: map[string]string{
			"[Content_Types].xml":                     "[CONTENT_TYPES].XML",
			"_rels/.rels":                             "_RELS/.RELS",
			"relocated/deck.xml":                      "ReLoCaTeD/Deck.XML",
			"relocated/_rels/deck.xml.rels":           "RELOCATED/_RELS/DECK.XML.RELS",
			"relocated/slides/slide-a.xml":            "RELOCATED/SLIDES/SLIDE-A.XML",
			"relocated/slides/_rels/slide-a.xml.rels": "RELOCATED/SLIDES/_RELS/SLIDE-A.XML.RELS",
		},
		mutate: func(parts map[string]string) {
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `Target="relocated/deck.xml"`, `Target="RELOCATED/DECK.XML"`, 1)
			parts["relocated/_rels/deck.xml.rels"] = strings.Replace(parts["relocated/_rels/deck.xml.rels"], `Target="slides/slide-a.xml"`, `Target="SLIDES/SLIDE-A.XML"`, 1)
			parts["[Content_Types].xml"] = strings.ReplaceAll(parts["[Content_Types].xml"], "application/", "APPLICATION/")
		},
	})
	var relationshipOwners []string
	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		if request.Reason == "pptx.relationship-map-preserve" {
			relationshipOwners = append(relationshipOwners, request.OwnerPart)
		}
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})}
	deck, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("case-insensitive OPC route: %v", err)
	}
	if deck.Slides[0].Source == nil || deck.Slides[0].Source.PartName != "RELOCATED/SLIDES/SLIDE-A.XML" {
		t.Fatalf("source anchor did not preserve actual ZIP spelling: %#v", deck.Slides[0].Source)
	}
	for _, expected := range []string{"_RELS/.RELS", "RELOCATED/_RELS/DECK.XML.RELS", "RELOCATED/SLIDES/_RELS/SLIDE-A.XML.RELS"} {
		found := false
		for _, owner := range relationshipOwners {
			if owner == expected {
				found = true
			}
		}
		if !found {
			t.Fatalf("actual-spelling relationship map %q was not capability-bound: %#v", expected, relationshipOwners)
		}
	}
}

func TestExtractNativePPTXRejectsHostileContentTypeDeclarationsAndPartCharacters(t *testing.T) {
	t.Parallel()

	tests := []nativeExtractFixtureOptions{
		{extraPartName: "bad name.xml"},
		{extraPartName: "bad?.xml"},
		{extraPartName: "bad#.xml"},
		{mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/RELOCATED/DECK.XML" ContentType="application/xml"/></Types>`, 1)
		}},
		{mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/missing.xml" ContentType="application/xml"/></Types>`, 1)
		}},
		{mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `ContentType="application/xml"/>`, `ContentType=" "/>`, 1)
		}},
		{mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `ContentType="application/xml"/>`, `ContentType="application/xml"><Bad/></Default>`, 1)
		}},
		{mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `ContentType="application/vnd.openxmlformats-package.relationships+xml"`, `ContentType="application/xml"`, 1)
		}},
		{mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Default Extension="unused" ContentType="application/octet-stream"/></Types>`, 1)
		}},
	}
	for index, options := range tests {
		options := options
		t.Run(fmt.Sprintf("case-%d", index), func(t *testing.T) {
			t.Parallel()
			if _, err := ExtractNativePPTX(nativeExtractFixture(t, options), nativeTestExtractOptions()); err == nil {
				t.Fatal("expected hostile content type/part rejection")
			}
		})
	}
}

func TestExtractNativePPTXTargetModeIsExact(t *testing.T) {
	t.Parallel()

	if _, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{rootMode: "Internal", slideMode: "Internal"}), nativeTestExtractOptions()); err != nil {
		t.Fatalf("explicit Internal must be accepted: %v", err)
	}
	for _, mode := range []string{"internal", " INTERNAL", "External", "Unknown"} {
		mode := mode
		t.Run(mode, func(t *testing.T) {
			t.Parallel()
			if _, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{rootMode: mode}), nativeTestExtractOptions()); err == nil {
				t.Fatalf("expected exact TargetMode enforcement for %q", mode)
			}
		})
	}
}

func TestExtractNativePPTXRejectsXMLOutsideRootAndProcessingInstructions(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name   string
		mutate func(string) string
	}{
		{name: "processing instruction", mutate: func(value string) string { return `<?unsafe value?>` + value }},
		{name: "wrong XML version", mutate: func(value string) string { return `<?xml version="1.1"?>` + value }},
		{name: "unknown declaration field", mutate: func(value string) string { return `<?xml version="1.0" extra="value"?>` + value }},
		{name: "declaration fields out of order", mutate: func(value string) string { return `<?xml encoding="UTF-8" version="1.0"?>` + value }},
		{name: "duplicate declaration field", mutate: func(value string) string { return `<?xml version="1.0" version="1.0"?>` + value }},
		{name: "declaration after leading content", mutate: func(value string) string { return `<!--before--><?xml version="1.0"?>` + value }},
		{name: "later declaration", mutate: func(value string) string { return value + `<?xml version="1.0"?>` }},
		{name: "text before root", mutate: func(value string) string { return `outside` + value }},
		{name: "text after root", mutate: func(value string) string { return value + `outside` }},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/deck.xml"] = test.mutate(parts["relocated/deck.xml"])
			}})
			if _, err := ExtractNativePPTX(payload, nativeTestExtractOptions()); err == nil {
				t.Fatal("expected XML boundary rejection")
			}
		})
	}
}

func TestExtractNativePPTXRejectsDuplicateModeledSingletons(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name   string
		part   string
		mutate func(string) string
	}{
		{name: "slide size", part: "relocated/deck.xml", mutate: func(value string) string {
			return strings.Replace(value, `</p:presentation>`, `<p:sldSz cx="1" cy="1"/></p:presentation>`, 1)
		}},
		{name: "common slide data", part: "relocated/slides/slide-a.xml", mutate: func(value string) string {
			return strings.Replace(value, `</p:sld>`, `<p:cSld><p:spTree/></p:cSld></p:sld>`, 1)
		}},
		{name: "nonvisual id", part: "relocated/slides/slide-a.xml", mutate: func(value string) string {
			return strings.Replace(value, `<p:cNvSpPr`, `<p:cNvPr id="3"/><p:cNvSpPr`, 1)
		}},
		{name: "transform offset", part: "relocated/slides/slide-a.xml", mutate: func(value string) string { return strings.Replace(value, `<a:ext`, `<a:off x="0" y="0"/><a:ext`, 1) }},
		{name: "paragraph properties", part: "relocated/slides/slide-a.xml", mutate: func(value string) string { return strings.Replace(value, `<a:r>`, `<a:pPr/><a:r>`, 1) }},
		{name: "run properties", part: "relocated/slides/slide-a.xml", mutate: func(value string) string {
			return strings.Replace(value, `<a:t xml:space`, `<a:rPr/><a:t xml:space`, 1)
		}},
		{name: "run text", part: "relocated/slides/slide-a.xml", mutate: func(value string) string {
			return strings.Replace(value, `<a:t xml:space="preserve">Hello </a:t>`, `<a:t xml:space="preserve">Hello </a:t><a:t>duplicate</a:t>`, 1)
		}},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts[test.part] = test.mutate(parts[test.part])
			}})
			if _, err := ExtractNativePPTX(payload, nativeTestExtractOptions()); err == nil {
				t.Fatal("expected duplicate singleton rejection")
			}
		})
	}
}

func TestExtractNativePPTXRejectsDialectMixingAndNoncanonicalIDs(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name   string
		strict bool
		part   string
		mutate func(string) string
	}{
		{name: "office relationship type", strict: true, part: "_rels/.rels", mutate: func(value string) string {
			return strings.Replace(value, relOfficeDocumentStrict, relOfficeDocumentTransitional, 1)
		}},
		{name: "relationship attribute namespace", strict: true, part: "relocated/deck.xml", mutate: func(value string) string {
			return strings.Replace(value, nsOfficeRelsStrict, nsOfficeRelsTransitional, 1)
		}},
		{name: "slide id leading zero", part: "relocated/deck.xml", mutate: func(value string) string { return strings.Replace(value, `id="256"`, `id="0256"`, 1) }},
		{name: "shape id leading zero", part: "relocated/slides/slide-a.xml", mutate: func(value string) string { return strings.Replace(value, `id="2"`, `id="02"`, 1) }},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: test.strict, mutate: func(parts map[string]string) {
				parts[test.part] = test.mutate(parts[test.part])
			}})
			if _, err := ExtractNativePPTX(payload, nativeTestExtractOptions()); err == nil {
				t.Fatal("expected dialect/native id rejection")
			}
		})
	}
}

func TestExtractNativePPTXUnsupportedShapeUsesObjectLocalCapability(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `<a:bodyPr/>`, `<a:bodyPr vert="vert"/>`, 1)
	}})
	var requests []NativePassthroughTokenRequest
	deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		requests = append(requests, request)
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})})
	if err != nil {
		t.Fatalf("extract unsupported shape with capability issuer: %v", err)
	}
	var objectRequest *NativePassthroughTokenRequest
	for index := range requests {
		if requests[index].ObjectID == "cNvPr-2" {
			objectRequest = &requests[index]
		}
	}
	if objectRequest == nil || !strings.HasPrefix(string(objectRequest.Payload), `<p:sp>`) || strings.Contains(string(objectRequest.Payload), `<p:sld`) {
		t.Fatalf("passthrough was not bound to the smallest owning object: %#v", requests)
	}
	if objectRequest.FingerprintSHA256 != nativeSHA256(objectRequest.Payload) {
		t.Fatal("passthrough fingerprint does not bind the exact object subtree")
	}
	if len(deck.Slides[0].Elements) != 1 || deck.Slides[0].Elements[0].Compatibility.Status != NativeCompatibilityStatusRefused || len(deck.Slides[0].Elements[0].Passthrough) != 1 {
		t.Fatalf("unsupported object was not made explicit: %#v", deck.Slides[0])
	}
	if len(deck.Compatibility.Diagnostics) == 0 {
		t.Fatal("deck did not aggregate unsupported-content diagnostics")
	}
}

func TestExtractNativePPTXElementFingerprintUsesExactObjectSubtree(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract: %v", err)
	}
	element := deck.Slides[0].Elements[0]
	if element.Source == nil || deck.Slides[0].Source == nil || element.Source.FingerprintSHA256 == deck.Slides[0].Source.FingerprintSHA256 {
		t.Fatalf("element fingerprint is not object-local: element=%#v slide=%#v", element.Source, deck.Slides[0].Source)
	}
}

func TestExtractNativePPTXScopesInitialIDsToPresentationContent(t *testing.T) {
	t.Parallel()

	first, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("first extract: %v", err)
	}
	secondPayload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `cx="12192000"`, `cx="12192001"`, 1)
	}})
	second, err := ExtractNativePPTX(secondPayload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("second extract: %v", err)
	}
	if first.DocumentID == second.DocumentID || first.Slides[0].ID == second.Slides[0].ID || first.Slides[0].Elements[0].ID == second.Slides[0].Elements[0].ID {
		t.Fatalf("independent documents collided: first=%q/%q/%q second=%q/%q/%q", first.DocumentID, first.Slides[0].ID, first.Slides[0].Elements[0].ID, second.DocumentID, second.Slides[0].ID, second.Slides[0].Elements[0].ID)
	}
}

func TestExtractNativePPTXValidatesPreviousBeforeIdentityReuse(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	previous.ContractVersion = "pptx-native/invalid"
	options := nativeTestExtractOptions()
	options.Previous = &previous
	if _, err := ExtractNativePPTX(payload, options); err == nil {
		t.Fatal("invalid Previous deck was trusted")
	}
}

func TestExtractNativePPTXRejectsCapabilityCollisionAcrossObjects(t *testing.T) {
	t.Parallel()

	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(NativePassthroughTokenRequest) (string, error) {
		return "token-collision", nil
	})}
	if _, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{}), options); err == nil {
		t.Fatal("distinct passthrough objects reused one opaque capability")
	}
}

func TestExtractNativePPTXRejectsDuplicateIDsAcrossShapeKindsAndSlideOrder(t *testing.T) {
	t.Parallel()

	tests := []nativeExtractFixtureOptions{
		{mutate: func(parts map[string]string) {
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, `<p:pic><p:nvPicPr><p:cNvPr id="2"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic></p:spTree>`, 1)
		}},
		{mutate: func(parts map[string]string) {
			parts["relocated/deck.xml"] = strings.Replace(parts["relocated/deck.xml"], `</p:sldIdLst>`, `<p:sldId id="257" r:id="rId7"/></p:sldIdLst>`, 1)
		}},
	}
	for index, options := range tests {
		if _, err := ExtractNativePPTX(nativeExtractFixture(t, options), nativeTestExtractOptions()); err == nil {
			t.Fatalf("duplicate global native identity case %d was accepted", index)
		}
	}
}

func TestExtractNativePPTXPreservesUnmodeledCriticalMarkup(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:sld>`, `<p:timing/></p:sld>`, 1)
	}})
	var sawTiming bool
	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		if request.Reason == "pptx.unsupported-slide-property" && string(request.Payload) == `<p:timing/>` {
			sawTiming = true
		}
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})}
	deck, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("extract critical markup: %v", err)
	}
	if !sawTiming || deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("unmodeled timing was not explicitly preserved: saw=%v compatibility=%#v", sawTiming, deck.Slides[0].Compatibility)
	}
}

func TestNativePPTXExtractorHasNoDOMOrLegacyParserDependency(t *testing.T) {
	t.Parallel()

	for _, name := range []string{"native_extract.go", "native_extract_connector.go", "native_extract_graph.go", "native_extract_group.go", "native_extract_picture.go", "native_extract_shape.go", "native_extract_table.go"} {
		source, err := os.ReadFile(name)
		if err != nil {
			t.Fatalf("read %s: %v", name, err)
		}
		for _, forbidden := range []string{"ParsePPTX(", "localName(", "golang.org/x/net/html", "document.createElement", "DOMParser", "mammoth", "konva"} {
			if bytes.Contains(source, []byte(forbidden)) {
				t.Fatalf("%s contains forbidden legacy/DOM dependency %q", name, forbidden)
			}
		}
	}
}

func TestNativePartAliasUsesASCIICaseEquivalenceOnly(t *testing.T) {
	t.Parallel()

	upperASCII, err := nativePartAlias("RELOCATED/%53LIDES/SLIDE.XML")
	if err != nil {
		t.Fatal(err)
	}
	lowerASCII, err := nativePartAlias("relocated/slides/slide.xml")
	if err != nil {
		t.Fatal(err)
	}
	if upperASCII != lowerASCII {
		t.Fatalf("ASCII case/percent aliases differ: %q != %q", upperASCII, lowerASCII)
	}
	nonASCIIUpper, _ := nativePartAlias("media/Ä.png")
	nonASCIILower, _ := nativePartAlias("media/ä.png")
	if nonASCIIUpper == nonASCIILower {
		t.Fatalf("non-ASCII case variants were incorrectly aliased: %q", nonASCIIUpper)
	}
	types := nativeExtractContentTypes{defaults: map[string]string{"Ä": "application/test"}, overrides: map[string]string{}}
	if got := types.forPart("media/value.ä"); got != "" {
		t.Fatalf("non-ASCII extension case variant incorrectly matched Default: %q", got)
	}
}

func TestExtractNativePPTXNamespaceDeclarationsDoNotCreateRootCapabilities(t *testing.T) {
	t.Parallel()

	seen := map[string]bool{}
	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		seen[request.Reason] = true
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})}
	if _, err := ExtractNativePPTX(nativeExtractFixture(t, nativeExtractFixtureOptions{}), options); err != nil {
		t.Fatalf("extract: %v", err)
	}
	for _, reason := range []string{"pptx.unsupported-presentation-markup", "pptx.unsupported-slide-markup", "pptx.unsupported-layout-markup", "pptx.unsupported-master-markup"} {
		if seen[reason] {
			t.Fatalf("namespace declaration created spurious whole-root capability %q", reason)
		}
	}
	if !seen["pptx.unsupported-theme-markup"] {
		t.Fatal("real theme name attribute was not preserved")
	}
}

func TestExtractNativePPTXDistinctUnsupportedNodesAndDependencyShapeIDs(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:sld>`, `<p:timing/><p:timing/></p:sld>`, 1)
		parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], `</p:spTree>`, `<p:sp><p:nvSpPr><p:cNvPr id="9"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr></p:sp></p:spTree>`, 1)
	}})
	var timingIDs []string
	var layoutShapeID string
	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		if request.Reason == "pptx.unsupported-slide-property" {
			timingIDs = append(timingIDs, request.ObjectID)
		}
		if request.Reason == "pptx.unsupported-layout-shape" {
			layoutShapeID = request.ObjectID
		}
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})}
	if _, err := ExtractNativePPTX(payload, options); err != nil {
		t.Fatalf("extract: %v", err)
	}
	if len(timingIDs) != 2 || timingIDs[0] == timingIDs[1] {
		t.Fatalf("distinct identical timing nodes collapsed: %#v", timingIDs)
	}
	if layoutShapeID != "layout-cNvPr-9" {
		t.Fatalf("layout shape did not use its safe cNvPr identity: %q", layoutShapeID)
	}
}

func TestExtractNativePPTXRejectsOpposingRelationshipTypeDialect(t *testing.T) {
	t.Parallel()

	tests := []struct {
		part   string
		target string
	}{
		{part: "_rels/.rels", target: "relocated/themes/theme.xml"},
		{part: "relocated/_rels/deck.xml.rels", target: "themes/theme.xml"},
		{part: "relocated/slides/_rels/slide-a.xml.rels", target: "../themes/theme.xml"},
		{part: "relocated/layouts/_rels/layout.xml.rels", target: "../themes/theme.xml"},
		{part: "relocated/masters/_rels/master.xml.rels", target: "../themes/theme.xml"},
	}
	for _, test := range tests {
		test := test
		t.Run(test.part, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				addition := fmt.Sprintf(`<Relationship Id="rOpposite" Type="%s/image" Target="%s"/>`, nsOfficeRelsStrict, test.target)
				parts[test.part] = strings.Replace(parts[test.part], `</Relationships>`, addition+`</Relationships>`, 1)
			}})
			if _, err := ExtractNativePPTX(payload, nativeTestExtractOptions()); err == nil {
				t.Fatal("opposing Strict relationship Type was accepted in Transitional package")
			}
		})
	}
}

func TestNativePPTXExtractorBoundsSharedPassthroughReferenceFanout(t *testing.T) {
	t.Parallel()

	factoryCalls := 0
	extractor := nativeExtractor{
		options: NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(NativePassthroughTokenRequest) (string, error) {
			factoryCalls++
			return "token-shared", nil
		})},
		sourceRevision:   "rev-" + strings.Repeat("a", 64),
		passthroughCache: map[string]NativePassthroughRef{},
		tokenOwners:      map[string]string{},
	}
	payload := []byte(`<a:themeElements/>`)
	var gotErr error
	for index := 0; index <= nativeExtractMaxEmittedPassthrough; index++ {
		slide := NativeSlide{ID: fmt.Sprintf("slide-%d", index), Passthrough: []NativePassthroughRef{}, Compatibility: NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}}}
		gotErr = extractor.markSlideUnsupported(&slide, "relocated/themes/theme.xml", "theme-shared", nativeSHA256(payload), payload, "pptx.unsupported-theme-dependency", "shared theme")
		if gotErr != nil {
			break
		}
	}
	if gotErr == nil || extractor.passthroughRefsEmitted != nativeExtractMaxEmittedPassthrough || factoryCalls != 1 {
		t.Fatalf("shared dependency fanout was not bounded before factory/materialization: err=%v refs=%d factoryCalls=%d", gotErr, extractor.passthroughRefsEmitted, factoryCalls)
	}
}

func TestExtractNativePPTXRejectsDuplicatePreviousSourceAnchorKeys(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial extract: %v", err)
	}
	clone := previous.Slides[0].Elements[0]
	clone.ID = "element-duplicate-source"
	clone.Compatibility.Diagnostics = append([]NativeDiagnostic(nil), clone.Compatibility.Diagnostics...)
	if clone.Compatibility.Diagnostics == nil {
		clone.Compatibility.Diagnostics = []NativeDiagnostic{}
	}
	for index := range clone.Compatibility.Diagnostics {
		originalScope := clone.Compatibility.Diagnostics[index].Scope
		if originalScope != nil {
			copiedScope := *originalScope
			if originalScope.SlideID != nil {
				value := *originalScope.SlideID
				copiedScope.SlideID = &value
			}
			if originalScope.ElementID != nil {
				value := *originalScope.ElementID
				copiedScope.ElementID = &value
			}
			if originalScope.PartName != nil {
				value := *originalScope.PartName
				copiedScope.PartName = &value
			}
			clone.Compatibility.Diagnostics[index].Scope = &copiedScope
		}
		scope := clone.Compatibility.Diagnostics[index].Scope
		if scope != nil && scope.ElementID != nil {
			*scope.ElementID = clone.ID
		}
	}
	previous.Slides[0].Elements = append(previous.Slides[0].Elements, clone)
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("test Previous must remain contract-valid: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	if _, err := ExtractNativePPTX(payload, options); err == nil {
		t.Fatal("duplicate Previous source-anchor keys were trusted")
	}
}

type nativeExtractFixtureOptions struct {
	strict        bool
	extraPartName string
	slideTarget   string
	rootMode      string
	slideMode     string
	mutate        func(map[string]string)
	rename        map[string]string
	secondSlide   bool
	reverseSlides bool
	extraParts    []nativeExtractZipPart
}

func nativeExtractFixture(t *testing.T, options nativeExtractFixtureOptions) []byte {
	t.Helper()
	presentationNS := nsPresentationTransitional
	drawingNS := nsDrawingTransitional
	officeRelsNS := nsOfficeRelsTransitional
	packageRelsNS := nsPackageRels
	officeDocumentType := relOfficeDocumentTransitional
	slideType := relSlideTransitional
	if options.strict {
		presentationNS = nsPresentationStrict
		drawingNS = nsDrawingStrict
		officeRelsNS = nsOfficeRelsStrict
		officeDocumentType = relOfficeDocumentStrict
		slideType = relSlideStrict
	}
	presentationPart := "relocated/deck.xml"
	slidePart := "relocated/slides/slide-a.xml"
	layoutPart := "relocated/layouts/layout.xml"
	masterPart := "relocated/masters/master.xml"
	themePart := "relocated/themes/theme.xml"
	slideTarget := "slides/slide-a.xml"
	if options.slideTarget != "" {
		slideTarget = options.slideTarget
	}

	rootMode := ""
	if options.rootMode != "" {
		rootMode = fmt.Sprintf(` TargetMode="%s"`, options.rootMode)
	}
	slideMode := ""
	if options.slideMode != "" {
		slideMode = fmt.Sprintf(` TargetMode="%s"`, options.slideMode)
	}
	parts := []nativeExtractZipPart{
		{name: "[Content_Types].xml", data: fmt.Sprintf(`<Types xmlns="%s"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/%s" ContentType="%s"/><Override PartName="/%s" ContentType="%s"/><Override PartName="/%s" ContentType="%s"/><Override PartName="/%s" ContentType="%s"/><Override PartName="/%s" ContentType="%s"/></Types>`, nsContentTypes, presentationPart, contentTypePresentation, slidePart, contentTypeSlide, layoutPart, contentTypeSlideLayout, masterPart, contentTypeSlideMaster, themePart, contentTypeTheme)},
		{name: "_rels/.rels", data: fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdRoot" Type="%s" Target="%s"%s/></Relationships>`, packageRelsNS, officeDocumentType, presentationPart, rootMode)},
		{name: presentationPart, data: fmt.Sprintf(`<p:presentation xmlns:p="%s" xmlns:r="%s"><p:sldIdLst><p:sldId id="256" r:id="rId7"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`, presentationNS, officeRelsNS)},
		{name: "relocated/_rels/deck.xml.rels", data: fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rId7" Type="%s" Target="%s"%s/></Relationships>`, packageRelsNS, slideType, slideTarget, slideMode)},
		{name: slidePart, data: fmt.Sprintf(`<p:sld xmlns:p="%s" xmlns:a="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="457200"/><a:ext cx="4572000" cy="914400"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="1" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t xml:space="preserve">Hello </a:t></a:r><a:r><a:rPr b="0" i="0" sz="3200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>world</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`, presentationNS, drawingNS)},
		{name: "relocated/slides/_rels/slide-a.xml.rels", data: fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdLayout" Type="%s" Target="../layouts/layout.xml"/></Relationships>`, packageRelsNS, map[bool]string{false: relSlideLayoutTransitional, true: relSlideLayoutStrict}[options.strict])},
		{name: layoutPart, data: fmt.Sprintf(`<p:sldLayout xmlns:p="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`, presentationNS)},
		{name: "relocated/layouts/_rels/layout.xml.rels", data: fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdMaster" Type="%s" Target="../masters/master.xml"/></Relationships>`, packageRelsNS, map[bool]string{false: relSlideMasterTransitional, true: relSlideMasterStrict}[options.strict])},
		{name: masterPart, data: fmt.Sprintf(`<p:sldMaster xmlns:p="%s"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>`, presentationNS)},
		{name: "relocated/masters/_rels/master.xml.rels", data: fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdTheme" Type="%s" Target="../themes/theme.xml"/></Relationships>`, packageRelsNS, map[bool]string{false: relThemeTransitional, true: relThemeStrict}[options.strict])},
		{name: themePart, data: fmt.Sprintf(`<a:theme xmlns:a="%s" name="Fixture"><a:themeElements/></a:theme>`, drawingNS)},
	}
	if options.secondSlide {
		secondSlidePart := "relocated/slides/slide-b.xml"
		for index := range parts {
			switch parts[index].name {
			case "[Content_Types].xml":
				parts[index].data = strings.Replace(parts[index].data, `</Types>`, fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/></Types>`, secondSlidePart, contentTypeSlide), 1)
			case presentationPart:
				ordered := `<p:sldId id="256" r:id="rId7"/><p:sldId id="257" r:id="rId8"/>`
				if options.reverseSlides {
					ordered = `<p:sldId id="257" r:id="rId8"/><p:sldId id="256" r:id="rId7"/>`
				}
				parts[index].data = strings.Replace(parts[index].data, `<p:sldId id="256" r:id="rId7"/>`, ordered, 1)
			case "relocated/_rels/deck.xml.rels":
				parts[index].data = strings.Replace(parts[index].data, `</Relationships>`, fmt.Sprintf(`<Relationship Id="rId8" Type="%s" Target="slides/slide-b.xml"/></Relationships>`, slideType), 1)
			}
		}
		secondSlideXML := ""
		for _, part := range parts {
			if part.name == slidePart {
				secondSlideXML = strings.Replace(part.data, `name="Title"`, `name="Second"`, 1)
				secondSlideXML = strings.Replace(secondSlideXML, `Hello `, `Second `, 1)
				break
			}
		}
		parts = append(parts,
			nativeExtractZipPart{name: secondSlidePart, data: secondSlideXML},
			nativeExtractZipPart{name: "relocated/slides/_rels/slide-b.xml.rels", data: fmt.Sprintf(`<Relationships xmlns="%s"><Relationship Id="rIdLayout" Type="%s" Target="../layouts/layout.xml"/></Relationships>`, packageRelsNS, map[bool]string{false: relSlideLayoutTransitional, true: relSlideLayoutStrict}[options.strict])},
		)
	}
	parts = append(parts, options.extraParts...)
	if options.mutate != nil {
		byName := make(map[string]string, len(parts))
		for _, part := range parts {
			byName[part.name] = part.data
		}
		options.mutate(byName)
		for index := range parts {
			parts[index].data = byName[parts[index].name]
		}
	}
	for index := range parts {
		if renamed := options.rename[parts[index].name]; renamed != "" {
			parts[index].name = renamed
		}
	}
	if options.extraPartName != "" {
		parts = append(parts, nativeExtractZipPart{name: options.extraPartName, data: "collision"})
	}
	return writeNativeExtractZip(t, parts)
}

type nativeExtractZipPart struct {
	name string
	data string
}

func writeNativeExtractZip(t *testing.T, parts []nativeExtractZipPart) []byte {
	t.Helper()
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	for _, part := range parts {
		header := &zip.FileHeader{Name: part.name, Method: zip.Store}
		header.SetMode(0o600)
		entry, err := writer.CreateHeader(header)
		if err != nil {
			t.Fatalf("create ZIP part %q: %v", part.name, err)
		}
		if _, err := entry.Write([]byte(part.data)); err != nil {
			t.Fatalf("write ZIP part %q: %v", part.name, err)
		}
	}
	if err := writer.Close(); err != nil {
		t.Fatalf("close ZIP: %v", err)
	}
	return output.Bytes()
}

func nativeTestExtractOptions() NativePPTXExtractOptions {
	return NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		digest := nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))
		return "token-" + digest[:24], nil
	})}
}
