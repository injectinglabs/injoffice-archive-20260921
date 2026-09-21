package pptxpatch

import (
	"bytes"
	"encoding/json"
	"strings"
	"testing"
)

func TestNativeSlideMutationsRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			original := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, secondSlide: true})
			before, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			slide := before.Slides[0]
			insert := NativePPTXMutation{OperationID: "insert", Kind: NativePPTXInsertSlide, SlideID: slide.ID, ExpectedFingerprintSHA256: slide.Source.FingerprintSHA256}
			produced := applySlideTestPayload(t, original, before, insert)
			deck, err := ExtractNativePPTX(produced, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if len(deck.Slides) != 3 || deck.Slides[2].Source.PartName != before.Slides[1].Source.PartName {
				t.Fatal("insert did not preserve order")
			}
			pkg, err := openNativeExtractPackage(produced)
			if err != nil {
				t.Fatal(err)
			}
			inserted := deck.Slides[1]
			if len(inserted.Elements) != 0 {
				t.Fatal("insertion copied existing slide content")
			}
			rels := string(pkg.parts[nativeRelationshipsPart(inserted.Source.PartName)])
			if !strings.Contains(rels, `Target="../layouts/layout.xml"`) {
				t.Fatalf("layout not retained: %s", rels)
			}
			oldPkg, _ := openNativeExtractPackage(original)
			for part, data := range oldPkg.parts {
				if part == "relocated/deck.xml" || part == "relocated/_rels/deck.xml.rels" || part == "[Content_Types].xml" {
					continue
				}
				if !bytes.Equal(data, pkg.parts[part]) {
					t.Fatalf("insert changed %s", part)
				}
			}
			// Insert again after the same anchor: IDs and OPC names must not collide.
			insert.SlideID = deck.Slides[0].ID
			insert.ExpectedFingerprintSHA256 = deck.Slides[0].Source.FingerprintSHA256
			again := applySlideTestPayload(t, produced, deck, insert)
			againDeck, err := ExtractNativePPTX(again, nativeMutationExtractOptions())
			if err != nil || len(againDeck.Slides) != 4 {
				t.Fatalf("repeated insert: %v", err)
			}
			fill := "2459AD"
			background := NativePPTXMutation{OperationID: "background", Kind: NativePPTXSetSlideBackground, SlideID: inserted.ID, ExpectedFingerprintSHA256: inserted.Source.FingerprintSHA256, Fill: &fill}
			colored := applySlideTestPayload(t, produced, deck, background)
			colorDeck, err := ExtractNativePPTX(colored, nativeMutationExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if nativeStringValue(colorDeck.Slides[1].Background) != fill {
				t.Fatal("background missing after reopen")
			}
			coloredPkg, _ := openNativeExtractPackage(colored)
			for part, data := range pkg.parts {
				if part != inserted.Source.PartName && !bytes.Equal(data, coloredPkg.parts[part]) {
					t.Fatalf("background changed %s", part)
				}
			}
			// Replace an existing background, leaving exactly one bg before spTree.
			fill = "202B3C"
			background.SlideID = colorDeck.Slides[1].ID
			background.ExpectedFingerprintSHA256 = colorDeck.Slides[1].Source.FingerprintSHA256
			recolored := applySlideTestPayload(t, colored, colorDeck, background)
			finalPkg, _ := openNativeExtractPackage(recolored)
			raw := string(finalPkg.parts[inserted.Source.PartName])
			if strings.Count(raw, "<p:bg ") != 1 || strings.Index(raw, "<p:bg ") > strings.Index(raw, "<p:spTree>") {
				t.Fatal("invalid background ordering/cardinality")
			}
		})
	}
}

func applySlideTestPayload(t *testing.T, original []byte, deck NativePPTXDeck, operation NativePPTXMutation) []byte {
	t.Helper()
	request := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{operation}}
	payload, err := json.Marshal(request)
	if err != nil {
		t.Fatal(err)
	}
	result, err := ApplyNativePPTXMutationPayload(original, "sha256:"+nativeSHA256(original), payload)
	if err != nil {
		t.Fatal(err)
	}
	return result
}

func TestNativeSlideMutationRefusals(t *testing.T) {
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{})
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	slide := deck.Slides[0]
	fill := "2459AD"
	valid := NativePPTXMutation{OperationID: "bg", Kind: NativePPTXSetSlideBackground, SlideID: slide.ID, ExpectedFingerprintSHA256: slide.Source.FingerprintSHA256, Fill: &fill}
	for _, test := range []struct {
		name   string
		mutate func(*NativePPTXMutationRequest)
	}{
		{"stale revision", func(r *NativePPTXMutationRequest) { r.ExpectedSourceRevision = "rev-" + strings.Repeat("0", 64) }},
		{"stale fingerprint", func(r *NativePPTXMutationRequest) {
			r.Operations[0].ExpectedFingerprintSHA256 = strings.Repeat("0", 64)
		}},
		{"missing slide", func(r *NativePPTXMutationRequest) { r.Operations[0].SlideID = "missing" }},
		{"element field", func(r *NativePPTXMutationRequest) { r.Operations[0].ElementID = "element" }},
		{"invalid color", func(r *NativePPTXMutationRequest) { bad := "red"; r.Operations[0].Fill = &bad }},
		{"mixed batch", func(r *NativePPTXMutationRequest) { r.Operations = append(r.Operations, valid) }},
		{"insert fill", func(r *NativePPTXMutationRequest) { r.Operations[0].Kind = NativePPTXInsertSlide }},
	} {
		t.Run(test.name, func(t *testing.T) {
			r := NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: []NativePPTXMutation{valid}}
			test.mutate(&r)
			out, err := ApplyNativePPTXMutations(original, r)
			if err == nil || out != nil {
				t.Fatal("invalid request produced a package")
			}
		})
	}
}

func TestNativeSlideBackgroundPreservesSourceAndOverridesLayout(t *testing.T) {
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/layouts/layout.xml"] = strings.Replace(parts["relocated/layouts/layout.xml"], "<p:cSld>", `<p:cSld><p:bg><p:bgPr><a:solidFill xmlns:a="`+nsDrawingTransitional+`"><a:srgbClr val="AABBCC"/></a:solidFill></p:bgPr></p:bg>`, 1)
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], "</p:cSld>", `<!--untouched opaque marker--></p:cSld>`, 1)
	}})
	deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	slide := deck.Slides[0]
	insert := NativePPTXMutation{OperationID: "insert", Kind: NativePPTXInsertSlide, SlideID: slide.ID, ExpectedFingerprintSHA256: slide.Source.FingerprintSHA256}
	inserted := applySlideTestPayload(t, original, deck, insert)
	after, err := ExtractNativePPTX(inserted, nativeMutationExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	if nativeStringValue(after.Slides[1].Background) != "AABBCC" {
		t.Fatal("new slide lost inherited layout background")
	}
	fill := "112233"
	operation := NativePPTXMutation{OperationID: "background", Kind: NativePPTXSetSlideBackground, SlideID: slide.ID, ExpectedFingerprintSHA256: slide.Source.FingerprintSHA256, Fill: &fill}
	colored := applySlideTestPayload(t, original, deck, operation)
	pkg, _ := openNativeExtractPackage(colored)
	if !bytes.Contains(pkg.parts[slide.Source.PartName], []byte("<!--untouched opaque marker-->")) {
		t.Fatal("background discarded opaque source XML")
	}
}

func TestNativeSlideInsertRefusesMissingLayout(t *testing.T) {
	original := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/_rels/slide-a.xml.rels"] = `<Relationships xmlns="` + nsPackageRels + `"/>`
	}})
	request := NativePPTXMutationRequest{ExpectedSourceRevision: "rev-" + nativeSHA256(original), Operations: []NativePPTXMutation{{OperationID: "insert", Kind: NativePPTXInsertSlide, SlideID: "slide", ExpectedFingerprintSHA256: strings.Repeat("a", 64)}}}
	if out, err := ApplyNativePPTXMutations(original, request); out != nil || err == nil || !strings.Contains(err.Error(), "layout") {
		t.Fatalf("missing layout: %v", err)
	}
}
