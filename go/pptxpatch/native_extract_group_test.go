package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXProjectsNestedGroupsExactlyInBothDialects(t *testing.T) {
	t.Parallel()

	for _, test := range []struct {
		name   string
		strict bool
	}{
		{name: "transitional"},
		{name: "strict", strict: true},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			presentationNS, drawingNS := nativeGroupTestNamespaces(test.strict)
			transformAttrs := ""
			if test.strict {
				transformAttrs = ` rot="0" flipH="false" flipV="0"`
			}
			groupXML := nativeNestedGroupXML(presentationNS, drawingNS, transformAttrs)
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: test.strict, mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, groupXML+`</p:spTree>`, 1)
			}})
			deck, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatalf("extract nested group: %v", err)
			}
			if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 2 {
				t.Fatalf("group changed slide z-order or disappeared: %#v", deck.Slides)
			}
			outer := deck.Slides[0].Elements[1]
			assertNativeGroupElement(t, outer, "Outer", "cNvPr-3", NativeTransform{
				X: int64Pointer(1_000_000), Y: int64Pointer(2_000_000), Cx: int64Pointer(6_000_000), Cy: int64Pointer(8_000_000),
			}, NativeTransform{
				X: int64Pointer(100), Y: int64Pointer(200), Cx: int64Pointer(300), Cy: int64Pointer(400),
			})
			if len(outer.Children) != 2 || outer.Children[0].Kind != NativeElementKindGroup || outer.Children[1].Kind != NativeElementKindShape {
				t.Fatalf("group child z-order was not source order: %#v", outer.Children)
			}
			inner := outer.Children[0]
			assertNativeGroupElement(t, inner, "Inner", "cNvPr-4", NativeTransform{
				X: int64Pointer(150), Y: int64Pointer(250), Cx: int64Pointer(100), Cy: int64Pointer(100),
			}, NativeTransform{
				X: int64Pointer(10), Y: int64Pointer(20), Cx: int64Pointer(50), Cy: int64Pointer(50),
			})
			if len(inner.Children) != 1 {
				t.Fatalf("nested group children: %#v", inner.Children)
			}
			assertNativeProjectedTransform(t, inner.Children[0], "cNvPr-5", NativeTransform{
				X: int64Pointer(20), Y: int64Pointer(30), Cx: int64Pointer(10), Cy: int64Pointer(5),
			})
			assertNativeProjectedTransform(t, outer.Children[1], "cNvPr-6", NativeTransform{
				X: int64Pointer(250), Y: int64Pointer(300), Cx: int64Pointer(25), Cy: int64Pointer(20),
			})
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("projected deck is not contract-valid: %#v", issues)
			}

			repeated, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatalf("repeat extract: %v", err)
			}
			if repeated.Slides[0].Elements[1].ID != outer.ID || repeated.Slides[0].Elements[1].Children[0].ID != inner.ID {
				t.Fatalf("group IDs are not stable: first=%q/%q second=%q/%q", outer.ID, inner.ID, repeated.Slides[0].Elements[1].ID, repeated.Slides[0].Elements[1].Children[0].ID)
			}
		})
	}
}

func TestValidateNativePPTXRejectsInexactGroupAffine(t *testing.T) {
	t.Parallel()

	zero, one, two, three := int64(0), int64(1), int64(2), int64(3)
	compatibility := NativeCompatibility{Status: NativeCompatibilityStatusEditable, Diagnostics: []NativeDiagnostic{}}
	element := NativeElement{
		Kind: NativeElementKindGroup, ID: "group", Provenance: NativeProvenanceAuthored,
		Transform:      NativeTransform{X: &zero, Y: &zero, Cx: &two, Cy: &two},
		ChildTransform: &NativeTransform{X: &zero, Y: &zero, Cx: &three, Cy: &three},
		Children: []NativeElement{{
			Kind: NativeElementKindConnector, ID: "leaf", Provenance: NativeProvenanceAuthored,
			Transform:   NativeTransform{X: &zero, Y: &zero, Cx: &one, Cy: &one},
			Passthrough: []NativePassthroughRef{}, Compatibility: compatibility,
		}},
		Passthrough: []NativePassthroughRef{}, Compatibility: compatibility,
	}
	deck := NativePPTXDeck{
		ContractVersion: NativePPTXContractVersion, DocumentID: "deck-group-affine", Origin: NativeOriginAuthored,
		Size: NativeSize{Cx: &one, Cy: &one}, Assets: []NativeAsset{},
		Slides: []NativeSlide{{
			ID: "slide-group-affine", Provenance: NativeProvenanceAuthored, Elements: []NativeElement{element},
			Passthrough: []NativePassthroughRef{}, Compatibility: compatibility,
		}},
		Compatibility: compatibility,
	}
	issues := ValidateNativePPTX(deck)
	for _, issue := range issues {
		if issue.Code == "native.groupTransform" {
			return
		}
	}
	t.Fatalf("expected inexact group transform refusal, got %#v", issues)
}

func TestExtractNativePPTXPreservesUnsupportedGroupAsOneOpaqueSubtree(t *testing.T) {
	t.Parallel()

	presentationNS, drawingNS := nativeGroupTestNamespaces(false)
	for _, test := range []struct {
		name      string
		xfrmAttrs string
		mutate    func(string) string
		code      string
	}{
		{name: "rotation", xfrmAttrs: ` rot="60000"`, code: "pptx.group-rotation-unavailable"},
		{name: "horizontal flip", xfrmAttrs: ` flipH="1"`, code: "pptx.group-flip-unavailable"},
		{name: "fractional EMU", code: "pptx.group-fractional-transform-unavailable", mutate: func(value string) string {
			return strings.Replace(value, `cx="6000000" cy="8000000"`, `cx="6000001" cy="8000000"`, 1)
		}},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			groupXML := nativeNestedGroupXML(presentationNS, drawingNS, test.xfrmAttrs)
			if test.mutate != nil {
				groupXML = test.mutate(groupXML)
			}
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, groupXML+`</p:spTree>`, 1)
			}})
			var groupRequests []NativePassthroughTokenRequest
			options := NativePPTXExtractOptions{TokenFactory: nativeAtomicTestTokenFactory(NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
				if request.Reason == test.code {
					groupRequests = append(groupRequests, request)
				}
				return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
			}))}
			deck, err := ExtractNativePPTX(payload, options)
			if err != nil {
				t.Fatalf("opaque group extraction: %v", err)
			}
			if len(deck.Slides[0].Elements) != 1 || deck.Slides[0].Elements[0].Source == nil || deck.Slides[0].Elements[0].Source.ObjectID != "cNvPr-2" {
				t.Fatalf("unsupported group leaked a partial projection: %#v", deck.Slides[0].Elements)
			}
			if len(groupRequests) != 1 || groupRequests[0].ObjectID != "cNvPr-3" || !strings.HasPrefix(string(groupRequests[0].Payload), `<p:grpSp`) || strings.Contains(string(groupRequests[0].Payload), `<p:sld`) {
				t.Fatalf("group was not capability-bound as one exact object-local subtree: %#v", groupRequests)
			}
			if groupRequests[0].FingerprintSHA256 != nativeSHA256(groupRequests[0].Payload) {
				t.Fatal("opaque group fingerprint does not bind its exact bytes")
			}
			if deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, test.code) {
				t.Fatalf("unsupported group was not explicit: %#v", deck.Slides[0].Compatibility)
			}
		})
	}
}

func TestExtractNativePPTXPreservesSchemaValidImplicitGroupTransforms(t *testing.T) {
	t.Parallel()

	for _, test := range []struct {
		name       string
		properties string
	}{
		{name: "missing transform", properties: `<p:grpSpPr/>`},
		{name: "empty transform", properties: `<p:grpSpPr><a:xfrm/></p:grpSpPr>`},
		{name: "partial transform", properties: `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/></a:xfrm></p:grpSpPr>`},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			group := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Implicit Transform"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>%s%s</p:grpSp>`, test.properties, nativeGroupRectXML(4, "Leaf", 0, 0, 1, 1))
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
			}})
			deck, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatalf("implicit group transform should be preserved, not reject the deck: %v", err)
			}
			if len(deck.Slides[0].Elements) != 1 || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.group-transform-unavailable") {
				t.Fatalf("implicit group transform was not one opaque subtree: %#v", deck.Slides[0])
			}
		})
	}
}

func TestExtractNativePPTXBubblesRefusedChildToOneExactGroupCapability(t *testing.T) {
	t.Parallel()

	refusedChild := strings.Replace(nativeGroupRectXML(4, "Refused Leaf", 0, 0, 10, 10), `</p:spPr>`, `<a:effectLst><a:blur rad="1"/></a:effectLst></p:spPr>`, 1)
	group := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Atomic Refused Child"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>%s%s</p:grpSp>`, nativeGroupRectXML(5, "Accepted Sibling", 20, 0, 10, 10), refusedChild)
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
	}})
	issuer := nativeAtomicTestTokenFactory(nativeTestExtractOptions().TokenFactory)
	deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: issuer})
	if err != nil {
		t.Fatalf("refused group child: %v", err)
	}
	if len(deck.Slides[0].Elements) != 1 || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.group-child-refused-unavailable") {
		t.Fatalf("refused child leaked a partial group projection: %#v", deck.Slides[0])
	}
	groupRequests := 0
	for _, request := range issuer.published {
		if request.Reason == "pptx.autoshape-refused" {
			t.Fatalf("refused descendant capability escaped group rollback: %#v", request)
		}
		if request.Reason == "pptx.group-child-refused-unavailable" {
			groupRequests++
			if request.ObjectID != "cNvPr-3" || !strings.HasPrefix(string(request.Payload), `<p:grpSp`) || !strings.Contains(string(request.Payload), `a:blur`) || request.FingerprintSHA256 != nativeSHA256(request.Payload) {
				t.Fatalf("group refusal is not bound to the exact complete subtree: %#v", request)
			}
		}
	}
	if groupRequests != 1 {
		t.Fatalf("expected exactly one atomic group capability, got %d: %#v", groupRequests, issuer.published)
	}
}

func TestExtractNativePPTXProjectsSubEmuGroupWorldPositionsAndUnicodeNames(t *testing.T) {
	t.Parallel()

	name := strings.Repeat("é", 600)
	group := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="%s"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3" cy="3"/><a:chOff x="0" y="0"/><a:chExt cx="2" cy="2"/></a:xfrm></p:grpSpPr>%s</p:grpSp>`, name, nativeGroupRectXML(4, "Half EMU Leaf", 1, 1, 1, 1))
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
	}})
	deck, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatalf("sub-EMU group projection: %v", err)
	}
	if len(deck.Slides[0].Elements) != 2 || deck.Slides[0].Elements[1].Kind != NativeElementKindGroup || deck.Slides[0].Elements[1].Name == nil || *deck.Slides[0].Elements[1].Name != name {
		t.Fatalf("exact group projection or Unicode name was lost: %#v", deck.Slides[0].Elements)
	}
}

func TestExtractNativePPTXNeverProjectsHiddenGroupedDescendants(t *testing.T) {
	t.Parallel()

	for _, dialect := range []struct {
		name   string
		strict bool
	}{{name: "transitional"}, {name: "strict", strict: true}} {
		dialect := dialect
		t.Run(dialect.name, func(t *testing.T) {
			t.Parallel()
			presentationNS, drawingNS := nativeGroupTestNamespaces(dialect.strict)
			base := nativeNestedGroupXML(presentationNS, drawingNS, "")
			picture := fmt.Sprintf(`<p:grpSp xmlns:p="%s" xmlns:a="%s"><p:nvGrpSpPr><p:cNvPr id="3" name="Hidden Picture Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="4" name="Hidden Picture" hidden="1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr></p:pic></p:grpSp>`, presentationNS, drawingNS)
			decorative := strings.Replace(base, `<p:cNvPr id="6" name="Outer Rect"/>`, `<p:cNvPr id="6" name="Outer Rect" hidden="true"><a:extLst><a:ext uri="decorative"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext></a:extLst></p:cNvPr>`, 1)
			for _, test := range []struct {
				name           string
				group          string
				wantDecorative bool
			}{
				{name: "shape", group: strings.Replace(base, `<p:cNvPr id="6" name="Outer Rect"/>`, `<p:cNvPr id="6" name="Outer Rect" hidden="1"/>`, 1)},
				{name: "picture", group: picture},
				{name: "nested-group", group: strings.Replace(base, `<p:cNvPr id="4" name="Inner"/>`, `<p:cNvPr id="4" name="Inner" hidden="true"/>`, 1)},
				{name: "decorative-shape", group: decorative, wantDecorative: true},
			} {
				test := test
				t.Run(test.name, func(t *testing.T) {
					payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: dialect.strict, mutate: func(parts map[string]string) {
						parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, test.group+`</p:spTree>`, 1)
					}})
					issuer := nativeAtomicTestTokenFactory(nativeTestExtractOptions().TokenFactory)
					deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: issuer})
					if err != nil {
						t.Fatalf("hidden grouped descendant: %v", err)
					}
					if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 1 || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.group-hidden-child-unavailable") {
						t.Fatalf("hidden descendant was projected or not diagnosed: %#v", deck.Slides)
					}
					foundOpaque := false
					for _, request := range issuer.published {
						if request.Reason == "pptx.group-hidden-child-unavailable" {
							foundOpaque = strings.HasPrefix(string(request.Payload), `<p:grpSp`) && (!test.wantDecorative || strings.Contains(string(request.Payload), "decorative"))
						}
					}
					if !foundOpaque {
						t.Fatalf("hidden subtree was not preserved exactly: %#v", issuer.published)
					}
				})
			}
		})
	}
}

func TestExtractNativePPTXRejectsMalformedOrOverdeepGroupsWithoutDeck(t *testing.T) {
	t.Parallel()

	presentationNS, drawingNS := nativeGroupTestNamespaces(false)
	base := nativeNestedGroupXML(presentationNS, drawingNS, "")
	tests := []struct {
		name  string
		group string
	}{
		{name: "duplicate child extent", group: strings.Replace(base, `</a:xfrm>`, `<a:chExt cx="300" cy="400"/></a:xfrm>`, 1)},
		{name: "zero child extent", group: strings.Replace(base, `chExt cx="300"`, `chExt cx="0"`, 1)},
		{name: "noncanonical coordinate", group: strings.Replace(base, `chOff x="100"`, `chOff x="0100"`, 1)},
		{name: "misordered transform", group: strings.Replace(base, `<a:off x="1000000" y="2000000"/><a:ext cx="6000000" cy="8000000"/>`, `<a:ext cx="6000000" cy="8000000"/><a:off x="1000000" y="2000000"/>`, 1)},
		{name: "over depth budget", group: nativeDeepGroupXML(presentationNS, drawingNS, nativeMaxDepth+1)},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, test.group+`</p:spTree>`, 1)
			}})
			deck, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
			if err == nil || len(deck.Slides) != 0 || len(deck.Assets) != 0 {
				t.Fatalf("malformed/overdeep group returned a partial deck: deck=%#v err=%v", deck, err)
			}
		})
	}
}

func TestExtractNativePPTXOpaqueGroupRollsBackNestedPictureProjection(t *testing.T) {
	t.Parallel()

	imagePart := "relocated/media/group.png"
	imageData := "\x89PNG\r\n\x1a\ngroup"
	group := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Picture Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="4" name="Nested Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="%s" r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="0"/><a:ext cx="10" cy="10"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic><p:grpSp><p:nvGrpSpPr><p:cNvPr id="5" name="Fractional Nested"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="101" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="3" cy="100"/></a:xfrm></p:grpSpPr>%s</p:grpSp></p:grpSp>`, nsOfficeRelsTransitional, nativeGroupRectXML(6, "Nested Rect", 0, 0, 1, 1))
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{
		extraParts: []nativeExtractZipPart{{name: imagePart, data: imageData}},
		mutate: func(parts map[string]string) {
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, fmt.Sprintf(`<Override PartName="/%s" ContentType="image/png"/></Types>`, imagePart), 1)
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
			parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`, fmt.Sprintf(`<Relationship Id="rIdImage" Type="%s" Target="../media/group.png"/></Relationships>`, relImageTransitional), 1)
		},
	})
	deck, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatalf("opaque nested-picture group: %v", err)
	}
	if len(deck.Assets) != 0 || len(deck.Slides[0].Elements) != 1 || nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.group-fractional-transform-unavailable") == false {
		t.Fatalf("opaque group leaked nested picture projection: assets=%#v elements=%#v compatibility=%#v", deck.Assets, deck.Slides[0].Elements, deck.Slides[0].Compatibility)
	}
}

func TestExtractNativePPTXProjectsStrictGroupedPictureAndRejectsOpposingRelationshipNamespace(t *testing.T) {
	t.Parallel()

	imagePart := "relocated/media/group-strict.png"
	imageData := "\x89PNG\r\n\x1a\nstrict-group"
	group := func(relationshipNamespace, extraRelationshipAttribute string) string {
		return fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Strict Picture Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr><p:pic><p:nvPicPr><p:cNvPr id="4" name="Strict Grouped Picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="%s" r:embed="rIdGroupImage"%s/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="10" cy="20"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:grpSp>`, relationshipNamespace, extraRelationshipAttribute)
	}
	fixture := func(t *testing.T, relationshipNamespace, extraRelationshipAttribute string) []byte {
		t.Helper()
		return nativeExtractFixture(t, nativeExtractFixtureOptions{
			strict:     true,
			extraParts: []nativeExtractZipPart{{name: imagePart, data: imageData}},
			mutate: func(parts map[string]string) {
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, fmt.Sprintf(`<Override PartName="/%s" ContentType="image/png"/></Types>`, imagePart), 1)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group(relationshipNamespace, extraRelationshipAttribute)+`</p:spTree>`, 1)
				parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`, fmt.Sprintf(`<Relationship Id="rIdGroupImage" Type="%s" Target="../MEDIA/GROUP-STRICT.PNG"/></Relationships>`, relImageStrict), 1)
			},
		})
	}

	deck, err := ExtractNativePPTX(fixture(t, nsOfficeRelsStrict, ""), nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatalf("strict grouped picture: %v", err)
	}
	if len(deck.Assets) != 1 || len(deck.Slides[0].Elements) != 2 || deck.Slides[0].Elements[1].Kind != NativeElementKindGroup || len(deck.Slides[0].Elements[1].Children) != 1 || deck.Slides[0].Elements[1].Children[0].Kind != NativeElementKindPicture {
		t.Fatalf("strict grouped picture was not projected exactly: assets=%#v elements=%#v", deck.Assets, deck.Slides[0].Elements)
	}

	refused, err := ExtractNativePPTX(fixture(t, nsOfficeRelsTransitional, ""), nativeAtomicTestExtractOptions())
	if err == nil || len(refused.Slides) != 0 || !strings.Contains(err.Error(), "mixes Strict and Transitional") {
		t.Fatalf("opposing relationship namespace was not rejected: deck=%#v err=%v", refused, err)
	}

	mixed, err := ExtractNativePPTX(fixture(t, nsOfficeRelsStrict, fmt.Sprintf(` xmlns:rt="%s" rt:link="rIdGroupImage"`, nsOfficeRelsTransitional)), nativeAtomicTestExtractOptions())
	if err == nil || len(mixed.Slides) != 0 || !strings.Contains(err.Error(), "mixes Strict and Transitional") {
		t.Fatalf("a correct relationship attribute did not mask an opposing-dialect attribute: deck=%#v err=%v", mixed, err)
	}
}

func TestExtractNativePPTXStagesNestedCapabilitiesUntilWholeGroupAcceptance(t *testing.T) {
	t.Parallel()

	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		slide := parts["relocated/slides/slide-a.xml"]
		shapeStart := strings.Index(slide, `<p:sp><p:nvSpPr><p:cNvPr id="2"`)
		if shapeStart < 0 {
			t.Fatal("fixture text shape is unavailable")
		}
		shapeEnd := strings.Index(slide[shapeStart:], `</p:sp>`)
		if shapeEnd < 0 {
			t.Fatal("fixture text shape end is unavailable")
		}
		shapeEnd += shapeStart + len(`</p:sp>`)
		textChild := strings.Replace(slide[shapeStart:shapeEnd], `id="2" name="Title"`, `id="3" name="Nested Text"`, 1)
		lateRefusal := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Late Fractional"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="101" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="3" cy="100"/></a:xfrm></p:grpSpPr>%s</p:grpSp>`, nativeGroupRectXML(5, "Late Leaf", 0, 0, 1, 1))
		group := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Atomic Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>%s%s</p:grpSp>`, textChild, lateRefusal)
		parts["relocated/slides/slide-a.xml"] = slide[:shapeStart] + group + slide[shapeEnd:]
	}})

	requests := []NativePassthroughTokenRequest{}
	issuedTokens := map[string]NativePassthroughTokenRequest{}
	options := NativePPTXExtractOptions{TokenFactory: nativeAtomicTestTokenFactory(NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		requests = append(requests, request)
		token := "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24]
		issuedTokens[token] = request
		return token, nil
	}))}
	deck, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("late-sibling group refusal: %v", err)
	}
	groupRequests := []NativePassthroughTokenRequest{}
	for _, request := range requests {
		if request.OwnerPart == "relocated/slides/slide-a.xml" && strings.HasPrefix(request.ObjectID, "cNvPr-") {
			groupRequests = append(groupRequests, request)
		}
	}
	if len(groupRequests) != 1 || groupRequests[0].ObjectID != "cNvPr-2" || groupRequests[0].Reason != "pptx.group-fractional-transform-unavailable" {
		t.Fatalf("external issuer observed an orphan nested capability: %#v", groupRequests)
	}
	if !strings.HasPrefix(string(groupRequests[0].Payload), `<p:grpSp`) || strings.Contains(string(groupRequests[0].Payload), `id="1"`) {
		t.Fatalf("final capability is not the exact object-local group: %#v", groupRequests[0])
	}
	if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 0 || deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("late refusal surfaced anything other than the opaque group: %#v", deck.Slides)
	}
	groupRefs := 0
	for _, ref := range deck.Slides[0].Passthrough {
		if ref.OwnerPart == groupRequests[0].OwnerPart && ref.FingerprintSHA256 == groupRequests[0].FingerprintSHA256 {
			groupRefs++
		}
	}
	if groupRefs != 1 {
		t.Fatalf("final group capability was not surfaced exactly once: %#v", deck.Slides[0].Passthrough)
	}
	surfacedTokens := map[string]bool{}
	for _, asset := range deck.Assets {
		for _, ref := range asset.Passthrough {
			surfacedTokens[ref.Token] = true
		}
	}
	var collectElementTokens func(NativeElement)
	collectElementTokens = func(element NativeElement) {
		for _, ref := range element.Passthrough {
			surfacedTokens[ref.Token] = true
		}
		for _, child := range element.Children {
			collectElementTokens(child)
		}
	}
	for _, slide := range deck.Slides {
		for _, ref := range slide.Passthrough {
			surfacedTokens[ref.Token] = true
		}
		for _, element := range slide.Elements {
			collectElementTokens(element)
		}
	}
	for token, request := range issuedTokens {
		if !surfacedTokens[token] {
			t.Fatalf("external issuer saw capability absent from final deck: token=%q request=%#v", token, request)
		}
	}
}

func TestExtractNativePPTXGroupTextBreakRefusalPublishesOnlyWholeGroup(t *testing.T) {
	t.Parallel()

	for _, test := range []struct {
		name   string
		strict bool
	}{
		{name: "transitional"},
		{name: "strict", strict: true},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: test.strict, mutate: func(parts map[string]string) {
				slide := parts["relocated/slides/slide-a.xml"]
				shapeStart := strings.Index(slide, `<p:sp><p:nvSpPr><p:cNvPr id="2"`)
				if shapeStart < 0 {
					t.Fatal("fixture text shape is unavailable")
				}
				shapeEnd := strings.Index(slide[shapeStart:], `</p:sp>`)
				if shapeEnd < 0 {
					t.Fatal("fixture text shape end is unavailable")
				}
				shapeEnd += shapeStart + len(`</p:sp>`)
				textChild := strings.Replace(slide[shapeStart:shapeEnd], `id="2" name="Title"`, `id="3" name="Nested Text"`, 1)
				textChild = strings.Replace(textChild, `</a:r><a:r>`, `</a:r><a:br/><a:r>`, 1)
				if !strings.Contains(textChild, `<a:br/>`) {
					t.Fatal("fixture did not expose a second text run")
				}
				group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="2" name="Text Break Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>` + textChild + `</p:grpSp>`
				parts["relocated/slides/slide-a.xml"] = slide[:shapeStart] + group + slide[shapeEnd:]
			}})

			issuer := nativeAtomicTestTokenFactory(nativeTestExtractOptions().TokenFactory)
			deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: issuer})
			if err != nil {
				t.Fatalf("grouped DrawingML break refusal: %v", err)
			}

			var groupRequests []NativePassthroughTokenRequest
			for _, request := range issuer.published {
				if request.OwnerPart == "relocated/slides/slide-a.xml" && strings.HasPrefix(request.ObjectID, "cNvPr-") {
					groupRequests = append(groupRequests, request)
				}
			}
			if len(groupRequests) != 1 || groupRequests[0].ObjectID != "cNvPr-2" || groupRequests[0].Reason != "pptx.group-child-refused-unavailable" {
				t.Fatalf("nested text capability escaped whole-group rollback: %#v", groupRequests)
			}
			if !strings.HasPrefix(string(groupRequests[0].Payload), `<p:grpSp`) || !strings.Contains(string(groupRequests[0].Payload), `<a:br/>`) {
				t.Fatalf("published capability is not the exact group subtree: %#v", groupRequests[0])
			}
			if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 0 || deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("grouped text refusal partially projected descendants: %#v", deck.Slides)
			}
			wantToken := "token-" + nativeSHA256([]byte(groupRequests[0].OwnerPart + "\x00" + groupRequests[0].ObjectID + "\x00" + groupRequests[0].Reason))[:24]
			groupRefs := 0
			for _, ref := range deck.Slides[0].Passthrough {
				if ref.Token == wantToken && ref.OwnerPart == groupRequests[0].OwnerPart && ref.FingerprintSHA256 == groupRequests[0].FingerprintSHA256 {
					groupRefs++
				}
			}
			if groupRefs != 1 {
				t.Fatalf("surface capability does not bind exactly one final issued group request: %#v", deck.Slides[0].Passthrough)
			}
		})
	}
}

func TestExtractNativePPTXGroupCapabilitiesCommitAtomically(t *testing.T) {
	t.Parallel()

	presentationNS, drawingNS := nativeGroupTestNamespaces(false)
	group := nativeNestedGroupXML(presentationNS, drawingNS, "")
	for _, idAndName := range []string{`id="5" name="Nested Rect"`, `id="6" name="Outer Rect"`} {
		group = strings.Replace(group, `<p:cNvPr `+idAndName+`/>`, `<p:cNvPr `+idAndName+`><a:extLst><a:ext uri="preserve"><x:metadata xmlns:x="urn:injoffice:test"/></a:ext></a:extLst></p:cNvPr>`, 1)
	}
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
	}})

	t.Run("legacy issuer is never called", func(t *testing.T) {
		calls := 0
		_, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(NativePassthroughTokenRequest) (string, error) {
			calls++
			return "must-not-publish", nil
		})})
		if err == nil || !strings.Contains(err.Error(), "requires an atomic passthrough token transaction") || calls != 0 {
			t.Fatalf("group reached a non-atomic issuer: calls=%d err=%v", calls, err)
		}
	})

	for _, test := range []struct {
		name       string
		issue      NativePassthroughTokenFactoryFunc
		failCommit bool
	}{
		{name: "issuer collision", issue: func(NativePassthroughTokenRequest) (string, error) { return "same-capability", nil }},
		{name: "late invalid candidate", issue: func() NativePassthroughTokenFactoryFunc {
			calls := 0
			return func(request NativePassthroughTokenRequest) (string, error) {
				calls++
				if calls == 2 {
					return "invalid capability with spaces", nil
				}
				return "candidate-" + nativeSHA256(request.Payload)[:24], nil
			}
		}()},
		{name: "commit failure", issue: func(request NativePassthroughTokenRequest) (string, error) {
			return "candidate-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
		}, failCommit: true},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			issuer := nativeAtomicTestTokenFactory(test.issue)
			issuer.failCommit = test.failCommit
			deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: issuer})
			if err == nil || len(deck.Slides) != 0 || len(issuer.published) != 0 {
				t.Fatalf("failed transaction published an orphan: deck=%#v published=%#v err=%v", deck, issuer.published, err)
			}
		})
	}
}

func TestNativeGroupStagingBudgetsAndRollbackAreBounded(t *testing.T) {
	t.Parallel()

	stager := &nativePassthroughTokenStager{requests: []nativeStagedPassthroughToken{}, payloadBytes: int64(nativeExtractMaxTotalExpandedBytes) - 1}
	request := NativePassthroughTokenRequest{ByteLength: 2, Payload: []byte("xx")}
	if _, err := stager.IssueNativePassthroughToken(request); err == nil || len(stager.requests) != 0 {
		t.Fatalf("aggregate staged payload budget was not enforced: requests=%#v err=%v", stager.requests, err)
	}
	stager.payloadBytes = 0
	for index := 0; index < nativeExtractMaxPassthrough; index++ {
		request := NativePassthroughTokenRequest{ByteLength: 0, Payload: []byte{}, ObjectID: fmt.Sprintf("object-%d", index)}
		if _, err := stager.IssueNativePassthroughToken(request); err != nil {
			t.Fatalf("stager rejected request at exact bound %d: %v", index, err)
		}
	}
	if _, err := stager.IssueNativePassthroughToken(NativePassthroughTokenRequest{ByteLength: 0, Payload: []byte{}}); err == nil {
		t.Fatal("stager accepted a request beyond its aggregate count bound")
	}
	backing := stager.requests
	stager.rollback(1, 0)
	if len(stager.requests) != 1 || stager.payloadBytes != 0 {
		t.Fatalf("bounded rollback retained staged payloads: %#v", stager)
	}
	if backing[1].request.Payload != nil || backing[1].placeholder != "" {
		t.Fatal("rollback retained rejected payload references in staging capacity")
	}
}

func TestNativeGroupRollbackKeepsPackageScopedPictureInspectionBounded(t *testing.T) {
	t.Parallel()

	payload := []byte("shared-media-payload")
	stager := &nativePassthroughTokenStager{requests: []nativeStagedPassthroughToken{}}
	extractor := nativeExtractor{
		pkg:     nativeExtractPackage{parts: map[string][]byte{"media/shared.png": payload}},
		options: NativePPTXExtractOptions{TokenFactory: stager}, sourceRevision: "revision", documentID: "document",
		passthroughCache: map[string]NativePassthroughRef{}, tokenOwners: map[string]string{},
		assets: []NativeAsset{}, assetByAlias: map[string]int{}, assetIDOwners: map[string]string{},
		picturePartInspections: map[string]nativePicturePartInspection{}, tokenStager: stager,
	}
	first := extractor.snapshotNativeGroupProjection()
	firstID, err := extractor.nativePictureAsset("media/shared.png", "image/png")
	if err != nil {
		t.Fatalf("first picture inspection: %v", err)
	}
	if extractor.mediaBytesInspected != int64(len(payload)) || len(extractor.picturePartInspections) != 1 {
		t.Fatalf("first picture inspection was not charged exactly once: bytes=%d cache=%#v", extractor.mediaBytesInspected, extractor.picturePartInspections)
	}
	extractor.rollbackNativeGroupProjection(first)
	if extractor.mediaBytesInspected != int64(len(payload)) || len(extractor.picturePartInspections) != 1 || len(extractor.assets) != 0 {
		t.Fatalf("rollback discarded work accounting/cache or retained projected output: %#v", extractor)
	}
	second := extractor.snapshotNativeGroupProjection()
	secondID, err := extractor.nativePictureAsset("media/shared.png", "image/png")
	if err != nil {
		t.Fatalf("second picture projection: %v", err)
	}
	if secondID != firstID || extractor.mediaBytesInspected != int64(len(payload)) || len(extractor.picturePartInspections) != 1 {
		t.Fatalf("shared media was re-inspected after rollback: first=%q second=%q bytes=%d cache=%#v", firstID, secondID, extractor.mediaBytesInspected, extractor.picturePartInspections)
	}
	extractor.rollbackNativeGroupProjection(second)

	extractor.mediaBytesInspected = nativeExtractMaxTotalMediaBytes - 1
	if _, err := extractor.nativePictureAsset("media/shared.png", "image/png"); err != nil {
		t.Fatalf("cached media should not consume the monotonic inspection budget again: %v", err)
	}
	extractor.pkg.parts["media/new.png"] = []byte("xx")
	if _, err := extractor.nativePictureAsset("media/new.png", "image/png"); err == nil || len(extractor.picturePartInspections) != 1 {
		t.Fatalf("uncached media bypassed the monotonic inspection budget: cache=%#v err=%v", extractor.picturePartInspections, err)
	}
}

func TestExtractNativePPTXGroupDiagnosticOverflowRollsBackWholeProjection(t *testing.T) {
	t.Parallel()

	var children strings.Builder
	for index := 0; index <= nativeMaxDiagnosticsPerScope; index++ {
		shape := nativeGroupRectXML(index+4, fmt.Sprintf("Decorative %d", index), int64(index), 0, 1, 1)
		identity := fmt.Sprintf(`<p:cNvPr id="%d" name="Decorative %d"/>`, index+4, index)
		metadata := fmt.Sprintf(`<p:cNvPr id="%d" name="Decorative %d"><a:extLst><a:ext uri="bounded"><x:metadata xmlns:x="urn:injoffice:test"/></a:ext></a:extLst></p:cNvPr>`, index+4, index)
		children.WriteString(strings.Replace(shape, identity, metadata, 1))
	}
	group := fmt.Sprintf(`<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Diagnostic Overflow"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="10"/><a:chOff x="0" y="0"/><a:chExt cx="2000" cy="10"/></a:xfrm></p:grpSpPr>%s</p:grpSp>`, children.String())
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
	}})
	issuer := nativeAtomicTestTokenFactory(nativeTestExtractOptions().TokenFactory)
	deck, err := ExtractNativePPTX(payload, NativePPTXExtractOptions{TokenFactory: issuer})
	if err != nil {
		t.Fatalf("diagnostic-overflow group: %v", err)
	}
	if len(deck.Slides[0].Elements) != 1 || !nativeDiagnosticsContain(deck.Slides[0].Compatibility.Diagnostics, "pptx.group-diagnostic-budget-unavailable") {
		t.Fatalf("overflowing group was partially surfaced: %#v", deck.Slides[0])
	}
	for _, request := range issuer.published {
		if request.Reason == "pptx.autoshape-nonvisual-unavailable" {
			t.Fatalf("rolled-back descendant capability was published: %#v", request)
		}
	}
}

func TestExtractNativePPTXRejectsMixedGroupDialectWithoutDeck(t *testing.T) {
	t.Parallel()

	presentationNS, drawingNS := nativeGroupTestNamespaces(true)
	group := nativeNestedGroupXML(presentationNS, drawingNS, "")
	group = strings.Replace(group, `<a:off x="1000000"`, fmt.Sprintf(`<at:off xmlns:at="%s" x="1000000"`, nsDrawingTransitional), 1)
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: true, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, group+`</p:spTree>`, 1)
	}})
	deck, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err == nil || !strings.Contains(err.Error(), "mixes Strict and Transitional") || len(deck.Slides) != 0 {
		t.Fatalf("mixed group dialect was not rejected atomically: deck=%#v err=%v", deck, err)
	}
}

func assertNativeGroupElement(t *testing.T, element NativeElement, name, objectID string, transform, childTransform NativeTransform) {
	t.Helper()
	if element.Kind != NativeElementKindGroup || element.Name == nil || *element.Name != name || element.Source == nil || element.Source.ObjectID != objectID {
		t.Fatalf("unexpected group identity: %#v", element)
	}
	if !nativeTransformsEqual(element.Transform, transform) {
		t.Fatalf("group %s transform: got %#v want %#v", name, element.Transform, transform)
	}
	if element.ChildTransform == nil || !nativeTransformsEqual(*element.ChildTransform, childTransform) {
		t.Fatalf("group %s child transform: got %#v want %#v", name, element.ChildTransform, childTransform)
	}
}

func assertNativeProjectedTransform(t *testing.T, element NativeElement, objectID string, transform NativeTransform) {
	t.Helper()
	if element.Source == nil || element.Source.ObjectID != objectID || !nativeTransformsEqual(element.Transform, transform) {
		t.Fatalf("projected child %s: got %#v want %#v", objectID, element, transform)
	}
}

func nativeTransformsEqual(left, right NativeTransform) bool {
	if left.X == nil || left.Y == nil || left.Cx == nil || left.Cy == nil || right.X == nil || right.Y == nil || right.Cx == nil || right.Cy == nil {
		return false
	}
	return *left.X == *right.X && *left.Y == *right.Y && *left.Cx == *right.Cx && *left.Cy == *right.Cy
}

func nativeDiagnosticsContain(diagnostics []NativeDiagnostic, code string) bool {
	for _, diagnostic := range diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

func nativeGroupTestNamespaces(strict bool) (string, string) {
	if strict {
		return nsPresentationStrict, nsDrawingStrict
	}
	return nsPresentationTransitional, nsDrawingTransitional
}

func nativeNestedGroupXML(presentationNS, drawingNS, outerTransformAttrs string) string {
	return fmt.Sprintf(`<p:grpSp xmlns:p="%s" xmlns:a="%s"><p:nvGrpSpPr><p:cNvPr id="3" name="Outer"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm%s><a:off x="1000000" y="2000000"/><a:ext cx="6000000" cy="8000000"/><a:chOff x="100" y="200"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr><p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Inner"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="150" y="250"/><a:ext cx="100" cy="100"/><a:chOff x="10" y="20"/><a:chExt cx="50" cy="50"/></a:xfrm></p:grpSpPr>%s</p:grpSp>%s</p:grpSp>`,
		presentationNS, drawingNS, outerTransformAttrs,
		nativeGroupRectXML(5, "Nested Rect", 20, 30, 10, 5),
		nativeGroupRectXML(6, "Outer Rect", 250, 300, 25, 20))
}

func nativeGroupRectXML(id int, name string, x, y, cx, cy int64) string {
	return fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%d" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="0" cap="flat" cmpd="sng" algn="ctr"><a:noFill/><a:prstDash val="solid"/><a:round/></a:ln></p:spPr></p:sp>`, id, name, x, y, cx, cy)
}

func nativeDeepGroupXML(presentationNS, drawingNS string, depth int) string {
	var value strings.Builder
	for index := 0; index < depth; index++ {
		id := index + 3
		fmt.Fprintf(&value, `<p:grpSp xmlns:p="%s" xmlns:a="%s"><p:nvGrpSpPr><p:cNvPr id="%d" name="Depth %d"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>`, presentationNS, drawingNS, id, index+1)
	}
	value.WriteString(nativeGroupRectXML(depth+3, "Leaf", 0, 0, 10, 10))
	for index := 0; index < depth; index++ {
		value.WriteString(`</p:grpSp>`)
	}
	return value.String()
}

type nativeAtomicTestTokenIssuer struct {
	issue      NativePassthroughTokenFactory
	published  []NativePassthroughTokenRequest
	failCommit bool
}

func nativeAtomicTestTokenFactory(issue NativePassthroughTokenFactory) *nativeAtomicTestTokenIssuer {
	return &nativeAtomicTestTokenIssuer{issue: issue, published: []NativePassthroughTokenRequest{}}
}

func nativeAtomicTestExtractOptions() NativePPTXExtractOptions {
	options := nativeTestExtractOptions()
	options.TokenFactory = nativeAtomicTestTokenFactory(options.TokenFactory)
	return options
}

func (issuer *nativeAtomicTestTokenIssuer) IssueNativePassthroughToken(NativePassthroughTokenRequest) (string, error) {
	return "", fmt.Errorf("direct non-transactional issuance is unavailable")
}

func (issuer *nativeAtomicTestTokenIssuer) BeginNativePassthroughTokenTransaction() (NativePassthroughTokenTransaction, error) {
	return &nativeAtomicTestTokenTransaction{issuer: issuer, staged: []NativePassthroughTokenRequest{}}, nil
}

type nativeAtomicTestTokenTransaction struct {
	issuer     *nativeAtomicTestTokenIssuer
	staged     []NativePassthroughTokenRequest
	rolledBack bool
}

func (transaction *nativeAtomicTestTokenTransaction) IssueNativePassthroughToken(request NativePassthroughTokenRequest) (string, error) {
	token, err := transaction.issuer.issue.IssueNativePassthroughToken(request)
	if err == nil {
		transaction.staged = append(transaction.staged, request)
	}
	return token, err
}

func (transaction *nativeAtomicTestTokenTransaction) CommitNativePassthroughTokens() error {
	if transaction.issuer.failCommit {
		return fmt.Errorf("injected atomic commit failure")
	}
	transaction.issuer.published = append(transaction.issuer.published, transaction.staged...)
	return nil
}

func (transaction *nativeAtomicTestTokenTransaction) RollbackNativePassthroughTokens() {
	transaction.rolledBack = true
	transaction.staged = nil
}
