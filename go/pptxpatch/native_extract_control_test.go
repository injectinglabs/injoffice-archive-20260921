package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

// controlSnapshotWMF builds a metafile shaped like the control snapshots
// PowerPoint writes: a solid face and a caption the decoder cannot paint.
func controlSnapshotWMF(withCaption bool) string {
	builder := &metafileBuilder{}
	builder.mapMode(nativeMetafileMapModeAnisotropic).windowOrg(0, 0).windowExt(12, 6)
	builder.brush(0x10, 0x20, 0x30).selectObject(0).patBlt(0, 0, 12, 6, nativeMetafileROPPatCopy)
	if withCaption {
		builder.record(nativeMetafileRecordExtTextOut, append(words(1, 1, 2, 0), []byte("Ok")...))
	}
	return string(builder.build())
}

// nativeControlFixtureOptions states the p:controls markup a fixture carries.
type nativeControlFixtureOptions struct {
	// controls replaces the whole p:controls body.
	controls string
	// mediaByPart is the bytes of each media part the controls reference.
	mediaByPart map[string]string
	// mimeByPart is each media part's declared content type.
	mimeByPart map[string]string
	// spTreePicture adds a p:pic to the shape tree claiming rIdControl1, so a
	// test can prove one relationship is never claimed twice.
	spTreePicture bool
}

func nativeControlFixture(t *testing.T, options nativeControlFixtureOptions) []byte {
	t.Helper()
	extraParts := []nativeExtractZipPart{}
	for part, data := range options.mediaByPart {
		extraParts = append(extraParts, nativeExtractZipPart{name: part, data: data})
	}
	return nativeExtractFixture(t, nativeExtractFixtureOptions{
		extraParts: extraParts,
		mutate: func(parts map[string]string) {
			overrides, relationships := "", ""
			index := 0
			for part := range options.mediaByPart {
				index++
				mime := options.mimeByPart[part]
				if mime == "" {
					mime = "image/x-wmf"
				}
				overrides += fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/>`, part, mime)
				relationships += fmt.Sprintf(`<Relationship Id="rIdControl%d" Type="%s" Target="/%s"/>`, index, relImageTransitional, part)
			}
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, overrides+`</Types>`, 1)
			parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`, relationships+`</Relationships>`, 1)
			if options.spTreePicture {
				picture := fmt.Sprintf(`<p:pic><p:nvPicPr><p:cNvPr id="9" name="Tree"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="%s" r:embed="rIdControl1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`, nsOfficeRelsTransitional)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, picture+`</p:spTree>`, 1)
			}
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:cSld>`, `<p:controls>`+options.controls+`</p:controls></p:cSld>`, 1)
		},
	})
}

// controlFallback is the shape PowerPoint writes: an mc:AlternateContent whose
// mc:Choice needs a namespace this reader does not implement and whose
// mc:Fallback carries an ordinary p:pic of the control.
func controlFallback(id int, relationshipID string, x, y int64) string {
	return fmt.Sprintf(`<mc:AlternateContent xmlns:mc="%s"><mc:Choice Requires="v"><p:control spid="_x0000_s10%d"/></mc:Choice><mc:Fallback><p:control name="CommandButton%d"><p:pic><p:nvPicPr><p:cNvPr id="%d" name="CommandButton%d"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip xmlns:r="%s" r:embed="%s"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:control></mc:Fallback></mc:AlternateContent>`,
		nsNativeMarkupCompatibility, id, id, 20+id, id, nsOfficeRelsTransitional, relationshipID, x, y)
}

// A slide whose every visible object is an ActiveX control has an empty
// p:spTree; the pictures live in p:cSld/p:controls. Painting the fallback each
// control states is what MCE requires of a consumer that does not implement
// the branch mc:Choice asks for.
func TestExtractNativePPTXPaintsControlFallbackPictures(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls:    controlFallback(1, "rIdControl1", 100, 200) + controlFallback(2, "rIdControl2", 300, 400),
		mediaByPart: map[string]string{"relocated/media/control1.wmf": controlSnapshotWMF(false), "relocated/media/control2.wmf": controlSnapshotWMF(true)},
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	pictures := 0
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindPicture {
			pictures++
		}
	}
	if pictures != 2 {
		t.Fatalf("slide painted %d control pictures, want 2: %#v", pictures, deck.Slides[0].Elements)
	}
	if len(deck.Assets) != 2 {
		t.Fatalf("expected one asset per control snapshot: %#v", deck.Assets)
	}
	for _, asset := range deck.Assets {
		// The asset is the raster, not the metafile: the contract carries no
		// metafile, so what the preview serves is the PNG derived from it.
		if asset.ContentType != "image/png" {
			t.Fatalf("control asset content type is %q, want image/png", asset.ContentType)
		}
		if asset.SourceTransform == nil || *asset.SourceTransform != NativeAssetSourceTransformWmfRasterV1 {
			t.Fatalf("control asset does not name its derivation: %#v", asset.SourceTransform)
		}
		if asset.SourceByteLength == nil || asset.Source == nil {
			t.Fatalf("control asset does not pin the part it derives from: %#v", asset)
		}
		if asset.SHA256 == asset.Source.FingerprintSHA256 {
			t.Fatal("derived asset digest equals its source digest, so the derivation was not applied")
		}
		if *asset.ByteLength == *asset.SourceByteLength {
			t.Fatal("derived asset length equals its source length, so the derivation was not applied")
		}
		// The capability is bound to the bytes the host reads, which is the
		// part, not the raster.
		if len(asset.Passthrough) != 1 || asset.Passthrough[0].FingerprintSHA256 != asset.Source.FingerprintSHA256 {
			t.Fatalf("control asset capability is not bound to its source part: %#v", asset.Passthrough)
		}
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid native deck: %#v", issues)
	}
}

// Only the picture is modeled. The control markup itself is still outside
// native v1 and has to stay preserved, so the slide says so.
func TestExtractNativePPTXPreservesTheControlMarkupItself(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls:    controlFallback(1, "rIdControl1", 100, 200),
		mediaByPart: map[string]string{"relocated/media/control1.wmf": controlSnapshotWMF(false)},
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	if !nativeSlideHasDiagnostic(deck.Slides[0], "pptx.unsupported-embedded-control") {
		t.Fatalf("slide does not disclose that the control markup is preserved: %#v", deck.Slides[0].Compatibility.Diagnostics)
	}
}

// A metafile that draws a caption is painted without it, because glyphs need
// font bytes this reader has no access to. That has to be stated on the
// element, not left for the viewer to notice.
func TestExtractNativePPTXDisclosesMetafileTextItCannotPaint(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls:    controlFallback(1, "rIdControl1", 100, 200),
		mediaByPart: map[string]string{"relocated/media/control1.wmf": controlSnapshotWMF(true)},
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	if !nativeSlideHasDiagnostic(deck.Slides[0], nativePictureMetafileTextCode) {
		t.Fatalf("slide does not disclose the metafile text it cannot paint: %#v", deck.Slides[0].Compatibility.Diagnostics)
	}
	// The picture still paints; the caption is a gap in it, not a reason to
	// throw the whole snapshot away.
	if nativeSlideHasDiagnostic(deck.Slides[0], nativePictureMetafileCode) {
		t.Fatal("a paintable snapshot was reported as unpaintable")
	}
}

// A metafile outside the modeled subset is not approximated. The picture stays
// exactly where an unreadable image already was — preserved, unpainted — and
// now names the construct that stopped it.
func TestExtractNativePPTXRefusesUnmodeledMetafileByName(t *testing.T) {
	t.Parallel()

	polygon := &metafileBuilder{}
	polygon.mapMode(nativeMetafileMapModeAnisotropic).windowExt(12, 6)
	polygon.record(0x0324, words(3, 0, 0, 4, 0, 2, 4))

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls:    controlFallback(1, "rIdControl1", 100, 200),
		mediaByPart: map[string]string{"relocated/media/control1.wmf": string(polygon.build())},
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	if !nativeSlideHasDiagnostic(deck.Slides[0], nativePictureMetafileCode) {
		t.Fatalf("slide does not disclose the refused metafile: %#v", deck.Slides[0].Compatibility.Diagnostics)
	}
	if len(deck.Assets) != 1 || deck.Assets[0].SourceTransform != nil || deck.Assets[0].ContentType != "image/x-wmf" {
		t.Fatalf("a refused metafile was still claimed as a derived raster: %#v", deck.Assets)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid native deck: %#v", issues)
	}
}

// Two elements may not claim one image relationship. A picture already in the
// shape tree keeps it, and the control that names it again contributes nothing.
func TestExtractNativePPTXDoesNotClaimOneImageRelationshipTwice(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls:      controlFallback(1, "rIdControl1", 100, 200),
		mediaByPart:   map[string]string{"relocated/media/control1.wmf": controlSnapshotWMF(false)},
		spTreePicture: true,
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	pictures := 0
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindPicture {
			pictures++
		}
	}
	if pictures != 1 {
		t.Fatalf("one image relationship produced %d pictures, want 1", pictures)
	}
}

// A p:control that names only its own relationship states no fallback picture,
// so there is nothing to paint and nothing to claim.
func TestExtractNativePPTXIgnoresControlsWithNoFallbackPicture(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls: `<p:control name="Bare" spid="_x0000_s1026"/>`,
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindPicture {
			t.Fatalf("a control with no fallback picture produced one: %#v", element)
		}
	}
	if !nativeSlideHasDiagnostic(deck.Slides[0], "pptx.unsupported-embedded-control") {
		t.Fatal("a control with no fallback picture is still preserved markup and must say so")
	}
}

func nativeSlideHasDiagnostic(slide NativeSlide, code string) bool {
	for _, diagnostic := range slide.Compatibility.Diagnostics {
		if diagnostic.Code == code {
			return true
		}
	}
	return false
}

// The contract only lets a host run a derivation when the deck pins both ends
// of it, so the two fields that describe one have to travel together.
func TestValidateNativePPTXBindsAssetSourceTransformToItsSource(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativeControlFixture(t, nativeControlFixtureOptions{
		controls:    controlFallback(1, "rIdControl1", 100, 200),
		mediaByPart: map[string]string{"relocated/media/control1.wmf": controlSnapshotWMF(false)},
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract controls: %v", err)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("the extractor emitted a deck its own validator rejects: %#v", issues)
	}
	for _, test := range []struct {
		name   string
		mutate func(*NativeAsset)
	}{
		{"length without a transform", func(a *NativeAsset) { a.SourceTransform = nil }},
		{"transform without a length", func(a *NativeAsset) { a.SourceByteLength = nil }},
		{"transform without a source", func(a *NativeAsset) { a.Source = nil }},
		{"transform that cannot produce this type", func(a *NativeAsset) { a.ContentType = "image/jpeg" }},
		{"unknown transform", func(a *NativeAsset) {
			unknown := NativeAssetSourceTransform("somethingElseV1")
			a.SourceTransform = &unknown
		}},
		{"length beyond the asset bound", func(a *NativeAsset) {
			oversized := int64(nativeMaxAssetBytes) + 1
			a.SourceByteLength = &oversized
		}},
	} {
		t.Run(test.name, func(t *testing.T) {
			bad := deck
			bad.Assets = append([]NativeAsset(nil), deck.Assets...)
			test.mutate(&bad.Assets[0])
			if issues := ValidateNativePPTX(bad); len(issues) == 0 {
				t.Fatal("validator accepted a derivation the host could not safely run")
			}
		})
	}
}
