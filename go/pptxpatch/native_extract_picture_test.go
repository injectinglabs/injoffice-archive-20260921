package pptxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXPicturesTransitionalAndStrict(t *testing.T) {
	t.Parallel()

	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			data := "\x89PNG\r\n\x1a\nfixture"
			deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{strict: strict, imageData: data}), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract picture: %v", err)
			}
			picture := nativeFixturePicture(t, deck.Slides[0])
			if picture.Name == nil || *picture.Name != "Photo 1" || picture.AssetID == nil || picture.Source == nil || picture.Source.RelationshipID == nil || *picture.Source.RelationshipID != "rIdImage" {
				t.Fatalf("picture identity/relationship was not extracted: %#v", picture)
			}
			if picture.Transform.X == nil || *picture.Transform.X != 1_828_800 || picture.Transform.Y == nil || *picture.Transform.Y != 914_400 || picture.Transform.Cx == nil || *picture.Transform.Cx != 3_657_600 || picture.Transform.Cy == nil || *picture.Transform.Cy != 2_743_200 {
				t.Fatalf("picture transform mismatch: %#v", picture.Transform)
			}
			if picture.Compatibility.Status != NativeCompatibilityStatusEditable || len(picture.Passthrough) != 0 {
				t.Fatalf("exact picture subset was not editable: %#v", picture.Compatibility)
			}
			if len(deck.Assets) != 1 {
				t.Fatalf("expected one asset: %#v", deck.Assets)
			}
			asset := deck.Assets[0]
			if asset.ID != *picture.AssetID || asset.ContentType != "image/png" || asset.SHA256 != nativeSHA256([]byte(data)) || asset.ByteLength == nil || *asset.ByteLength != int64(len(data)) || asset.DataBase64 != nil || asset.Source == nil || asset.Source.PartName != "relocated/media/image.png" || asset.Source.FingerprintSHA256 != asset.SHA256 || len(asset.Passthrough) != 1 || asset.Passthrough[0].OwnerPart != asset.Source.PartName || asset.Passthrough[0].FingerprintSHA256 != asset.SHA256 {
				t.Fatalf("asset metadata/source mismatch: %#v", asset)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid native deck: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXPictureStretchDefaults(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, content := range []string{"", " \n\t", "<a:fillRect/>"} {
			t.Run(fmt.Sprintf("strict=%v/content=%q", strict, content), func(t *testing.T) {
				deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{strict: strict, stretchContent: &content}), nativeTestExtractOptions())
				if err != nil {
					t.Fatal(err)
				}
				picture := nativeFixturePicture(t, deck.Slides[0])
				if picture.Compatibility.Status != NativeCompatibilityStatusEditable {
					t.Fatalf("default stretch not modeled: %#v", picture.Compatibility)
				}
				if issues := ValidateNativePPTX(deck); len(issues) != 0 {
					t.Fatalf("invalid deck: %#v", issues)
				}
			})
		}
	}
}

func TestExtractNativePPTXPictureStretchRetainsSafetyBoundaries(t *testing.T) {
	for _, content := range []string{"rogue", `<a:fillRect l="1000"/>`, `<a:unknown/>`} {
		t.Run(content, func(t *testing.T) {
			deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{stretchContent: &content}), nativeTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			if picture := nativeFixturePicture(t, deck.Slides[0]); picture.Compatibility.Status == NativeCompatibilityStatusEditable {
				t.Fatal("unmodeled stretch became editable")
			}
		})
	}
	duplicate := `<a:fillRect/><a:fillRect/>`
	if _, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{stretchContent: &duplicate}), nativeTestExtractOptions()); err == nil {
		t.Fatal("duplicate rectangle accepted")
	}
}

func TestExtractNativePPTXPictureRoutesRelocatedCaseAndPercentAsset(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{
		imagePart:   "RELOCATED/MEDIA/%49MAGE.PNG",
		imageTarget: "../MEDIA/IMAGE.PNG",
		imageMIME:   "IMAGE/PNG",
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract aliased picture: %v", err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	if len(deck.Assets) != 1 || deck.Assets[0].Source == nil || deck.Assets[0].Source.PartName != "RELOCATED/MEDIA/%49MAGE.PNG" || picture.AssetID == nil || *picture.AssetID != deck.Assets[0].ID {
		t.Fatalf("actual asset spelling/alias identity was not preserved: picture=%#v assets=%#v", picture, deck.Assets)
	}
}

func TestExtractNativePPTXPictureDeduplicatesSharedExactPart(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{pictureCount: 2}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract shared picture: %v", err)
	}
	var pictures []NativeElement
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindPicture {
			pictures = append(pictures, element)
		}
	}
	if len(pictures) != 2 || len(deck.Assets) != 1 || pictures[0].AssetID == nil || pictures[1].AssetID == nil || *pictures[0].AssetID != deck.Assets[0].ID || *pictures[1].AssetID != deck.Assets[0].ID || pictures[0].ID == pictures[1].ID {
		t.Fatalf("shared part was not deduplicated with distinct element identities: pictures=%#v assets=%#v", pictures, deck.Assets)
	}
}

func TestExtractNativePPTXPictureDoesNotCollapseDistinctPartsWithSameBytes(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{
		pictureCount:      2,
		secondImagePart:   "relocated/media/image-copy.png",
		secondImageTarget: "../media/image-copy.png",
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract distinct same-byte picture parts: %v", err)
	}
	if len(deck.Assets) != 2 || deck.Assets[0].ID == deck.Assets[1].ID || deck.Assets[0].Source.PartName == deck.Assets[1].Source.PartName || deck.Assets[0].SHA256 != deck.Assets[1].SHA256 {
		t.Fatalf("distinct source parts collapsed or collided: %#v", deck.Assets)
	}
}

func TestExtractNativePPTXPictureSourceOnlyAssetRequiresBoundCapability(t *testing.T) {
	t.Parallel()

	var assetRequest *NativePassthroughTokenRequest
	options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		if request.Reason == "pptx.picture-asset-source" {
			copy := request
			assetRequest = &copy
		}
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	})}
	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{pictureCount: 2}), options)
	if err != nil {
		t.Fatalf("extract picture asset capability: %v", err)
	}
	if len(deck.Assets) != 1 || deck.Assets[0].DataBase64 != nil || len(deck.Assets[0].Passthrough) != 1 || assetRequest == nil {
		t.Fatalf("source-only asset was fetch-authorized by metadata alone: asset=%#v request=%#v", deck.Assets, assetRequest)
	}
	asset := deck.Assets[0]
	if assetRequest.SourceRevision != *deck.SourceRevision || assetRequest.OwnerPart != asset.Source.PartName || assetRequest.ObjectID != asset.Source.ObjectID || assetRequest.FingerprintSHA256 != asset.SHA256 || assetRequest.ByteLength != *asset.ByteLength || int64(len(assetRequest.Payload)) != *asset.ByteLength || assetRequest.Reason != "pptx.picture-asset-source" {
		t.Fatalf("asset capability request is not fully bound: asset=%#v request=%#v", asset, assetRequest)
	}
}

func TestExtractNativePPTXPictureCropRotationAndEffectsArePreserveOnly(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{
		xfrmAttrs:   ` rot="60000" flipH="true"`,
		sourceRect:  `<a:srcRect l="-10000" t="0" r="0" b="0"/>`,
		blipContent: `<a:alphaBiLevel thresh="50000"/>`,
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract fidelity-limited picture: %v", err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	if picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(picture.Passthrough) != 1 {
		t.Fatalf("picture semantics were silently approximated: %#v", picture)
	}
	codes := map[string]bool{}
	for _, diagnostic := range picture.Compatibility.Diagnostics {
		codes[diagnostic.Code] = true
	}
	for _, code := range []string{"pptx.picture-crop-unavailable", "pptx.picture-transform-unavailable", "pptx.picture-effects-unavailable"} {
		if !codes[code] {
			t.Fatalf("missing picture fidelity diagnostic %q: %#v", code, picture.Compatibility.Diagnostics)
		}
	}
}

func TestExtractNativePPTXPicturePositiveCropRetainsExactSource(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		data := "\x89PNG\r\n\x1a\nexact-cropped-image"
		input := nativePictureFixture(t, nativePictureFixtureOptions{strict: strict, imageData: data, sourceRect: `<a:srcRect l="12500" t="25000" r="37500"/>`})
		before := nativeSHA256(input)
		deck, err := ExtractNativePPTX(input, nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		picture := nativeFixturePicture(t, deck.Slides[0])
		if picture.Compatibility.Status != NativeCompatibilityStatusEditable || len(picture.Passthrough) != 0 {
			t.Fatalf("qualified crop became opaque: %#v", picture)
		}
		crop := picture.Crop
		if crop == nil || *crop.Left != 12500 || *crop.Top != 25000 || *crop.Right != 37500 || *crop.Bottom != 0 {
			t.Fatalf("source crop not retained exactly: %#v", crop)
		}
		if deck.Assets[0].SHA256 != nativeSHA256([]byte(data)) || *deck.Assets[0].ByteLength != int64(len(data)) || nativeSHA256(input) != before {
			t.Fatal("crop changed source package or image bytes")
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatalf("invalid crop contract: %#v", issues)
		}
	}
}

func TestExtractNativePPTXPictureOutsetAndDegenerateCropsStayOpaque(t *testing.T) {
	t.Parallel()
	for _, rectangle := range []string{`l="-1"`, `t="-1"`, `r="-1"`, `b="-1"`, `l="50000" r="50000"`, `t="99999" b="1"`, `l="100000"`} {
		deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{sourceRect: `<a:srcRect ` + rectangle + `/>`}), nativeTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		picture := nativeFixturePicture(t, deck.Slides[0])
		if picture.Crop != nil || picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
			t.Fatalf("unsupported crop %s was approximated: %#v", rectangle, picture)
		}
	}
}

func TestNativePPTXPictureCropContractRejectsMalformedInsets(t *testing.T) {
	t.Parallel()
	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{sourceRect: `<a:srcRect l="12500"/>`}), nativeTestExtractOptions())
	if err != nil {
		t.Fatal(err)
	}
	payload, err := MarshalNativePPTXJSON(deck)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := DecodeNativePPTXJSON(payload); err != nil {
		t.Fatal(err)
	}
	for _, replacement := range []string{`"left":-1`, `"left":100000`, `"left":0.5`, `"left":50000,"right":50000`, `"unknown":12500`, `"left":null`} {
		malformed := strings.Replace(string(payload), `"left":12500`, replacement, 1)
		if malformed == string(payload) {
			t.Fatal("crop test did not mutate payload")
		}
		if _, err := DecodeNativePPTXJSON([]byte(malformed)); err == nil {
			t.Fatalf("accepted malformed crop: %s", replacement)
		}
	}
	for _, edges := range [][4]int64{{-1, 0, 0, 0}, {100000, 0, 0, 0}, {50000, 0, 50000, 0}, {0, 99999, 0, 1}} {
		picture := &deck.Slides[0].Elements[0]
		picture.Crop = &NativePictureCrop{Left: &edges[0], Top: &edges[1], Right: &edges[2], Bottom: &edges[3]}
		if len(ValidateNativePPTX(deck)) == 0 {
			t.Fatalf("accepted invalid opposing insets: %v", edges)
		}
	}
}

func TestExtractNativePPTXPictureExplicitNoOpCropRotationAndFlipsStayExact(t *testing.T) {
	t.Parallel()

	baseline, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("baseline picture: %v", err)
	}
	explicit, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{
		xfrmAttrs:  ` rot="0" flipH="false" flipV="0"`,
		sourceRect: `<a:srcRect l="0" t="0" r="0" b="0"/>`,
	}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("explicit no-op picture metadata: %v", err)
	}
	baselinePicture := nativeFixturePicture(t, baseline.Slides[0])
	explicitPicture := nativeFixturePicture(t, explicit.Slides[0])
	if explicitPicture.Compatibility.Status != NativeCompatibilityStatusEditable ||
		*explicitPicture.Transform.X != *baselinePicture.Transform.X || *explicitPicture.Transform.Y != *baselinePicture.Transform.Y ||
		*explicitPicture.Transform.Cx != *baselinePicture.Transform.Cx || *explicitPicture.Transform.Cy != *baselinePicture.Transform.Cy {
		t.Fatalf("explicit no-op crop/rotation/flips changed rendering semantics: baseline=%#v explicit=%#v", baselinePicture, explicitPicture)
	}
}

func TestExtractNativePPTXPictureLinkedRelationshipIsExactAndExternal(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{linkRelationshipID: "rIdLink"}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract embedded picture with external fallback link: %v", err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	if picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
		t.Fatalf("linked picture semantics were not conservative: %#v", picture)
	}
	for _, test := range []struct {
		name    string
		options nativePictureFixtureOptions
	}{
		{name: "dangling link", options: nativePictureFixtureOptions{linkRelationshipID: "rIdLink", omitLinkRelationship: true}},
		{name: "wrong link type", options: nativePictureFixtureOptions{linkRelationshipID: "rIdLink", linkRelationshipType: relChartTransitional}},
		{name: "internal link", options: nativePictureFixtureOptions{linkRelationshipID: "rIdLink", linkRelationshipMode: "Internal"}},
	} {
		t.Run(test.name, func(t *testing.T) {
			if _, err := ExtractNativePPTX(nativePictureFixture(t, test.options), nativeTestExtractOptions()); err == nil {
				t.Fatal("invalid linked-picture relationship was accepted")
			}
		})
	}
}

func TestExtractNativePPTXPictureOpaqueNonvisualMetadataIsPreserveOnly(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{nonVisualContent: `<p:ph type="pic"/>`}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract inherited picture metadata: %v", err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	if picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(picture.Passthrough) != 1 {
		t.Fatalf("inherited/opaque picture metadata was not preserved explicitly: %#v", picture)
	}
}

func TestExtractNativePPTXPictureRejectsDuplicateModeledSingleton(t *testing.T) {
	t.Parallel()

	if _, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{duplicateBlip: true}), nativeTestExtractOptions()); err == nil {
		t.Fatal("duplicate modeled picture blip was accepted")
	}
}

func TestExtractNativePPTXPictureUnsupportedMIMEIsExplicit(t *testing.T) {
	t.Parallel()

	deck, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{imageMIME: "image/tiff"}), nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("extract preserve-only TIFF: %v", err)
	}
	picture := nativeFixturePicture(t, deck.Slides[0])
	if picture.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || len(deck.Assets) != 1 || deck.Assets[0].ContentType != "image/tiff" {
		t.Fatalf("unsupported image MIME was not explicit: picture=%#v assets=%#v", picture, deck.Assets)
	}
	found := false
	for _, diagnostic := range picture.Compatibility.Diagnostics {
		found = found || diagnostic.Code == "pptx.picture-content-type-unavailable"
	}
	if !found {
		t.Fatalf("missing unsupported MIME diagnostic: %#v", picture.Compatibility.Diagnostics)
	}
}

func TestExtractNativePPTXPictureHostileRelationshipAndMIME(t *testing.T) {
	t.Parallel()

	tests := []struct {
		name    string
		options nativePictureFixtureOptions
	}{
		{name: "missing target", options: nativePictureFixtureOptions{omitImagePart: true}},
		{name: "wrong relationship type", options: nativePictureFixtureOptions{relationshipType: relChartTransitional}},
		{name: "external embed", options: nativePictureFixtureOptions{relationshipMode: "External", imageTarget: "https://example.invalid/image.png", omitImagePart: true}},
		{name: "non-image MIME", options: nativePictureFixtureOptions{imageMIME: "application/octet-stream"}},
		{name: "duplicate relationship id", options: nativePictureFixtureOptions{duplicateRelationshipID: true}},
	}
	for _, test := range tests {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			if _, err := ExtractNativePPTX(nativePictureFixture(t, test.options), nativeTestExtractOptions()); err == nil {
				t.Fatal("hostile relationship/MIME was accepted")
			}
		})
	}
}

func TestExtractNativePPTXPictureIdentityAndAssetReuseAcrossSlideReorder(t *testing.T) {
	t.Parallel()

	payload := nativePictureFixture(t, nativePictureFixtureOptions{secondSlide: true})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial picture extract: %v", err)
	}
	oldAssetID := previous.Assets[0].ID
	previous.Assets[0].ID = "asset-reused"
	pictureIDs := map[string]string{}
	for slideIndex := range previous.Slides {
		for elementIndex := range previous.Slides[slideIndex].Elements {
			element := &previous.Slides[slideIndex].Elements[elementIndex]
			if element.Kind != NativeElementKindPicture {
				continue
			}
			element.ID = "picture-reused-" + fmt.Sprint(slideIndex)
			*element.AssetID = "asset-reused"
			pictureIDs[element.Source.PartName] = element.ID
		}
	}
	if oldAssetID == previous.Assets[0].ID || len(pictureIDs) != 2 || len(ValidateNativePPTX(previous)) != 0 {
		t.Fatalf("test Previous is not valid after durable-id rewrite: %#v", ValidateNativePPTX(previous))
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	reordered, err := ExtractNativePPTX(nativePictureFixture(t, nativePictureFixtureOptions{secondSlide: true, reverseSlides: true}), options)
	if err != nil {
		t.Fatalf("reordered picture extract: %v", err)
	}
	if len(reordered.Assets) != 1 || reordered.Assets[0].ID != "asset-reused" || reordered.Slides[0].Source.PartName != "relocated/slides/slide-b.xml" {
		t.Fatalf("asset/slide reorder identity mismatch: %#v", reordered)
	}
	for _, slide := range reordered.Slides {
		picture := nativeFixturePicture(t, slide)
		if picture.ID != pictureIDs[slide.Source.PartName] || picture.AssetID == nil || *picture.AssetID != "asset-reused" {
			t.Fatalf("picture identity changed for %q: %#v", slide.Source.PartName, picture)
		}
	}
}

func TestExtractNativePPTXFiltersPreviousAssetIdentityToMatchedSlideRelationships(t *testing.T) {
	t.Parallel()

	fixtureOptions := nativePictureFixtureOptions{
		secondSlide:            true,
		secondSlideImagePart:   "relocated/media/second-slide.png",
		secondSlideImageTarget: "../media/second-slide.png",
	}
	payload := nativePictureFixture(t, fixtureOptions)
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial picture extract: %v", err)
	}
	if len(previous.Assets) != 2 {
		t.Fatalf("fixture did not create distinct per-slide assets: %#v", previous.Assets)
	}
	previous.Slides[1].Source.ObjectID = "sldId-999"
	secondPicture := nativeFixturePicture(t, previous.Slides[1])
	var poisonedAsset *NativeAsset
	for index := range previous.Assets {
		if previous.Assets[index].Source != nil && previous.Assets[index].Source.PartName == "relocated/media/second-slide.png" {
			poisonedAsset = &previous.Assets[index]
			break
		}
	}
	if poisonedAsset == nil || secondPicture.AssetID == nil || *secondPicture.AssetID != poisonedAsset.ID {
		t.Fatalf("second-slide asset linkage missing: picture=%#v assets=%#v", secondPicture, previous.Assets)
	}
	poisonedAsset.ID = "asset-poisoned"
	for elementIndex := range previous.Slides[1].Elements {
		if previous.Slides[1].Elements[elementIndex].Kind == NativeElementKindPicture {
			*previous.Slides[1].Elements[elementIndex].AssetID = poisonedAsset.ID
		}
	}
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("partial-anchor asset Previous must remain contract-valid: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	current, err := ExtractNativePPTX(payload, options)
	if err != nil {
		t.Fatalf("extract with partial-anchor asset Previous: %v", err)
	}
	currentPicture := nativeFixturePicture(t, current.Slides[1])
	if currentPicture.AssetID == nil || *currentPicture.AssetID == "asset-poisoned" {
		t.Fatalf("unmatched slide relationship poisoned the current asset identity: picture=%#v assets=%#v", currentPicture, current.Assets)
	}
	for _, asset := range current.Assets {
		if asset.ID == "asset-poisoned" {
			t.Fatalf("unmatched asset identity survived filtering: %#v", current.Assets)
		}
	}
}

func TestExtractNativePPTXPictureRejectsDuplicatePreviousAssetAnchor(t *testing.T) {
	t.Parallel()

	payload := nativePictureFixture(t, nativePictureFixtureOptions{})
	previous, err := ExtractNativePPTX(payload, nativeTestExtractOptions())
	if err != nil {
		t.Fatalf("initial picture extract: %v", err)
	}
	clone := previous.Assets[0]
	clone.ID = "asset-duplicate-source"
	previous.Assets = append(previous.Assets, clone)
	if issues := ValidateNativePPTX(previous); len(issues) != 0 {
		t.Fatalf("malicious Previous must remain contract-valid for anchor audit: %#v", issues)
	}
	options := nativeTestExtractOptions()
	options.Previous = &previous
	if _, err := ExtractNativePPTX(payload, options); err == nil {
		t.Fatal("duplicate Previous asset source anchor was trusted")
	}
}

func TestNativePPTXPictureAssetBudgetsAreCumulativeAndPreMaterialization(t *testing.T) {
	t.Parallel()

	extractor := nativeExtractor{mediaBytesEmitted: nativeExtractMaxTotalMediaBytes - 1}
	if err := extractor.reserveNativeAssetBudget(2, 0); err == nil || extractor.mediaBytesEmitted != nativeExtractMaxTotalMediaBytes-1 {
		t.Fatalf("media budget was not fail-closed before mutation: err=%v bytes=%d", err, extractor.mediaBytesEmitted)
	}
	extractor = nativeExtractor{assetBase64Emitted: nativeMaxTotalInlineAssetBase64CodeUnits - 1}
	if err := extractor.reserveNativeAssetBudget(0, 2); err == nil || extractor.assetBase64Emitted != nativeMaxTotalInlineAssetBase64CodeUnits-1 {
		t.Fatalf("base64 budget was not fail-closed before mutation: err=%v units=%d", err, extractor.assetBase64Emitted)
	}
	extractor = nativeExtractor{assets: make([]NativeAsset, nativeMaxAssets)}
	if _, err := extractor.nativePictureAsset("media/image.png", "image/png"); err == nil {
		t.Fatal("asset count budget was not checked before package access or capability issuance")
	}
}

type nativePictureFixtureOptions struct {
	stretchContent          *string
	strict                  bool
	secondSlide             bool
	reverseSlides           bool
	pictureCount            int
	imagePart               string
	imageTarget             string
	imageMIME               string
	imageData               string
	secondImagePart         string
	secondImageTarget       string
	secondSlideImagePart    string
	secondSlideImageTarget  string
	relationshipType        string
	relationshipMode        string
	xfrmAttrs               string
	sourceRect              string
	pictureGeometry         string
	blipContent             string
	nonVisualContent        string
	omitImagePart           bool
	duplicateRelationshipID bool
	duplicateBlip           bool
	linkRelationshipID      string
	linkRelationshipType    string
	linkRelationshipMode    string
	omitLinkRelationship    bool
}

func nativePictureFixture(t *testing.T, options nativePictureFixtureOptions) []byte {
	t.Helper()
	relsNS := nsOfficeRelsTransitional
	relType := relImageTransitional
	if options.strict {
		relsNS = nsOfficeRelsStrict
		relType = relImageStrict
	}
	if options.relationshipType != "" {
		relType = options.relationshipType
		if options.strict && options.relationshipType == relChartTransitional {
			relType = relChartStrict
		}
	}
	imagePart := options.imagePart
	if imagePart == "" {
		imagePart = "relocated/media/image.png"
	}
	imageTarget := options.imageTarget
	if imageTarget == "" {
		imageTarget = "../media/image.png"
	}
	imageMIME := options.imageMIME
	if imageMIME == "" {
		imageMIME = "image/png"
	}
	imageData := options.imageData
	if imageData == "" {
		imageData = "\x89PNG\r\n\x1a\nfixture"
	}
	pictureCount := options.pictureCount
	if pictureCount == 0 {
		pictureCount = 1
	}
	extraParts := []nativeExtractZipPart{}
	if !options.omitImagePart {
		extraParts = append(extraParts, nativeExtractZipPart{name: imagePart, data: imageData})
	}
	if options.secondImagePart != "" {
		extraParts = append(extraParts, nativeExtractZipPart{name: options.secondImagePart, data: imageData})
	}
	if options.secondSlideImagePart != "" {
		extraParts = append(extraParts, nativeExtractZipPart{name: options.secondSlideImagePart, data: imageData + "-second-slide"})
	}
	mode := ""
	if options.relationshipMode != "" {
		mode = fmt.Sprintf(` TargetMode="%s"`, options.relationshipMode)
	}
	return nativeExtractFixture(t, nativeExtractFixtureOptions{
		strict: options.strict, secondSlide: options.secondSlide, reverseSlides: options.reverseSlides, extraParts: extraParts,
		mutate: func(parts map[string]string) {
			if !options.omitImagePart {
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/></Types>`, imagePart, imageMIME), 1)
			}
			if options.secondImagePart != "" {
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/></Types>`, options.secondImagePart, imageMIME), 1)
			}
			if options.secondSlideImagePart != "" {
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, fmt.Sprintf(`<Override PartName="/%s" ContentType="%s"/></Types>`, options.secondSlideImagePart, imageMIME), 1)
			}
			addPictures := func(slidePart, relationshipsPart, primaryTarget string) {
				pictures := ""
				relationships := ""
				for index := 0; index < pictureCount; index++ {
					relID := "rIdImage"
					if pictureCount > 1 {
						relID = fmt.Sprintf("rIdImage%d", index+1)
					}
					linkAttribute := ""
					if options.linkRelationshipID != "" {
						linkAttribute = fmt.Sprintf(` r:link="%s"`, options.linkRelationshipID)
					}
					secondBlip := ""
					if options.duplicateBlip {
						secondBlip = fmt.Sprintf(`<a:blip xmlns:r="%s" r:embed="%s"/>`, relsNS, relID)
					}
					pictures += fmt.Sprintf(`<p:pic><p:nvPicPr><p:cNvPr id="%d" name="Photo %d"/><p:cNvPicPr/><p:nvPr>%s</p:nvPr></p:nvPicPr><p:blipFill><a:blip xmlns:r="%s" r:embed="%s"%s>%s</a:blip>%s%s<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm%s><a:off x="1828800" y="914400"/><a:ext cx="3657600" cy="2743200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`, 3+index, index+1, options.nonVisualContent, relsNS, relID, linkAttribute, options.blipContent, secondBlip, options.sourceRect, options.xfrmAttrs)
					target := primaryTarget
					if index == 1 && options.secondImageTarget != "" {
						target = options.secondImageTarget
					}
					relationships += fmt.Sprintf(`<Relationship Id="%s" Type="%s" Target="%s"%s/>`, relID, relType, target, mode)
				}
				if options.linkRelationshipID != "" && !options.omitLinkRelationship {
					linkType := options.linkRelationshipType
					if linkType == "" {
						linkType = map[bool]string{false: relImageTransitional, true: relImageStrict}[options.strict]
					} else if options.strict && linkType == relChartTransitional {
						linkType = relChartStrict
					}
					linkMode := options.linkRelationshipMode
					if linkMode == "" {
						linkMode = "External"
					}
					linkTarget := "https://example.invalid/fallback.png"
					if linkMode == "Internal" {
						linkTarget = imageTarget
					}
					relationships += fmt.Sprintf(`<Relationship Id="%s" Type="%s" Target="%s" TargetMode="%s"/>`, options.linkRelationshipID, linkType, linkTarget, linkMode)
				}
				if options.duplicateRelationshipID {
					relationships += fmt.Sprintf(`<Relationship Id="rIdImage" Type="%s" Target="%s"%s/>`, relType, imageTarget, mode)
				}
				if options.stretchContent != nil {
					pictures = strings.ReplaceAll(pictures, `<a:fillRect/>`, *options.stretchContent)
				}
				if options.pictureGeometry != "" {
					pictures = strings.ReplaceAll(pictures, `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>`, options.pictureGeometry)
				}
				parts[slidePart] = strings.Replace(parts[slidePart], `</p:spTree>`, pictures+`</p:spTree>`, 1)
				parts[relationshipsPart] = strings.Replace(parts[relationshipsPart], `</Relationships>`, relationships+`</Relationships>`, 1)
			}
			addPictures("relocated/slides/slide-a.xml", "relocated/slides/_rels/slide-a.xml.rels", imageTarget)
			if options.secondSlide {
				secondSlideTarget := imageTarget
				if options.secondSlideImageTarget != "" {
					secondSlideTarget = options.secondSlideImageTarget
				}
				addPictures("relocated/slides/slide-b.xml", "relocated/slides/_rels/slide-b.xml.rels", secondSlideTarget)
			}
		},
	})
}

func nativeFixturePicture(t *testing.T, slide NativeSlide) NativeElement {
	t.Helper()
	for _, element := range slide.Elements {
		if element.Kind == NativeElementKindPicture {
			return element
		}
	}
	t.Fatalf("slide has no native picture: %#v", slide)
	return NativeElement{}
}
