package main

import (
	"bytes"
	"fmt"
	"os"
	"testing"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func TestFixtureIsReproducibleAndNativeEditable(t *testing.T) {
	first, err := buildPackage()
	if err != nil {
		t.Fatal(err)
	}
	second, err := buildPackage()
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, second) {
		t.Fatal("playground fixture generation is not deterministic")
	}
	checkedIn, err := os.ReadFile("../../testdata/playground_northstar_review.pptx")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(first, checkedIn) {
		t.Fatal("checked-in playground fixture is stale; run go run ./cmd/playground_sample")
	}

	token := 0
	deck, err := pptxpatch.ExtractNativePPTX(first, pptxpatch.NativePPTXExtractOptions{
		TokenFactory: pptxpatch.NativePassthroughTokenFactoryFunc(func(pptxpatch.NativePassthroughTokenRequest) (string, error) {
			token++
			return fmt.Sprintf("fixture-token-%d", token), nil
		}),
	})
	if err != nil {
		t.Fatal(err)
	}
	if len(deck.Slides) != 3 {
		t.Fatalf("slides=%d, want 3", len(deck.Slides))
	}
	editable := 0
	var reviewTitle, decisionMarker bool
	for _, slide := range deck.Slides {
		for _, element := range slide.Elements {
			if element.Compatibility.Status != pptxpatch.NativeCompatibilityStatusEditable {
				continue
			}
			editable++
			if element.Name != nil && *element.Name == "Review title" {
				reviewTitle = true
			}
			if element.Name != nil && *element.Name == "Decision marker" {
				decisionMarker = true
			}
		}
	}
	if editable != 20 || !reviewTitle || !decisionMarker {
		t.Fatalf("editable=%d reviewTitle=%v decisionMarker=%v", editable, reviewTitle, decisionMarker)
	}
}
