package docxpatch

import (
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"testing"
)

// The host default size has one definition per language, and the two must
// agree: the Go extractor names the proven source shape and the TypeScript
// consumer applies the size. A drift here would silently re-size every package
// that states no size, so it is pinned rather than trusted.
func TestHostDefaultSizeHalfPointsMatchTypeScript(t *testing.T) {
	path := filepath.Join("..", "..", "packages", "docs", "src", "nativeAbsentFontSizeV1.ts")
	source, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	pattern := regexp.MustCompile(`"([a-z-]+)":\s*(\d+),`)
	found := map[string]int{}
	for _, match := range pattern.FindAllStringSubmatch(string(source), -1) {
		halfPoints, err := strconv.Atoi(match[2])
		if err != nil {
			t.Fatal(err)
		}
		found[match[1]] = halfPoints
	}
	// Read out of the Tf operators of Microsoft Word 16.112's own PDF exports,
	// which place text on a 1/300 in grid: a package with no w:docDefaults
	// record is written at 50 units (12 pt), one whose record states no w:sz at
	// 42 units (10 pt).
	want := map[string]int{NativeDocxAbsentDocumentDefaultsV1: 24, NativeDocxSizelessDocumentDefaultsV1: 20}
	if len(found) != len(want) {
		t.Fatalf("TypeScript declares %d host defaults, Go declares %d: %v", len(found), len(want), found)
	}
	for shape, halfPoints := range want {
		if found[shape] != halfPoints {
			t.Fatalf("%s: TypeScript %d half-points, Go %d", shape, found[shape], halfPoints)
		}
		if goHalfPoints, ok := NativeDocxHostDefaultSizeHalfPointsV1(shape); !ok || goHalfPoints != halfPoints {
			t.Fatalf("%s: Go resolver returned %d/%v", shape, goHalfPoints, ok)
		}
	}
	if _, ok := NativeDocxHostDefaultSizeHalfPointsV1(""); ok {
		t.Fatal("an unproven shape must select no host default")
	}
}
