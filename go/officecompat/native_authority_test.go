package officecompat_test

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

// The corpus is evidence for native contract extraction. Keep the production
// extraction boundary free of reconstructive browser/document pipelines.
func TestNativeOfficeExtractionHasNoLegacyReconstructionDependency(t *testing.T) {
	patterns := []string{
		filepath.Join("..", "docxpatch", "native_extract*.go"),
		filepath.Join("..", "pptxpatch", "native_extract*.go"),
		filepath.Join("..", "xlsxpatch", "native_extract*.go"),
	}
	forbidden := []string{
		"mammoth", "luckyexcel", "jszip", "domparser", "innerhtml",
		"outerhtml", "document.createelement", "golang.org/x/net/html",
	}
	for _, pattern := range patterns {
		paths, err := filepath.Glob(pattern)
		if err != nil {
			t.Fatal(err)
		}
		if len(paths) == 0 {
			t.Fatalf("native extractor source missing for %s", pattern)
		}
		for _, path := range paths {
			if strings.HasSuffix(path, "_test.go") {
				continue
			}
			data, err := os.ReadFile(path)
			if err != nil {
				t.Fatal(err)
			}
			lower := strings.ToLower(string(data))
			for _, token := range forbidden {
				if strings.Contains(lower, token) {
					t.Fatalf("native extractor %s contains forbidden legacy dependency %q", path, token)
				}
			}
		}
	}
}
