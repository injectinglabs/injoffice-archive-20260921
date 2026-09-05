package xlsxpatch

import (
	"os"
	"path/filepath"
	"strings"
	"testing"
)

func TestNativeGetCorpusExtractMatchesManifest(t *testing.T) {
	for _, spec := range NativeGetCorpusCases() {
		t.Run(spec.ID, func(t *testing.T) {
			raw, got, err := BuildNativeGetCorpusFile(spec.ID)
			if err != nil {
				t.Fatal(err)
			}
			if got.ID != spec.ID || got.File != spec.File {
				t.Fatalf("manifest row drifted: %#v", got)
			}
			if len(raw) < 100 {
				t.Fatalf("package too small: %d", len(raw))
			}
			workbook, err := ExtractNativeWorkbookV2(raw)
			if spec.Extract == "error" {
				if err == nil {
					t.Fatal("expected extract to refuse")
				}
				if !strings.Contains(err.Error(), spec.ExtractError) {
					t.Fatalf("extract error %v, want %q", err, spec.ExtractError)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if err := nativeGetCorpusCheckUnsupported(workbook.Unsupported, spec); err != nil {
				t.Fatal(err)
			}
			switch spec.DrawingCoverage {
			case "charts":
				charts, err := ReadCharts(raw)
				if err != nil {
					t.Fatal(err)
				}
				if len(charts) != spec.Charts {
					t.Fatalf("charts=%d want %d", len(charts), spec.Charts)
				}
				anchors, err := ReadChartAnchors(raw)
				if err != nil || len(anchors) == 0 {
					t.Fatalf("chart anchors missing: %v %#v", err, anchors)
				}
			case "refuse":
				ins, err := Inspect(raw)
				if err != nil {
					t.Fatal(err)
				}
				if len(ins.DrawingParts) == 0 {
					t.Fatal("shape fixture has no drawing part")
				}
				charts, err := ReadCharts(raw)
				if err != nil {
					t.Fatal(err)
				}
				if len(charts) != 0 {
					t.Fatalf("shape fixture projected charts: %#v", charts)
				}
			case "none":
				ins, err := Inspect(raw)
				if err != nil {
					t.Fatal(err)
				}
				if len(ins.DrawingParts) != 0 || len(ins.ChartParts) != 0 {
					t.Fatalf("table fixture has drawings: %#v", ins)
				}
			}
		})
	}
}

func TestNativeGetCorpusCommittedFilesExtract(t *testing.T) {
	dir := filepath.Join("testdata", "native-get-corpus")
	for _, spec := range NativeGetCorpusCases() {
		t.Run(spec.ID, func(t *testing.T) {
			path := filepath.Join(dir, spec.File)
			raw, err := os.ReadFile(path)
			if err != nil {
				t.Fatalf("missing committed corpus file %s; run go run ./cmd/nativegetcorpus", path)
			}
			workbook, err := ExtractNativeWorkbookV2(raw)
			if spec.Extract == "error" {
				if err == nil || !strings.Contains(err.Error(), spec.ExtractError) {
					t.Fatalf("committed file extract=%v want %q", err, spec.ExtractError)
				}
				return
			}
			if err != nil {
				t.Fatal(err)
			}
			if err := nativeGetCorpusCheckUnsupported(workbook.Unsupported, spec); err != nil {
				t.Fatal(err)
			}
		})
	}
}
