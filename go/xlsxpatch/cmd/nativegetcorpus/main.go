// Command nativegetcorpus writes the native GET edge-case .xlsx files plus
// extract JSON the playground and tests share.
package main

import (
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, "nativegetcorpus:", err)
		os.Exit(1)
	}
}

func run() error {
	root, err := repoRoot()
	if err != nil {
		return err
	}
	testdata := filepath.Join(root, "go", "xlsxpatch", "testdata", "native-get-corpus")
	playground := filepath.Join(root, "apps", "playground", "public", "native-corpus")
	if err := os.MkdirAll(testdata, 0o755); err != nil {
		return err
	}
	if err := os.MkdirAll(playground, 0o755); err != nil {
		return err
	}

	cases := xlsxpatch.NativeGetCorpusCases()
	manifest, err := json.MarshalIndent(struct {
		Cases []xlsxpatch.NativeGetCorpusCase `json:"cases"`
	}{Cases: cases}, "", "  ")
	if err != nil {
		return err
	}
	manifest = append(manifest, '\n')
	if err := os.WriteFile(filepath.Join(testdata, "MANIFEST.json"), manifest, 0o644); err != nil {
		return err
	}
	if err := os.WriteFile(filepath.Join(playground, "manifest.json"), manifest, 0o644); err != nil {
		return err
	}

	for _, spec := range cases {
		raw, _, err := xlsxpatch.BuildNativeGetCorpusFile(spec.ID)
		if err != nil {
			return fmt.Errorf("%s: %w", spec.ID, err)
		}
		if err := os.WriteFile(filepath.Join(testdata, spec.File), raw, 0o644); err != nil {
			return err
		}
		if err := os.WriteFile(filepath.Join(playground, spec.File), raw, 0o644); err != nil {
			return err
		}
		if spec.Extract != "ok" {
			continue
		}
		workbook, err := xlsxpatch.ExtractNativeWorkbookV2(raw)
		if err != nil {
			return fmt.Errorf("%s extract: %w", spec.ID, err)
		}
		encoded, err := xlsxpatch.EncodeNativeWorkbookV2(workbook)
		if err != nil {
			return err
		}
		jsonName := spec.ID + ".workbook.json"
		if err := os.WriteFile(filepath.Join(testdata, jsonName), append(encoded, '\n'), 0o644); err != nil {
			return err
		}
		if err := os.WriteFile(filepath.Join(playground, jsonName), append(encoded, '\n'), 0o644); err != nil {
			return err
		}
		if spec.DrawingCoverage != "charts" {
			continue
		}
		charts, err := xlsxpatch.ReadCharts(raw)
		if err != nil {
			return err
		}
		anchors, err := xlsxpatch.ReadChartAnchors(raw)
		if err != nil {
			return err
		}
		envelope, err := json.MarshalIndent(struct {
			Charts  []xlsxpatch.ChartInfo            `json:"charts"`
			Anchors map[string]xlsxpatch.ChartAnchor `json:"anchors"`
		}{Charts: charts, Anchors: anchors}, "", "  ")
		if err != nil {
			return err
		}
		chartName := spec.ID + ".charts.json"
		if err := os.WriteFile(filepath.Join(testdata, chartName), append(envelope, '\n'), 0o644); err != nil {
			return err
		}
		if err := os.WriteFile(filepath.Join(playground, chartName), append(envelope, '\n'), 0o644); err != nil {
			return err
		}
	}
	fmt.Fprintf(os.Stderr, "wrote %d corpus files to %s and %s\n", len(cases), testdata, playground)
	return nil
}

func repoRoot() (string, error) {
	wd, err := os.Getwd()
	if err != nil {
		return "", err
	}
	dir := wd
	for {
		if _, err := os.Stat(filepath.Join(dir, "go", "xlsxpatch")); err == nil {
			if _, err := os.Stat(filepath.Join(dir, "apps", "playground")); err == nil {
				return dir, nil
			}
		}
		parent := filepath.Dir(dir)
		if parent == dir {
			return "", fmt.Errorf("run from the injoffice repository")
		}
		dir = parent
	}
}
