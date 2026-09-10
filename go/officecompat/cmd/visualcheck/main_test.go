package main

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"image"
	"image/color"
	"image/png"
	"os"
	"path/filepath"
	"testing"

	officecompat "github.com/injectinglabs/injoffice/go/officecompat"
)

func hash(data []byte) string { sum := sha256.Sum256(data); return hex.EncodeToString(sum[:]) }

func fixture(t *testing.T) (string, manifest) {
	t.Helper()
	dir := t.TempDir()
	img := image.NewRGBA(image.Rect(0, 0, 2, 2))
	img.Set(0, 0, color.RGBA{255, 0, 0, 255})
	var buf bytes.Buffer
	if err := png.Encode(&buf, img); err != nil {
		t.Fatal(err)
	}
	for name, data := range map[string][]byte{"source.pdf": []byte("test source identity"), "reference.png": buf.Bytes(), "candidate.png": buf.Bytes()} {
		if err := os.WriteFile(filepath.Join(dir, name), data, 0600); err != nil {
			t.Fatal(err)
		}
	}
	return dir, manifest{Version: 1, Cases: []visualCase{{ID: "pdf-page", Format: "pdf", Source: "source.pdf", SourceSHA256: hash([]byte("test source identity")), ReferenceRenderer: "reviewed test renderer v1", CandidateRenderer: "candidate test renderer v1", Limits: officecompat.VisualLimits{MaxEncodedBytes: 1024, MaxWidth: 10, MaxHeight: 10, MaxPixels: 100}, Pages: []pagePair{{Reference: "reference.png", Candidate: "candidate.png", ReferenceSHA256: hash(buf.Bytes())}}}}}
}

func execute(t *testing.T, dir string, spec manifest) (result, error) {
	t.Helper()
	data, err := json.Marshal(spec)
	if err != nil {
		t.Fatal(err)
	}
	path := filepath.Join(dir, "manifest.json")
	if err := os.WriteFile(path, data, 0600); err != nil {
		t.Fatal(err)
	}
	return run(path)
}

func TestIdenticalPagesPass(t *testing.T) {
	dir, spec := fixture(t)
	got, err := execute(t, dir, spec)
	if err != nil || !got.Passed || len(got.Cases[0].Pages) != 1 || got.Cases[0].Pages[0].Report == nil {
		t.Fatalf("%+v %v", got, err)
	}
}

func TestChangedPixelsFail(t *testing.T) {
	dir, spec := fixture(t)
	var buf bytes.Buffer
	if err := png.Encode(&buf, image.NewRGBA(image.Rect(0, 0, 2, 2))); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(dir, "candidate.png"), buf.Bytes(), 0600); err != nil {
		t.Fatal(err)
	}
	got, err := execute(t, dir, spec)
	if err != nil || got.Passed || got.Cases[0].Pages[0].Report.DifferentPixels != 1 {
		t.Fatalf("%+v %v", got, err)
	}
}

func TestInvalidManifestCases(t *testing.T) {
	for name, mutate := range map[string]func(*manifest){
		"empty":       func(m *manifest) { m.Cases = nil },
		"version":     func(m *manifest) { m.Version = 2 },
		"duplicate":   func(m *manifest) { m.Cases = append(m.Cases, m.Cases[0]) },
		"digest":      func(m *manifest) { m.Cases[0].SourceSHA256 = "wrong" },
		"no pages":    func(m *manifest) { m.Cases[0].Pages = nil },
		"no renderer": func(m *manifest) { m.Cases[0].ReferenceRenderer = "" },
		"format":      func(m *manifest) { m.Cases[0].Format = "html" },
		"limits":      func(m *manifest) { m.Cases[0].Limits.MaxPixels = 0 },
		"escape":      func(m *manifest) { m.Cases[0].Source = "../source.pdf" },
	} {
		t.Run(name, func(t *testing.T) {
			dir, spec := fixture(t)
			mutate(&spec)
			if _, err := execute(t, dir, spec); err == nil {
				t.Fatal("accepted invalid manifest")
			}
		})
	}
}

func TestPageFailuresRemainInReport(t *testing.T) {
	for _, kind := range []string{"missing", "bad reference", "dimension mismatch"} {
		t.Run(kind, func(t *testing.T) {
			dir, spec := fixture(t)
			switch kind {
			case "missing":
				spec.Cases[0].Pages[0].Candidate = "missing.png"
			case "bad reference":
				spec.Cases[0].Pages[0].ReferenceSHA256 = "bad"
			case "dimension mismatch":
				var buf bytes.Buffer
				if err := png.Encode(&buf, image.NewRGBA(image.Rect(0, 0, 3, 3))); err != nil {
					t.Fatal(err)
				}
				if err := os.WriteFile(filepath.Join(dir, "candidate.png"), buf.Bytes(), 0600); err != nil {
					t.Fatal(err)
				}
			}
			got, err := execute(t, dir, spec)
			if err != nil || got.Passed || got.Cases[0].Pages[0].Error == "" {
				t.Fatalf("%+v %v", got, err)
			}
		})
	}
}

func TestSymlinkEscape(t *testing.T) {
	dir, spec := fixture(t)
	outside := filepath.Join(t.TempDir(), "source.pdf")
	if err := os.WriteFile(outside, []byte("test source identity"), 0600); err != nil {
		t.Fatal(err)
	}
	if err := os.Symlink(outside, filepath.Join(dir, "link.pdf")); err != nil {
		t.Skip(err)
	}
	spec.Cases[0].Source = "link.pdf"
	if _, err := execute(t, dir, spec); err == nil {
		t.Fatal("followed external symlink")
	}
}
