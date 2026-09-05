package officecompat_test

import (
	"bytes"
	"encoding/json"
	"errors"
	"image"
	"image/color"
	"image/png"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat"
)

var visualTestLimits = officecompat.VisualLimits{
	MaxEncodedBytes: 1 << 20,
	MaxWidth:        100,
	MaxHeight:       100,
	MaxPixels:       10_000,
}

func TestVisualComparePNGIsStableAcrossContainerEncoding(t *testing.T) {
	pixels := []color.NRGBA{
		{R: 12, G: 34, B: 56, A: 255},
		{R: 200, G: 100, B: 50, A: 128},
		{R: 0, G: 0, B: 0, A: 0},
		{R: 255, G: 255, B: 255, A: 255},
	}
	reference := encodePNG(t, 2, 2, pixels, png.BestSpeed)
	candidate := encodePNG(t, 2, 2, pixels, png.BestCompression)

	report, err := officecompat.ComparePNG(reference, candidate, visualTestLimits, officecompat.VisualTolerance{})
	if err != nil {
		t.Fatal(err)
	}
	if !report.Matches() {
		t.Fatalf("identical pixels should match exactly: %+v", report)
	}
	if report.DifferentPixels != 0 || report.DifferentPixelsPPM != 0 || report.MeanAbsoluteChannelErrorPPM != 0 || report.MaximumChannelDelta != 0 {
		t.Fatalf("unexpected difference: %+v", report)
	}
	if report.ReferencePixelSHA256 == "" || report.ReferencePixelSHA256 != report.CandidatePixelSHA256 {
		t.Fatalf("unexpected normalized digests: %+v", report)
	}
}

func TestVisualComparePNGHonorsExplicitIntegerTolerance(t *testing.T) {
	referencePixels := solidPixels(4, color.NRGBA{R: 100, G: 100, B: 100, A: 255})
	candidatePixels := append([]color.NRGBA(nil), referencePixels...)
	candidatePixels[0].R = 105
	reference := encodePNG(t, 2, 2, referencePixels, png.DefaultCompression)
	candidate := encodePNG(t, 2, 2, candidatePixels, png.DefaultCompression)
	tolerance := officecompat.VisualTolerance{
		PerChannelDelta:                3,
		MaxDifferentPixelsPPM:          250_000,
		MaxMeanAbsoluteChannelErrorPPM: 2_000,
	}

	report, err := officecompat.ComparePNG(reference, candidate, visualTestLimits, tolerance)
	if err != nil {
		t.Fatal(err)
	}
	if report.DifferentPixels != 1 || report.DifferentPixelsPPM != 250_000 {
		t.Fatalf("different pixel report: %+v", report)
	}
	if report.MeanAbsoluteChannelErrorPPM != 1_634 || report.MaximumChannelDelta != 5 {
		t.Fatalf("channel delta report: %+v", report)
	}
	if !report.Matches() {
		t.Fatalf("boundary tolerance should be inclusive: %+v", report)
	}
	if err := officecompat.RequirePNGMatch(reference, candidate, visualTestLimits, tolerance); err != nil {
		t.Fatalf("boundary tolerance should pass: %v", err)
	}

	tooStrict := tolerance
	tooStrict.MaxDifferentPixelsPPM--
	err = officecompat.RequirePNGMatch(reference, candidate, visualTestLimits, tooStrict)
	var mismatch *officecompat.VisualMismatchError
	if !errors.As(err, &mismatch) {
		t.Fatalf("expected VisualMismatchError, got %v", err)
	}
	expectedMismatchReport := report
	expectedMismatchReport.Tolerance = tooStrict
	if mismatch.Report != expectedMismatchReport || !strings.Contains(err.Error(), "1/4 pixels differ (250000 ppm)") {
		t.Fatalf("unstable mismatch evidence: %#v, %v", mismatch, err)
	}
}

func TestVisualComparePNGMeanErrorGateIsIndependent(t *testing.T) {
	reference := encodePNG(t, 1, 1, solidPixels(1, color.NRGBA{A: 255}), png.DefaultCompression)
	candidate := encodePNG(t, 1, 1, solidPixels(1, color.NRGBA{R: 255, A: 255}), png.DefaultCompression)
	tolerance := officecompat.VisualTolerance{
		PerChannelDelta:                255,
		MaxDifferentPixelsPPM:          0,
		MaxMeanAbsoluteChannelErrorPPM: 333_332,
	}

	report, err := officecompat.ComparePNG(reference, candidate, visualTestLimits, tolerance)
	if err != nil {
		t.Fatal(err)
	}
	if report.DifferentPixels != 0 || report.MeanAbsoluteChannelErrorPPM != 333_334 {
		t.Fatalf("unexpected independent metrics: %+v", report)
	}
	if report.Matches() {
		t.Fatal("mean channel error must fail even when the pixel classifier ignores the delta")
	}
}

func TestVisualComparePNGUsesVisualAlphaOverWhite(t *testing.T) {
	reference := encodePNG(t, 1, 1, []color.NRGBA{{R: 255, A: 0}}, png.DefaultCompression)
	candidate := encodePNG(t, 1, 1, []color.NRGBA{{B: 255, A: 0}}, png.DefaultCompression)

	report, err := officecompat.ComparePNG(reference, candidate, visualTestLimits, officecompat.VisualTolerance{})
	if err != nil {
		t.Fatal(err)
	}
	if !report.Matches() || report.ReferencePixelSHA256 != report.CandidatePixelSHA256 {
		t.Fatalf("transparent hidden colors should have identical visual pixels: %+v", report)
	}
}

func TestVisualComparePNGRejectsDimensionMismatch(t *testing.T) {
	reference := encodePNG(t, 1, 1, solidPixels(1, color.NRGBA{A: 255}), png.DefaultCompression)
	candidate := encodePNG(t, 2, 1, solidPixels(2, color.NRGBA{A: 255}), png.DefaultCompression)

	_, err := officecompat.ComparePNG(reference, candidate, visualTestLimits, officecompat.VisualTolerance{})
	if err == nil || err.Error() != "officecompat: visual dimensions differ: reference 1x1, candidate 2x1" {
		t.Fatalf("unexpected error: %v", err)
	}
}

func TestVisualComparePNGRejectsInvalidAndTruncatedInput(t *testing.T) {
	valid := encodePNG(t, 1, 1, solidPixels(1, color.NRGBA{A: 255}), png.DefaultCompression)
	tests := []struct {
		name      string
		input     []byte
		wantError string
	}{
		{name: "not png", input: []byte("not a png"), wantError: "decode reference PNG config"},
		{name: "truncated", input: valid[:len(valid)-4], wantError: "decode reference PNG"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := officecompat.ComparePNG(test.input, valid, visualTestLimits, officecompat.VisualTolerance{})
			if err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("expected %q, got %v", test.wantError, err)
			}
		})
	}
}

func TestVisualComparePNGEnforcesResourceBoundsBeforePixelDecode(t *testing.T) {
	large := encodePNG(t, 3, 3, solidPixels(9, color.NRGBA{R: 20, A: 255}), png.NoCompression)
	tests := []struct {
		name      string
		limits    officecompat.VisualLimits
		wantError string
	}{
		{
			name:      "encoded bytes",
			limits:    officecompat.VisualLimits{MaxEncodedBytes: uint64(len(large) - 1), MaxWidth: 10, MaxHeight: 10, MaxPixels: 100},
			wantError: "exceeds encoded-byte limit",
		},
		{
			name:      "width",
			limits:    officecompat.VisualLimits{MaxEncodedBytes: 1 << 20, MaxWidth: 2, MaxHeight: 10, MaxPixels: 100},
			wantError: "dimensions exceed limit",
		},
		{
			name:      "height",
			limits:    officecompat.VisualLimits{MaxEncodedBytes: 1 << 20, MaxWidth: 10, MaxHeight: 2, MaxPixels: 100},
			wantError: "dimensions exceed limit",
		},
		{
			name:      "pixels",
			limits:    officecompat.VisualLimits{MaxEncodedBytes: 1 << 20, MaxWidth: 10, MaxHeight: 10, MaxPixels: 8},
			wantError: "pixel count exceeds limit",
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := officecompat.ComparePNG(large, large, test.limits, officecompat.VisualTolerance{})
			if err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("expected %q, got %v", test.wantError, err)
			}
		})
	}
}

func TestVisualComparePNGRejectsUnsafePolicyValues(t *testing.T) {
	valid := encodePNG(t, 1, 1, solidPixels(1, color.NRGBA{A: 255}), png.DefaultCompression)
	tests := []struct {
		name      string
		limits    officecompat.VisualLimits
		tolerance officecompat.VisualTolerance
		wantError string
	}{
		{name: "zero limits", limits: officecompat.VisualLimits{}, wantError: "every visual resource limit must be positive"},
		{
			name:      "overflow-scale pixel limit",
			limits:    officecompat.VisualLimits{MaxEncodedBytes: 1, MaxWidth: 1, MaxHeight: 1, MaxPixels: ^uint64(0)},
			wantError: "visual pixel limit exceeds safe maximum",
		},
		{
			name:      "different ppm",
			limits:    visualTestLimits,
			tolerance: officecompat.VisualTolerance{MaxDifferentPixelsPPM: 1_000_001},
			wantError: "visual tolerance ppm values must be at most 1000000",
		},
		{
			name:      "mean ppm",
			limits:    visualTestLimits,
			tolerance: officecompat.VisualTolerance{MaxMeanAbsoluteChannelErrorPPM: 1_000_001},
			wantError: "visual tolerance ppm values must be at most 1000000",
		},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := officecompat.ComparePNG(valid, valid, test.limits, test.tolerance)
			if err == nil || !strings.Contains(err.Error(), test.wantError) {
				t.Fatalf("expected %q, got %v", test.wantError, err)
			}
		})
	}
}

func TestVisualReportJSONIsStableAndIntegerOnly(t *testing.T) {
	reference := encodePNG(t, 1, 1, []color.NRGBA{{R: 10, G: 20, B: 30, A: 255}}, png.DefaultCompression)
	candidate := encodePNG(t, 1, 1, []color.NRGBA{{R: 11, G: 20, B: 30, A: 255}}, png.DefaultCompression)
	report, err := officecompat.ComparePNG(reference, candidate, visualTestLimits, officecompat.VisualTolerance{})
	if err != nil {
		t.Fatal(err)
	}
	encoded, err := json.Marshal(report)
	if err != nil {
		t.Fatal(err)
	}
	want := `{"width":1,"height":1,"pixels":1,"tolerance":{"perChannelDelta":0,"maxDifferentPixelsPpm":0,"maxMeanAbsoluteChannelErrorPpm":0},"differentPixels":1,"differentPixelsPpm":1000000,"meanAbsoluteChannelErrorPpm":1308,"maximumChannelDelta":1,"referencePixelSha256":"6951bbd9c2178b2a9d01714687307afda1eb6161ff8c398486a76683e7b42925","candidatePixelSha256":"7a4e42ecc489870c12f800d342e617162943f3e5f281cd133c6f4db1ffc54118"}`
	if string(encoded) != want {
		t.Fatalf("report is not stable integer JSON: %s", encoded)
	}
}

func encodePNG(t *testing.T, width, height int, pixels []color.NRGBA, compression png.CompressionLevel) []byte {
	t.Helper()
	if len(pixels) != width*height {
		t.Fatalf("pixel count %d does not match %dx%d", len(pixels), width, height)
	}
	value := image.NewNRGBA(image.Rect(0, 0, width, height))
	for index, pixel := range pixels {
		value.SetNRGBA(index%width, index/width, pixel)
	}
	var encoded bytes.Buffer
	encoder := png.Encoder{CompressionLevel: compression}
	if err := encoder.Encode(&encoded, value); err != nil {
		t.Fatal(err)
	}
	return encoded.Bytes()
}

func solidPixels(count int, value color.NRGBA) []color.NRGBA {
	pixels := make([]color.NRGBA, count)
	for index := range pixels {
		pixels[index] = value
	}
	return pixels
}
