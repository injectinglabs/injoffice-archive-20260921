package officecompat

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"fmt"
	"image"
	"image/color"
	"image/png"
)

const visualRatioScale uint64 = 1_000_000

// Keep every ratio cross-product below uint64 overflow even when callers pass
// hostile limit values. Real qualification budgets should be far smaller.
const maxSafeVisualPixels uint64 = ^uint64(0) / (3 * 255 * visualRatioScale)

// VisualLimits bounds the encoded and decoded work accepted by ComparePNG.
// Every field must be positive. Limits apply independently to the reference
// and candidate images.
type VisualLimits struct {
	MaxEncodedBytes uint64 `json:"maxEncodedBytes"`
	MaxWidth        uint32 `json:"maxWidth"`
	MaxHeight       uint32 `json:"maxHeight"`
	MaxPixels       uint64 `json:"maxPixels"`
}

// VisualTolerance defines an explicit, integer-only visual comparison policy.
//
// PerChannelDelta is the largest 8-bit visual RGB channel difference ignored
// when classifying a pixel. MaxDifferentPixelsPPM and
// MaxMeanAbsoluteChannelErrorPPM are millionths in the inclusive range
// [0, 1_000_000]. A zero-value tolerance therefore requires an exact match.
type VisualTolerance struct {
	PerChannelDelta                uint8  `json:"perChannelDelta"`
	MaxDifferentPixelsPPM          uint32 `json:"maxDifferentPixelsPpm"`
	MaxMeanAbsoluteChannelErrorPPM uint32 `json:"maxMeanAbsoluteChannelErrorPpm"`
}

// VisualReport is a stable, machine-readable summary of one PNG comparison.
// Ratios are rounded upward so a non-zero difference is never reported as
// zero. Pixel digests cover normalized visual RGB bytes, not PNG container
// metadata, compression, or transparent hidden color values.
type VisualReport struct {
	Width                       uint32          `json:"width"`
	Height                      uint32          `json:"height"`
	Pixels                      uint64          `json:"pixels"`
	Tolerance                   VisualTolerance `json:"tolerance"`
	DifferentPixels             uint64          `json:"differentPixels"`
	DifferentPixelsPPM          uint32          `json:"differentPixelsPpm"`
	MeanAbsoluteChannelErrorPPM uint32          `json:"meanAbsoluteChannelErrorPpm"`
	MaximumChannelDelta         uint8           `json:"maximumChannelDelta"`
	ReferencePixelSHA256        string          `json:"referencePixelSha256"`
	CandidatePixelSHA256        string          `json:"candidatePixelSha256"`
}

// Matches reports whether the comparison satisfies its recorded tolerance.
// Because report ratios are rounded upward, inclusive integer comparisons
// remain fail-closed at tolerance boundaries.
func (report VisualReport) Matches() bool {
	if report.Pixels == 0 || !validTolerance(report.Tolerance) {
		return false
	}
	if report.DifferentPixelsPPM > report.Tolerance.MaxDifferentPixelsPPM {
		return false
	}
	return report.MeanAbsoluteChannelErrorPPM <= report.Tolerance.MaxMeanAbsoluteChannelErrorPPM
}

// VisualMismatchError reports a valid comparison that exceeded its explicit
// tolerance. Invalid or oversized images are returned as ordinary errors.
type VisualMismatchError struct {
	Report VisualReport
}

func (err *VisualMismatchError) Error() string {
	return fmt.Sprintf(
		"officecompat: visual mismatch: %d/%d pixels differ (%d ppm), mean channel error %d ppm",
		err.Report.DifferentPixels,
		err.Report.Pixels,
		err.Report.DifferentPixelsPPM,
		err.Report.MeanAbsoluteChannelErrorPPM,
	)
}

// ComparePNG compares two externally rendered PNG images. It does not render
// Office documents or invoke a provider. Images are normalized deterministically
// to 8-bit visual RGB by compositing alpha over opaque white; PNG
// metadata, compression, and fully transparent hidden colors do not affect the
// result.
func ComparePNG(reference, candidate []byte, limits VisualLimits, tolerance VisualTolerance) (VisualReport, error) {
	if err := validateVisualLimits(limits); err != nil {
		return VisualReport{}, err
	}
	if !validTolerance(tolerance) {
		return VisualReport{}, fmt.Errorf("officecompat: visual tolerance ppm values must be at most %d", visualRatioScale)
	}

	referenceImage, referenceConfig, err := decodeBoundedPNG("reference", reference, limits)
	if err != nil {
		return VisualReport{}, err
	}
	candidateImage, candidateConfig, err := decodeBoundedPNG("candidate", candidate, limits)
	if err != nil {
		return VisualReport{}, err
	}
	if referenceConfig.Width != candidateConfig.Width || referenceConfig.Height != candidateConfig.Height {
		return VisualReport{}, fmt.Errorf(
			"officecompat: visual dimensions differ: reference %dx%d, candidate %dx%d",
			referenceConfig.Width,
			referenceConfig.Height,
			candidateConfig.Width,
			candidateConfig.Height,
		)
	}

	width := uint32(referenceConfig.Width)
	height := uint32(referenceConfig.Height)
	pixels := uint64(width) * uint64(height)
	referenceDigest := sha256.New()
	candidateDigest := sha256.New()
	var differentPixels uint64
	var channelDeltaSum uint64
	var maximumChannelDelta uint8
	var referenceRGB [3]byte
	var candidateRGB [3]byte

	for y := 0; y < int(height); y++ {
		for x := 0; x < int(width); x++ {
			referenceRGB = visualRGB(referenceImage.At(referenceImage.Bounds().Min.X+x, referenceImage.Bounds().Min.Y+y))
			candidateRGB = visualRGB(candidateImage.At(candidateImage.Bounds().Min.X+x, candidateImage.Bounds().Min.Y+y))
			referenceDigest.Write(referenceRGB[:])
			candidateDigest.Write(candidateRGB[:])

			pixelDifferent := false
			for channel := 0; channel < len(referenceRGB); channel++ {
				delta := absoluteByteDifference(referenceRGB[channel], candidateRGB[channel])
				channelDeltaSum += uint64(delta)
				if delta > maximumChannelDelta {
					maximumChannelDelta = delta
				}
				if delta > tolerance.PerChannelDelta {
					pixelDifferent = true
				}
			}
			if pixelDifferent {
				differentPixels++
			}
		}
	}

	report := VisualReport{
		Width:                       width,
		Height:                      height,
		Pixels:                      pixels,
		Tolerance:                   tolerance,
		DifferentPixels:             differentPixels,
		DifferentPixelsPPM:          uint32(ceilRatio(differentPixels, pixels)),
		MeanAbsoluteChannelErrorPPM: uint32(ceilRatio(channelDeltaSum, pixels*3*255)),
		MaximumChannelDelta:         maximumChannelDelta,
		ReferencePixelSHA256:        hex.EncodeToString(referenceDigest.Sum(nil)),
		CandidatePixelSHA256:        hex.EncodeToString(candidateDigest.Sum(nil)),
	}
	return report, nil
}

// RequirePNGMatch returns VisualMismatchError when valid images exceed the
// supplied tolerance.
func RequirePNGMatch(reference, candidate []byte, limits VisualLimits, tolerance VisualTolerance) error {
	report, err := ComparePNG(reference, candidate, limits, tolerance)
	if err != nil {
		return err
	}
	if !report.Matches() {
		return &VisualMismatchError{Report: report}
	}
	return nil
}

func validateVisualLimits(limits VisualLimits) error {
	if limits.MaxEncodedBytes == 0 || limits.MaxWidth == 0 || limits.MaxHeight == 0 || limits.MaxPixels == 0 {
		return fmt.Errorf("officecompat: every visual resource limit must be positive")
	}
	if limits.MaxPixels > maxSafeVisualPixels {
		return fmt.Errorf("officecompat: visual pixel limit exceeds safe maximum %d", maxSafeVisualPixels)
	}
	return nil
}

func validTolerance(tolerance VisualTolerance) bool {
	return tolerance.MaxDifferentPixelsPPM <= uint32(visualRatioScale) &&
		tolerance.MaxMeanAbsoluteChannelErrorPPM <= uint32(visualRatioScale)
}

func decodeBoundedPNG(label string, encoded []byte, limits VisualLimits) (image.Image, image.Config, error) {
	if uint64(len(encoded)) > limits.MaxEncodedBytes {
		return nil, image.Config{}, fmt.Errorf(
			"officecompat: %s PNG exceeds encoded-byte limit: %d > %d",
			label,
			len(encoded),
			limits.MaxEncodedBytes,
		)
	}
	config, err := png.DecodeConfig(bytes.NewReader(encoded))
	if err != nil {
		return nil, image.Config{}, fmt.Errorf("officecompat: decode %s PNG config: %w", label, err)
	}
	if config.Width <= 0 || config.Height <= 0 {
		return nil, image.Config{}, fmt.Errorf("officecompat: %s PNG has invalid dimensions %dx%d", label, config.Width, config.Height)
	}
	if uint64(config.Width) > uint64(limits.MaxWidth) || uint64(config.Height) > uint64(limits.MaxHeight) {
		return nil, image.Config{}, fmt.Errorf(
			"officecompat: %s PNG dimensions exceed limit: %dx%d (maximum %dx%d)",
			label,
			config.Width,
			config.Height,
			limits.MaxWidth,
			limits.MaxHeight,
		)
	}
	pixels := uint64(config.Width) * uint64(config.Height)
	if pixels > limits.MaxPixels {
		return nil, image.Config{}, fmt.Errorf(
			"officecompat: %s PNG pixel count exceeds limit: %d > %d",
			label,
			pixels,
			limits.MaxPixels,
		)
	}
	decoded, err := png.Decode(bytes.NewReader(encoded))
	if err != nil {
		return nil, image.Config{}, fmt.Errorf("officecompat: decode %s PNG: %w", label, err)
	}
	if decoded.Bounds().Dx() != config.Width || decoded.Bounds().Dy() != config.Height {
		return nil, image.Config{}, fmt.Errorf("officecompat: %s PNG decoded dimensions changed", label)
	}
	return decoded, config, nil
}

func visualRGB(value color.Color) [3]byte {
	nrgba := color.NRGBAModel.Convert(value).(color.NRGBA)
	return [3]byte{
		compositeOverWhite(nrgba.R, nrgba.A),
		compositeOverWhite(nrgba.G, nrgba.A),
		compositeOverWhite(nrgba.B, nrgba.A),
	}
}

func compositeOverWhite(channel, alpha uint8) byte {
	return byte((uint32(channel)*uint32(alpha) + 255*uint32(255-alpha) + 127) / 255)
}

func absoluteByteDifference(left, right byte) uint8 {
	if left >= right {
		return left - right
	}
	return right - left
}

func ceilRatio(numerator, denominator uint64) uint64 {
	if numerator == 0 {
		return 0
	}
	return (numerator*visualRatioScale + denominator - 1) / denominator
}
