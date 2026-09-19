package pptxpatch

import "fmt"

// nativePictureCropCode discloses a crop whose authored a:srcRect reaches
// outside the source image and was clamped back onto the representable
// lattice. The move is bounded below one device pixel of the painted picture,
// so it can never stand in for a real outset frame.
const nativePictureCropCode = "pptx.picture-crop-approximate"

// nativePictureCropDevicePixelEmu is one 96 DPI device pixel. A clamp that
// moves the painted image by less than this cannot be seen at any raster this
// preview produces.
const nativePictureCropDevicePixelEmu = 9525

// projectNativeApproximatePictureCrop clamps an a:srcRect inset that reaches
// outside its source image back onto the representable lattice, and reports
// whether the clamp is small enough to paint.
//
// ST_Percentage insets may be negative (ECMA-376 §20.1.10.40 via §20.1.10.41):
// a negative inset OUTSETS the sampled rectangle past the image edge, so the
// destination shows blank beyond it. Native PPTX v1 models a crop as four
// non-negative insets and refused every such picture, including ones whose
// outset is a rounding crumb — `croppedTo0.pptx` authors `r="-1"`, one
// thousandth of one percent of its image.
//
// This is the DOCX approximate image-extent pattern (#315): the projection
// proposes the nearest representable insets and the caller hands them to the
// unmodified crop model. The proposal is kept only when the painted image
// moves by less than one device pixel on both axes, measured against the
// authored sampled fraction, so a genuine outset frame — which really does
// show emptiness the reader cannot paint — still refuses exactly as before.
//
// A degenerate crop (an axis that samples nothing) is never projected: there
// is no nearest representable crop, only a different picture.
func projectNativeApproximatePictureCrop(insets [4]int64, cx, cy int64) (*NativePictureCrop, bool) {
	if cx <= 0 || cy <= 0 {
		return nil, false
	}
	painted := insets
	for index := range painted {
		if painted[index] < 0 {
			painted[index] = 0
		}
	}
	axes := [2]struct{ low, high, extent int }{{0, 2, 0}, {1, 3, 1}}
	extents := [2]int64{cx, cy}
	for _, axis := range axes {
		authored := 100_000 - insets[axis.low] - insets[axis.high]
		clamped := 100_000 - painted[axis.low] - painted[axis.high]
		if authored <= 0 || clamped <= 0 {
			return nil, false
		}
		// The authored sampled fraction maps onto the whole destination edge,
		// so `moved / authored` of that edge is how far the image travels.
		moved := authored - clamped
		if moved < 0 {
			return nil, false
		}
		if extents[axis.extent]*moved >= nativePictureCropDevicePixelEmu*authored {
			return nil, false
		}
	}
	if painted == ([4]int64{}) {
		// Every authored inset was an outset crumb: the picture is uncropped.
		return nil, true
	}
	return &NativePictureCrop{Left: &painted[0], Top: &painted[1], Right: &painted[2], Bottom: &painted[3]}, true
}

// nativeApproximatePictureCropMessage names both the authored insets and the
// painted ones, so the move is auditable from the diagnostic alone.
func nativeApproximatePictureCropMessage(source [4]int64, painted *NativePictureCrop) string {
	values := [4]int64{}
	if painted != nil {
		values = [4]int64{*painted.Left, *painted.Top, *painted.Right, *painted.Bottom}
	}
	return fmt.Sprintf("Read-only approximate picture crop: the authored a:srcRect reaches outside its image at l=%d t=%d r=%d b=%d and is painted at the nearest representable l=%d t=%d r=%d b=%d. The clamp moves the painted image by less than one 96 DPI device pixel on both axes; a larger outset still refuses.",
		source[0], source[1], source[2], source[3], values[0], values[1], values[2], values[3])
}
