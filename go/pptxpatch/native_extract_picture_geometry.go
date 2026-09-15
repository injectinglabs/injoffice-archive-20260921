package pptxpatch

import (
	"encoding/xml"
	"unicode/utf8"
)

// nativePictureGeometryPreviewCode marks a picture whose non-rectangular
// prstGeom outline was evaluated through the fingerprinted DrawingML preset
// catalog. The outline is a read-only clip preview: the picture stays
// preserve-only, no source bytes change, and mutation keeps refusing every
// element that carries evaluated geometry.
const nativePictureGeometryPreviewCode = "pptx.picture-geometry-preview"

// nativePictureGeometryDiagnosticBound keeps an untrusted preset name or
// evaluator error from pushing a diagnostic past the contract message bound.
const nativePictureGeometryDiagnosticBound = 512

// validateNativePictureGeometry keeps the exact rect / default roundRect
// contract unchanged and routes every other preset, including authored
// adjustments, through the same bounded catalog evaluator as AutoShapes.
// Evaluation failure retains the legacy preserve-only gap so the renderer
// still refuses the picture instead of guessing an outline.
//
// ECMA-376 CT_PresetGeometry2D declares avLst with minOccurs=0; an omitted list
// means default adjustments exactly like an empty one.
func validateNativePictureGeometry(geometry *nativeXMLNode, dialect nativeExtractDialect, width, height int64, transformExact bool, gaps *nativePictureGapSet, clip **string, evaluated **NativeEvaluatedGeometry) error {
	adjustments, err := nativeSingleton(geometry, dialect.drawing, "avLst", false)
	if err != nil {
		return err
	}
	preset, presetOK := exactNativeAttr(geometry, "", "prst")
	exactMarkup := presetOK &&
		requireOnlyNativeAttrs(geometry, xml.Name{Local: "prst"}) == nil &&
		requireOnlyNativeChildren(geometry, xml.Name{Space: dialect.drawing, Local: "avLst"}) == nil &&
		(adjustments == nil || requireEmptyNativeElement(adjustments) == nil)
	switch {
	case exactMarkup && preset == "rect":
		return nil
	case exactMarkup && preset == "roundRect":
		// ECMA presetShapeDefinitions: default adj=16667; x1=ss*adj/100000.
		// Keep the source preset, not a generic approximate corner radius.
		*clip = stringPointer("roundRect")
		return nil
	}
	if !transformExact {
		gaps.add("pptx.picture-geometry-unavailable", "non-rectangle picture geometry with rotation or flips is preserved but not evaluated in native PPTX v1")
		return nil
	}
	result, evalErr := evaluateNativePresetSource(geometry, dialect.drawing, width, height)
	if evalErr != nil {
		gaps.add("pptx.picture-geometry-unavailable", "picture geometry is outside the evaluated preset catalog: "+nativeBoundedDiagnosticText(evalErr.Error(), nativePictureGeometryDiagnosticBound))
		return nil
	}
	*evaluated = result
	gaps.add(nativePictureGeometryPreviewCode, "DrawingML preset catalog outline evaluated from source as a read-only picture clip; the picture remains preserve-only")
	return nil
}

func nativeBoundedDiagnosticText(text string, limit int) string {
	if utf8.RuneCountInString(text) <= limit {
		return text
	}
	runes := []rune(text)
	return string(runes[:limit]) + "…"
}
