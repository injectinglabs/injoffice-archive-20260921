package docxpatch

import "encoding/xml"

// ECMA-376 20.4.2.6 adds these extents to the inline object's layout box,
// without resizing its image. Actual effect-bearing picture markup remains
// independently refused by nativePictureBoundedTransform/blip validation.
func nativeInlineEffectExtents(node *nativeXMLNode) (*NativeDrawingCropV1, bool) {
	if !nativeExactLeaf(node, xml.Name{Local: "l"}, xml.Name{Local: "t"}, xml.Name{Local: "r"}, xml.Name{Local: "b"}) {
		return nil, false
	}
	values := [4]int64{}
	for i, name := range []string{"l", "t", "r", "b"} {
		value, ok := nativeNonnegativeInt64Attr(node, "", name)
		if !ok || value > 91440000 {
			return nil, false
		}
		values[i] = value
	}
	return &NativeDrawingCropV1{Left: nativeInt64(values[0]), Top: nativeInt64(values[1]), Right: nativeInt64(values[2]), Bottom: nativeInt64(values[3])}, true
}
