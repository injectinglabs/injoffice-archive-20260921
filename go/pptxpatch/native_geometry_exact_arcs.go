package pptxpatch

import "math"

// An integer-looking float alone is not evidence of exact source arithmetic.
func nativeGeometryExactIntegerAttribute(g nativeGeometryGuides, node *nativeXMLNode, key string, value float64) bool {
	token, ok := exactNativeAttr(node, "", key)
	if !ok || !nativeGeometryFinite(value) || math.Trunc(value) != value {
		return false
	}
	exact := g.exactOperand(token)
	return exact != nil && exact.IsInt() && exact.Num().IsInt64() && exact.Num().Int64() == int64(value)
}

// Certify the operations actually used by arc: cardinal offsets, center
// subtraction, and quarter-turn endpoint addition are safe integer arithmetic.
// The caller still charges all input and previously accumulated uncertainty,
// and arc still enforces radius fit, coordinate and command limits.
func nativeGeometryExactCardinalArc(g nativeGeometryGuides, node *nativeXMLNode, p nativeGeometryPathBuilder, exactPen bool, values []float64) bool {
	if !exactPen || !p.hasPen || p.sx != 1 || p.sy != 1 || len(values) != 4 {
		return false
	}
	for i, key := range []string{"wR", "hR", "stAng", "swAng"} {
		if !nativeGeometryExactIntegerAttribute(g, node, key, values[i]) {
			return false
		}
	}
	const quarter int64 = 5400000
	const turn int64 = 4 * quarter
	rx, ry, start, sweep := int64(values[0]), int64(values[1]), int64(values[2]), int64(values[3])
	if rx <= 0 || ry <= 0 || start%quarter != 0 || sweep%quarter != 0 || sweep < -turn || sweep > turn {
		return false
	}
	safe := func(v int64) bool { return v >= -9007199254740991 && v <= 9007199254740991 }
	if !nativeGeometryFinite(p.pen.x) || !nativeGeometryFinite(p.pen.y) || math.Trunc(p.pen.x) != p.pen.x || math.Trunc(p.pen.y) != p.pen.y {
		return false
	}
	offset := func(angle int64) (int64, int64) {
		switch (angle%turn + turn) % turn {
		case 0:
			return rx, 0
		case quarter:
			return 0, ry
		case 2 * quarter:
			return -rx, 0
		default:
			return 0, -ry
		}
	}
	start %= turn
	ox, oy := offset(start)
	cx, cy := int64(p.pen.x)-ox, int64(p.pen.y)-oy
	if !safe(cx) || !safe(cy) {
		return false
	}
	segments := sweep / quarter
	direction := int64(1)
	if segments < 0 {
		segments = -segments
		direction = -1
	}
	for i := int64(1); i <= segments; i++ {
		x, y := offset(start + i*quarter*direction)
		if !safe(cx+x) || !safe(cy+y) {
			return false
		}
	}
	return true
}
