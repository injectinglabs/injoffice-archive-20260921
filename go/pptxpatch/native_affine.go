package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

// nativeSourceAffine retains source transform semantics independently of the
// legacy quarter-turn transport. ST_Angle uses 60,000 units per degree.
type nativeSourceAffine struct {
	Rotation int64
	FlipH    bool
	FlipV    bool
}

func parseNativeSourceAffine(node *nativeXMLNode) (nativeSourceAffine, error) {
	var result nativeSourceAffine
	if node == nil {
		return result, fmt.Errorf("missing source transform")
	}
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "rot"}, xml.Name{Local: "flipH"}, xml.Name{Local: "flipV"}); err != nil {
		return result, err
	}
	if value, ok := exactNativeAttr(node, "", "rot"); ok {
		angle, err := parseCanonicalNativeInt(value, -2147483648, 2147483647)
		if err != nil {
			return result, fmt.Errorf("invalid source transform angle")
		}
		result.Rotation = (angle%21600000 + 21600000) % 21600000
	}
	for _, item := range []struct {
		name   string
		target *bool
	}{{"flipH", &result.FlipH}, {"flipV", &result.FlipV}} {
		if value, ok := exactNativeAttr(node, "", item.name); ok {
			flag, err := nativeBool(value)
			if err != nil {
				return nativeSourceAffine{}, err
			}
			*item.target = flag
		}
	}
	return result, nil
}
