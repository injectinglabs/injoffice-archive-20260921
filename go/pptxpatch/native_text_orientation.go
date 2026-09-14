package pptxpatch

import "fmt"

// Body rotation is a signed ST_Angle, retained verbatim as an integer. It is
// separate from the shape transform and only normalized for trigonometry.
// The text-layout extractor still owns all other bodyPr admission rules.
type nativeTextOrientation struct {
	Rotation    int64
	Upright     bool
	HasRotation bool
	HasUpright  bool
}

func parseNativeTextOrientation(node *nativeXMLNode) (nativeTextOrientation, error) {
	var result nativeTextOrientation
	if node == nil {
		return result, fmt.Errorf("missing text body properties")
	}
	for _, attribute := range node.Attrs {
		switch attribute.Name.Local {
		case "rot":
			if attribute.Name.Space != "" || result.HasRotation {
				return nativeTextOrientation{}, fmt.Errorf("ambiguous text rotation attribute")
			}
			result.HasRotation = true
			value, err := parseCanonicalNativeInt(attribute.Value, -2147483648, 2147483647)
			if err != nil {
				return nativeTextOrientation{}, fmt.Errorf("invalid text rotation")
			}
			result.Rotation = value
		case "upright":
			if attribute.Name.Space != "" || result.HasUpright {
				return nativeTextOrientation{}, fmt.Errorf("ambiguous text upright attribute")
			}
			result.HasUpright = true
			value, err := nativeBool(attribute.Value)
			if err != nil {
				return nativeTextOrientation{}, fmt.Errorf("invalid text upright flag")
			}
			result.Upright = value
		}
	}
	return result, nil
}
func (orientation nativeTextOrientation) normalizedRotation() int64 {
	return (orientation.Rotation%21600000 + 21600000) % 21600000
}

func nativeTextBodyHasOrientation(body *NativeTextBodyLayout) bool {
	return body != nil && ((body.RotationAngle60000 != nil && *body.RotationAngle60000 != 0) || (body.Upright != nil && *body.Upright))
}
