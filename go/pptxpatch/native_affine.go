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

// Only new affine/rational source groups widen the preview boundary. Keep the
// previously supported exact authored/parsed group mutation rules unchanged.
func nativeComplexAffineGroup(element NativeElement) bool {
	if element.Kind != NativeElementKindGroup {
		return false
	}
	if nativeHasSourceAffine(element.Transform) {
		return true
	}
	if element.ChildTransform != nil {
		_, _, _, _, err := nativeGroupAffineComponents(element.Transform, *element.ChildTransform)
		return err != nil
	}
	return false
}
func nativeHasGraphicFrameDescendant(elements []NativeElement) bool {
	for _, element := range elements {
		if element.Kind == NativeElementKindTable || element.Kind == NativeElementKindChart || nativeHasGraphicFrameDescendant(element.Children) {
			return true
		}
	}
	return false
}
func nativeMarkAffineDescendants(elements []NativeElement) {
	for index := range elements {
		element := &elements[index]
		element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
		// One inherited marker is sufficient even across nested affine groups.
		marked := false
		for _, diagnostic := range element.Compatibility.Diagnostics {
			if diagnostic.Code == "pptx.source-affine-ancestor-preview" {
				marked = true
			}
		}
		if !marked && len(element.Compatibility.Diagnostics) < nativeMaxDiagnosticsPerScope {
			element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: "pptx.source-affine-ancestor-preview", Message: "an ancestor source affine transform makes this target preview-only"})
		}
		nativeMarkAffineDescendants(element.Children)
	}
}
