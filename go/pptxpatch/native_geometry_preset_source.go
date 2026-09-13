package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// NativePPTXPresetNames returns the complete supported DrawingML preset catalog
// in deterministic order. The returned slice never aliases internal state.
func NativePPTXPresetNames() ([]string, error) { return nativePresetNames() }

// EvaluateNativePPTXPresetGeometry evaluates a preset for caller-authored native
// shape geometry. A containing NativeElement must remain preserve-only until a
// preset/custom-geometry serializer is qualified; evaluation does not grant
// source mutation authority.
func EvaluateNativePPTXPresetGeometry(name string, widthEmu, heightEmu int64, adjustments map[string]int64) (*NativeEvaluatedGeometry, error) {
	return evaluateNativePresetGeometry(name, adjustments, widthEmu, heightEmu)
}

func evaluateNativePresetSource(node *nativeXMLNode, ns string, width, height int64) (*NativeEvaluatedGeometry, error) {
	if node == nil || node.Name != (xml.Name{Space: ns, Local: "prstGeom"}) {
		return nil, fmt.Errorf("missing preset geometry")
	}
	if err := requireOnlyNativeAttrs(node, xml.Name{Local: "prst"}); err != nil {
		return nil, err
	}
	if err := requireOnlyNativeChildren(node, xml.Name{Space: ns, Local: "avLst"}); err != nil {
		return nil, err
	}
	name, ok := exactNativeAttr(node, "", "prst")
	if !ok {
		return nil, fmt.Errorf("missing preset name")
	}
	list, err := nativeSingleton(node, ns, "avLst", false)
	if err != nil {
		return nil, err
	}
	adjustments := map[string]int64{}
	if list != nil {
		if err := requireOnlyNativeAttrs(list); err != nil {
			return nil, err
		}
		if err := requireOnlyNativeChildren(list, xml.Name{Space: ns, Local: "gd"}); err != nil {
			return nil, err
		}
		if len(list.Children) > nativeGeometryMaxGuides {
			return nil, fmt.Errorf("preset adjustment budget exceeded")
		}
		for _, guide := range list.Children {
			if err := requireOnlyNativeAttrs(guide, xml.Name{Local: "name"}, xml.Name{Local: "fmla"}); err != nil {
				return nil, err
			}
			if err := requireOnlyNativeChildren(guide); err != nil {
				return nil, err
			}
			key, ok := exactNativeAttr(guide, "", "name")
			if !ok {
				return nil, fmt.Errorf("missing preset adjustment name")
			}
			if _, exists := adjustments[key]; exists {
				return nil, fmt.Errorf("duplicate preset adjustment")
			}
			formula, _ := exactNativeAttr(guide, "", "fmla")
			fields := strings.Fields(formula)
			if len(fields) != 2 || fields[0] != "val" {
				return nil, fmt.Errorf("preset adjustment requires a literal val formula")
			}
			value, err := strconv.ParseInt(fields[1], 10, 64)
			if err != nil || value > nativeMaxSafeInteger || value < -nativeMaxSafeInteger {
				return nil, fmt.Errorf("invalid preset adjustment value")
			}
			adjustments[key] = value
		}
	}
	return EvaluateNativePPTXPresetGeometry(name, width, height, adjustments)
}
