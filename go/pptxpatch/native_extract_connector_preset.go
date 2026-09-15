package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

// nativeConnectorPresetPreviewCode is the declared read-only policy for
// connectors whose geometry comes from the fingerprinted DrawingML preset
// catalog rather than the exact straight-line contract.
const nativeConnectorPresetPreviewCode = "pptx.connector-preset-preview"

// nativeConnectorPresetFamily is the closed set of DrawingML connector presets
// a <p:cxnSp> may route through the catalog evaluator. Every member is a single
// open fill="none" path, so a connector can never acquire a painted interior.
var nativeConnectorPresetFamily = map[string]bool{
	"line": true, "straightConnector1": true,
	"bentConnector2": true, "bentConnector3": true, "bentConnector4": true, "bentConnector5": true,
	"curvedConnector2": true, "curvedConnector3": true, "curvedConnector4": true, "curvedConnector5": true,
}

// nativeConnectorUsesExactStraightPath reports whether the connector still
// qualifies for the pre-existing exact editable projection: a straight preset
// with an empty adjustment list and no rotation. Custom or missing geometry also
// stays on that path so its existing refusal semantics are unchanged.
func nativeConnectorUsesExactStraightPath(xfrm, preset, custom *nativeXMLNode, dialect nativeExtractDialect) bool {
	if custom != nil || preset == nil {
		return true
	}
	name, _ := exactNativeAttr(preset, "", "prst")
	if name != "line" && name != "straightConnector1" {
		return false
	}
	adjustments := nativeChild(preset, dialect.drawing, "avLst")
	if adjustments == nil || requireEmptyNativeElement(adjustments) != nil {
		return false
	}
	if rotation, ok := exactNativeAttr(xfrm, "", "rot"); ok && rotation != "0" {
		return false
	}
	return true
}

// validateNativeConnectorPresetTransform projects the source frame and its
// orientation as a bounded rational affine preview. Connectors never use the
// legacy quarter-turn transport, which the contract restricts to text/shapes.
func validateNativeConnectorPresetTransform(node *nativeXMLNode, dialect nativeExtractDialect, gaps *nativeConnectorGapSet) (NativeTransform, error) {
	x, y, cx, cy, err := parseNativeConnectorFrame(node, dialect)
	if err != nil {
		return NativeTransform{}, err
	}
	// Unmodeled transform attributes or non-canonical orientation values stay
	// an element-level refusal with zero orientation; they never abort the slide.
	orientation, err := parseNativeSourceAffine(node)
	if err != nil {
		gaps.add("pptx.connector-transform-unavailable", "connector transform attributes or orientation are outside the exact subset", true)
		orientation = nativeSourceAffine{}
	}
	result := NativeTransform{X: int64Pointer(x), Y: int64Pointer(y), Cx: int64Pointer(cx), Cy: int64Pointer(cy)}
	if orientation.Rotation != 0 || orientation.FlipH || orientation.FlipV {
		result.RotationAngle = int64Pointer(orientation.Rotation)
		if orientation.FlipH {
			result.FlipH = boolPointer(true)
		}
		if orientation.FlipV {
			result.FlipV = boolPointer(true)
		}
		gaps.add("pptx.source-affine-preview", "DrawingML orientation uses bounded rational affine preview; transformed targets remain read-only", false)
	}
	return result, nil
}

// evaluateNativeConnectorPreset routes a connector-family preset with its
// literal adjustments through the fingerprinted catalog. The result must be a
// single open stroked path; anything else refuses rather than approximating.
func evaluateNativeConnectorPreset(preset *nativeXMLNode, dialect nativeExtractDialect, width, height int64, gaps *nativeConnectorGapSet) *NativeEvaluatedGeometry {
	name, ok := exactNativeAttr(preset, "", "prst")
	if !ok || !nativeConnectorPresetFamily[name] {
		gaps.add("pptx.connector-geometry-unavailable", "only the DrawingML connector preset family is evaluated for connectors; other presets are preserved but not approximated", true)
		return nil
	}
	geometry, err := evaluateNativePresetSource(preset, dialect.drawing, width, height)
	if err != nil {
		gaps.add("pptx.connector-geometry-unavailable", "connector preset geometry is outside the evaluated profile: "+err.Error(), true)
		return nil
	}
	if err := requireOpenStrokedConnectorGeometry(geometry); err != nil {
		gaps.add("pptx.connector-geometry-unavailable", "evaluated connector geometry is not a single open stroked path: "+err.Error(), true)
		return nil
	}
	gaps.add(nativeConnectorPresetPreviewCode, "DrawingML connector preset catalog paths, literal adjustments, and source orientation are evaluated from source; connector geometry remains read-only", false)
	return geometry
}

func requireOpenStrokedConnectorGeometry(geometry *NativeEvaluatedGeometry) error {
	if geometry == nil || len(geometry.Paths) != 1 {
		return fmt.Errorf("expected exactly one path")
	}
	path := geometry.Paths[0]
	if path.FillMode != "none" || !path.Stroke {
		return fmt.Errorf("connector paths must be stroked without fill")
	}
	if len(path.Commands) < 2 || path.Commands[0].Kind != "moveTo" {
		return fmt.Errorf("connector paths start with moveTo and one drawing command")
	}
	for _, command := range path.Commands[1:] {
		switch command.Kind {
		case "lineTo", "cubicBezierTo", "quadBezierTo":
		default:
			return fmt.Errorf("unsupported connector path command %q", command.Kind)
		}
	}
	// Arrowheads take their direction from the terminal segments; a zero-length
	// first or last segment (e.g. an elbow adjustment at 0 or 100000) has no
	// tangent, so it refuses here instead of failing later in a renderer.
	if nativeGeometrySegmentIsDegenerate(path.Commands[0], path.Commands[1]) {
		return fmt.Errorf("first connector segment has zero length")
	}
	last := len(path.Commands) - 1
	if last >= 2 && nativeGeometrySegmentIsDegenerate(path.Commands[last-1], path.Commands[last]) {
		return fmt.Errorf("last connector segment has zero length")
	}
	return nil
}

// nativeGeometrySegmentIsDegenerate reports whether every point of a drawing
// command coincides with the pen position left by the previous command.
func nativeGeometrySegmentIsDegenerate(previous, command NativeGeometryCommand) bool {
	if previous.X == nil || previous.Y == nil {
		return false
	}
	penX, penY := *previous.X, *previous.Y
	for _, point := range [][2]*int64{{command.X, command.Y}, {command.X1, command.Y1}, {command.X2, command.Y2}} {
		if point[0] == nil || point[1] == nil {
			continue
		}
		if *point[0] != penX || *point[1] != penY {
			return false
		}
	}
	return true
}

// resolveNativeConnectorStyle projects a connector's style-matrix outline into
// an owned paint view. Only lnRef selects paint: connector presets carry no
// fillable path, so fillRef is checked for shape but never painted. The local
// <a:ln> inherits the selected matrix outline per ECMA-376 §20.1.4.2.19.
// Raw shape XML and source anchors are never rewritten by this projection.
func resolveNativeConnectorStyle(properties, style, themeRoot *nativeXMLNode, dialect nativeExtractDialect, theme nativeResolvedTheme) (*nativeXMLNode, error) {
	if style == nil {
		return properties, nil
	}
	if requireOnlyNativeAttrs(style) != nil || requireOnlyNativeChildren(style, xml.Name{Space: dialect.drawing, Local: "lnRef"}, xml.Name{Space: dialect.drawing, Local: "fillRef"}, xml.Name{Space: dialect.drawing, Local: "effectRef"}, xml.Name{Space: dialect.drawing, Local: "fontRef"}) != nil {
		return nil, fmt.Errorf("unmodeled connector style markup")
	}
	refs := map[string]*nativeXMLNode{}
	for _, name := range []string{"lnRef", "fillRef", "effectRef", "fontRef"} {
		ref, err := nativeSingleton(style, dialect.drawing, name, true)
		if err != nil {
			return nil, err
		}
		if requireOnlyNativeAttrs(ref, xml.Name{Local: "idx"}) != nil || len(ref.Children) > 1 || !onlyNativeXMLSpace(ref.Text) {
			return nil, fmt.Errorf("unmodeled style reference markup")
		}
		refs[name] = ref
	}
	if index, _ := exactNativeAttr(refs["fontRef"], "", "idx"); index != "major" && index != "minor" {
		return nil, fmt.Errorf("unsupported font style reference")
	}
	if index, _ := exactNativeAttr(refs["effectRef"], "", "idx"); index != "0" {
		return nil, fmt.Errorf("nonzero effect matrix reference remains unsupported")
	}
	if value, _ := exactNativeAttr(refs["fillRef"], "", "idx"); value != "0" {
		if _, err := parseCanonicalNativeInt(value, 1, 3); err != nil {
			return nil, fmt.Errorf("fill matrix index outside bounded entries")
		}
	}
	lineIndexValue, _ := exactNativeAttr(refs["lnRef"], "", "idx")
	if lineIndexValue == "0" {
		return properties, nil
	}
	lineIndex, err := parseCanonicalNativeInt(lineIndexValue, 1, 3)
	if err != nil {
		return nil, fmt.Errorf("line matrix index outside bounded first three entries")
	}
	color, err := exactNativeSolidColor(&nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "solidFill"}, Children: refs["lnRef"].Children, Text: refs["lnRef"].Text}, dialect, theme)
	if err != nil {
		return nil, err
	}
	elements, err := nativeSingleton(themeRoot, dialect.drawing, "themeElements", true)
	if err != nil {
		return nil, err
	}
	matrix, err := nativeSingleton(elements, dialect.drawing, "fmtScheme", true)
	if err != nil {
		return nil, err
	}
	list, err := nativeSingleton(matrix, dialect.drawing, "lnStyleLst", true)
	if err != nil {
		return nil, err
	}
	if requireOnlyNativeAttrs(list) != nil || !onlyNativeXMLSpace(list.Text) || lineIndex > int64(len(list.Children)) {
		return nil, fmt.Errorf("line style matrix entry unavailable")
	}
	selected := list.Children[lineIndex-1]
	if selected.Name != (xml.Name{Space: dialect.drawing, Local: "ln"}) {
		return nil, fmt.Errorf("only exact line matrix entries are supported")
	}
	inherited, err := nativeStylePlaceholderColor(selected, dialect, color)
	if err != nil {
		return nil, err
	}
	// The selected matrix entry must itself be exact before a local override
	// can conceal malformed or duplicate theme markup.
	matrixGaps := nativeShapeGapSet{}
	if _, err := validateNativeAutoShapeLine(&nativeXMLNode{Children: []*nativeXMLNode{inherited}}, dialect, theme, false, &matrixGaps); err != nil {
		return nil, err
	}
	if len(matrixGaps.values) != 0 {
		return nil, fmt.Errorf("unmodeled selected line matrix entry")
	}
	local, err := nativeSingleton(properties, dialect.drawing, "ln", false)
	if err != nil {
		return nil, err
	}
	result := *properties
	result.Children = make([]*nativeXMLNode, 0, len(properties.Children)+1)
	for _, child := range properties.Children {
		if child == local {
			continue
		}
		result.Children = append(result.Children, child)
	}
	result.Children = append(result.Children, mergeNativeInheritedLine(inherited, local, dialect.drawing))
	return &result, nil
}
