package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"sort"
	"strings"
)

const nativePlaceholderPreviewCode = "pptx.placeholder-inheritance-approximate"

// nativePlaceholderPreviewHairlineWidthEmu paints an inherited outline that
// declares no width (DrawingML "thinnest possible" line) as one 96 DPI pixel.
const nativePlaceholderPreviewHairlineWidthEmu = 9525

// nativePlaceholderPreview is the read-only projection of a slide placeholder
// whose chain the exact title/body resolver did not qualify. It exists only
// behind AllowInheritedTextPreview and is disclosed per element.
type nativePlaceholderPreview struct {
	kind        NativePlaceholderType
	sourceKind  string
	layers      []*nativeXMLNode
	omitted     map[string]bool
	hasTextBody bool
	// Effective frame paint resolved master → layout → slide, nearest wins.
	// Both stay nil when the chain declares no paint or only unmodeled paint.
	fill     *string
	stroke   *NativeStroke
	hairline bool
}

func (preview *nativePlaceholderPreview) omit(name string) {
	if preview.omitted == nil {
		preview.omitted = map[string]bool{}
	}
	if len(preview.omitted) < nativeMaxInheritedOmissionNames {
		preview.omitted[name] = true
	}
}

func (preview *nativePlaceholderPreview) omissions() []string {
	names := make([]string, 0, len(preview.omitted))
	for name := range preview.omitted {
		names = append(names, name)
	}
	sort.Strings(names)
	return names
}

// nativePlaceholderFamily maps the ST_PlaceholderType values whose text
// inherits master txStyles to their title/body family. Every other kind
// (dates, footers, slide numbers, pictures, charts, tables, media, diagrams)
// stays outside this preview.
func nativePlaceholderFamily(kind string) (NativePlaceholderType, string, bool) {
	switch kind {
	case "title":
		return NativePlaceholderTypeTitle, "title", true
	case "ctrTitle":
		return NativePlaceholderTypeCtrTitle, "title", true
	case "subTitle":
		return NativePlaceholderTypeSubTitle, "body", true
	case "", "body", "obj":
		return NativePlaceholderTypeBody, "body", true
	}
	return "", "", false
}

// resolveNativePlaceholderPreview resolves layout-by-index then master-by-
// family for title, ctrTitle, subTitle, body and obj placeholders. Ancestor
// geometry, body properties and solid frame paint cascade; ancestor list styles
// and the master title/body style become preview layers; ancestor prompt text
// is never painted and is disclosed. The returned node is an owned view
// retaining the raw source offsets of the slide shape.
func (extractor *nativeExtractor) resolveNativePlaceholderPreview(node *nativeXMLNode, dialect nativeExtractDialect) (*nativeXMLNode, *nativePlaceholderPreview, error) {
	identity, err := nativePlaceholderMetadata(node, dialect)
	if identity == nil || err != nil {
		return node, nil, err
	}
	kind, family, ok := nativePlaceholderFamily(identity.kind)
	if !ok {
		return nil, nil, unsupportedNativePlaceholder("placeholder type " + identity.kind + " is outside the read-only title/body inheritance families")
	}
	// Same bottom layer as the non-placeholder preview: presentation
	// defaultTextStyle, then the master title/body style and placeholder chain.
	preview := &nativePlaceholderPreview{kind: kind, sourceKind: identity.kind, layers: []*nativeXMLNode{extractor.presentationTextPreviewStyle}}
	layout, layoutIdentity, err := nativeMatchingPlaceholder(extractor.slideDependencies.layoutRoot, *identity, true, dialect)
	if err != nil && identity.kind != "" {
		layout, layoutIdentity, err = nativeMatchingPlaceholder(extractor.slideDependencies.layoutRoot, *identity, false, dialect)
	}
	if err != nil {
		return nil, nil, unsupportedNativePlaceholder("placeholder has no unambiguous layout match: " + err.Error())
	}
	_, layoutFamily, ok := nativePlaceholderFamily(layoutIdentity.kind)
	if !ok || layoutFamily != family {
		return nil, nil, unsupportedNativePlaceholder("placeholder type conflicts with its layout placeholder family")
	}
	if identity.kind == "" {
		preview.kind, _, _ = nativePlaceholderFamily(layoutIdentity.kind)
		preview.sourceKind = layoutIdentity.kind
	}
	var master *nativeXMLNode
	if extractor.slideDependencies.masterRoot != nil {
		master, _, err = nativeMatchingPlaceholder(extractor.slideDependencies.masterRoot, nativePlaceholderIdentity{kind: family}, false, dialect)
		if err != nil {
			master = nil
			preview.omit("master placeholder (" + family + ") unmatched")
		}
		styles, err := nativeSingleton(extractor.slideDependencies.masterRoot, dialect.presentation, "txStyles", false)
		if err != nil {
			return nil, nil, err
		}
		if styles != nil {
			if requireOnlyNativeAttrs(styles) != nil || requireOnlyNativeChildren(styles, xml.Name{Space: dialect.presentation, Local: "titleStyle"}, xml.Name{Space: dialect.presentation, Local: "bodyStyle"}, xml.Name{Space: dialect.presentation, Local: "otherStyle"}) != nil {
				return nil, nil, unsupportedNativePlaceholder("unmodeled master text styles")
			}
			style, err := nativeSingleton(styles, dialect.presentation, family+"Style", false)
			if err != nil {
				return nil, nil, err
			}
			preview.layers = append(preview.layers, style)
		}
	}
	var transform, bodyProperties *nativeXMLNode
	for _, shape := range []*nativeXMLNode{master, layout, node} {
		if shape == nil {
			continue
		}
		ancestor := shape != node
		if requireOnlyNativeAttrs(shape) != nil || requireOnlyNativeChildren(shape, xml.Name{Space: dialect.presentation, Local: "nvSpPr"}, xml.Name{Space: dialect.presentation, Local: "spPr"}, xml.Name{Space: dialect.presentation, Local: "style"}, xml.Name{Space: dialect.presentation, Local: "txBody"}) != nil || !onlyNativeXMLSpace(shape.Text) {
			return nil, nil, unsupportedNativePlaceholder("unmodeled placeholder shape markup")
		}
		if err := extractor.validateNativePlaceholderPreviewNonVisual(shape, ancestor, dialect, preview); err != nil {
			return nil, nil, err
		}
		if style, err := nativeSingleton(shape, dialect.presentation, "style", false); err != nil {
			return nil, nil, err
		} else if style != nil {
			if !ancestor {
				return nil, nil, unsupportedNativePlaceholder("placeholder shape style references remain outside the inheritance preview")
			}
			preview.omit("ancestor p:style")
		}
		properties, err := nativeSingleton(shape, dialect.presentation, "spPr", true)
		if err != nil {
			return nil, nil, err
		}
		localTransform, err := extractor.validateNativePlaceholderPreviewShapeProperties(properties, ancestor, dialect, preview)
		if err != nil {
			return nil, nil, err
		}
		if localTransform != nil {
			transform = localTransform
		}
		body, err := nativeSingleton(shape, dialect.presentation, "txBody", false)
		if err != nil {
			return nil, nil, err
		}
		if body == nil {
			continue
		}
		if requireOnlyNativeAttrs(body) != nil || requireOnlyNativeChildren(body, xml.Name{Space: dialect.drawing, Local: "bodyPr"}, xml.Name{Space: dialect.drawing, Local: "lstStyle"}, xml.Name{Space: dialect.drawing, Local: "p"}) != nil || !onlyNativeXMLSpace(body.Text) {
			return nil, nil, unsupportedNativePlaceholder("unmodeled placeholder text body")
		}
		localBody, err := nativeSingleton(body, dialect.drawing, "bodyPr", false)
		if err != nil {
			return nil, nil, err
		}
		if localBody != nil {
			if err := validateNativePlaceholderPreviewBodyProperties(localBody, dialect); err != nil {
				return nil, nil, err
			}
			bodyProperties = mergeNativePlaceholderBodyProperties(bodyProperties, localBody, dialect)
		}
		localList, err := nativeSingleton(body, dialect.drawing, "lstStyle", false)
		if err != nil {
			return nil, nil, err
		}
		if ancestor {
			preview.layers = append(preview.layers, localList)
			for _, paragraph := range body.Children {
				if paragraph.Name == (xml.Name{Space: dialect.drawing, Local: "p"}) && (len(paragraph.Children) != 0 || len(paragraph.Attrs) != 0 || !onlyNativeXMLSpace(paragraph.Text)) {
					preview.omit("ancestor prompt paragraphs")
					break
				}
			}
		} else {
			preview.hasTextBody = true
		}
	}
	if transform == nil {
		return nil, nil, unsupportedNativePlaceholder("placeholder geometry remains unresolved through the layout/master chain")
	}
	if bodyProperties == nil {
		bodyProperties = &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "bodyPr"}}
	}
	result := *node
	result.Children = make([]*nativeXMLNode, 0, len(node.Children)+1)
	projectedBody := false
	for _, child := range node.Children {
		switch child.Name {
		case xml.Name{Space: dialect.presentation, Local: "spPr"}:
			projected := *child
			projected.Attrs = nil
			projected.Children = []*nativeXMLNode{transform}
			result.Children = append(result.Children, &projected)
		case xml.Name{Space: dialect.presentation, Local: "txBody"}:
			projected := *child
			projected.Children = []*nativeXMLNode{bodyProperties}
			list, err := nativeSingleton(child, dialect.drawing, "lstStyle", false)
			if err != nil {
				return nil, nil, err
			}
			if list == nil {
				list = &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "lstStyle"}}
			}
			projected.Children = append(projected.Children, list)
			for _, item := range child.Children {
				if item.Name != (xml.Name{Space: dialect.drawing, Local: "bodyPr"}) && item.Name != (xml.Name{Space: dialect.drawing, Local: "lstStyle"}) {
					projected.Children = append(projected.Children, item)
				}
			}
			result.Children = append(result.Children, &projected)
			projectedBody = true
		case xml.Name{Space: dialect.presentation, Local: "nvSpPr"}:
			result.Children = append(result.Children, nativeProjectPlaceholderPreviewNonVisual(child, dialect))
		default:
			result.Children = append(result.Children, child)
		}
	}
	if !projectedBody {
		// PowerPoint paints no prompt for an empty slide placeholder; the owned
		// view carries only the inherited frame so nothing is invented.
		result.Children = append(result.Children, &nativeXMLNode{Name: xml.Name{Space: dialect.presentation, Local: "txBody"}, Children: []*nativeXMLNode{bodyProperties, {Name: xml.Name{Space: dialect.drawing, Local: "lstStyle"}}}})
	}
	return &result, preview, nil
}

func (extractor *nativeExtractor) validateNativePlaceholderPreviewNonVisual(shape *nativeXMLNode, ancestor bool, dialect nativeExtractDialect, preview *nativePlaceholderPreview) error {
	nv, err := nativeSingleton(shape, dialect.presentation, "nvSpPr", true)
	if err != nil {
		return err
	}
	if requireOnlyNativeAttrs(nv) != nil || requireOnlyNativeChildren(nv, xml.Name{Space: dialect.presentation, Local: "cNvPr"}, xml.Name{Space: dialect.presentation, Local: "cNvSpPr"}, xml.Name{Space: dialect.presentation, Local: "nvPr"}) != nil || !onlyNativeXMLSpace(nv.Text) {
		return unsupportedNativePlaceholder("unmodeled placeholder nonvisual metadata")
	}
	cNvPr, err := nativeSingleton(nv, dialect.presentation, "cNvPr", true)
	if err != nil {
		return err
	}
	if requireOnlyNativeAttrs(cNvPr, xml.Name{Local: "id"}, xml.Name{Local: "name"}, xml.Name{Local: "descr"}, xml.Name{Local: "title"}, xml.Name{Local: "hidden"}) != nil || requireOnlyNativeChildren(cNvPr, xml.Name{Space: dialect.drawing, Local: "extLst"}) != nil || !onlyNativeXMLSpace(cNvPr.Text) {
		return unsupportedNativePlaceholder("unmodeled placeholder identity metadata")
	}
	if value, ok := exactNativeAttr(cNvPr, "", "hidden"); ok {
		hidden, err := nativeBool(value)
		if err != nil {
			return err
		}
		if hidden && !ancestor {
			return unsupportedNativePlaceholder("hidden placeholder is not painted")
		}
	}
	for _, name := range []string{"descr", "title"} {
		if _, ok := exactNativeAttr(cNvPr, "", name); ok && !ancestor {
			preview.omit("p:cNvPr@" + name)
		}
	}
	if len(cNvPr.Children) != 0 && !ancestor {
		preview.omit("p:cNvPr/a:extLst")
	}
	cNvSpPr, err := nativeSingleton(nv, dialect.presentation, "cNvSpPr", true)
	if err != nil {
		return err
	}
	if requireOnlyNativeAttrs(cNvSpPr, xml.Name{Local: "txBox"}) != nil || requireOnlyNativeChildren(cNvSpPr, xml.Name{Space: dialect.drawing, Local: "spLocks"}) != nil || !onlyNativeXMLSpace(cNvSpPr.Text) {
		return unsupportedNativePlaceholder("unmodeled placeholder shape metadata")
	}
	if len(cNvSpPr.Children) != 0 && !ancestor {
		preview.omit("p:cNvSpPr/a:spLocks")
	}
	nvPr, err := nativeSingleton(nv, dialect.presentation, "nvPr", true)
	if err != nil {
		return err
	}
	if requireOnlyNativeAttrs(nvPr) != nil || requireOnlyNativeChildren(nvPr, xml.Name{Space: dialect.presentation, Local: "ph"}) != nil || !onlyNativeXMLSpace(nvPr.Text) {
		return unsupportedNativePlaceholder("unmodeled placeholder nonvisual properties")
	}
	return nil
}

// Solid fill and solid outline resolve through the same chain as geometry:
// each layer's declaration replaces the previous one, so a nearer a:noFill
// paints nothing. Gradient, pattern and picture fills, effects and unmodeled
// outline markup clear the inherited paint and are disclosed on ancestors;
// on the slide placeholder itself they refuse as before.
func (extractor *nativeExtractor) validateNativePlaceholderPreviewShapeProperties(properties *nativeXMLNode, ancestor bool, dialect nativeExtractDialect, preview *nativePlaceholderPreview) (*nativeXMLNode, error) {
	if requireOnlyNativeAttrs(properties, xml.Name{Local: "bwMode"}) != nil || !onlyNativeXMLSpace(properties.Text) {
		return nil, unsupportedNativePlaceholder("unmodeled placeholder shape properties")
	}
	allowed := []xml.Name{{Space: dialect.drawing, Local: "xfrm"}, {Space: dialect.drawing, Local: "prstGeom"}, {Space: dialect.drawing, Local: "noFill"}, {Space: dialect.drawing, Local: "solidFill"}, {Space: dialect.drawing, Local: "ln"}}
	if ancestor {
		allowed = append(allowed, xml.Name{Space: dialect.drawing, Local: "gradFill"}, xml.Name{Space: dialect.drawing, Local: "pattFill"}, xml.Name{Space: dialect.drawing, Local: "blipFill"}, xml.Name{Space: dialect.drawing, Local: "effectLst"}, xml.Name{Space: dialect.drawing, Local: "extLst"})
	}
	if requireOnlyNativeChildren(properties, allowed...) != nil {
		return nil, unsupportedNativePlaceholder("placeholder shape paint is outside the inherited frame preview")
	}
	for _, name := range allowed {
		if _, err := nativeSingleton(properties, name.Space, name.Local, false); err != nil {
			return nil, err
		}
	}
	if _, ok := exactNativeAttr(properties, "", "bwMode"); ok {
		preview.omit("p:spPr@bwMode")
	}
	fills := 0
	for _, child := range properties.Children {
		switch child.Name.Local {
		case "noFill", "solidFill", "gradFill", "pattFill", "blipFill":
			fills++
		}
	}
	if fills > 1 {
		if !ancestor {
			return nil, unsupportedNativePlaceholder("placeholder declares conflicting fills")
		}
		preview.fill = nil
		preview.omit("ancestor conflicting fills")
	}
	var transform *nativeXMLNode
	for _, child := range properties.Children {
		switch child.Name.Local {
		case "xfrm":
			gaps := nativeShapeGapSet{}
			if _, err := validateNativeAutoShapeTransform(child, dialect, &gaps); err != nil {
				return nil, err
			}
			if gaps.refused() || requireOnlyNativeAttrs(child) != nil {
				return nil, unsupportedNativePlaceholder("placeholder transform is outside the exact subset")
			}
			transform = child
		case "prstGeom":
			preset, ok := exactNativeAttr(child, "", "prst")
			if !ok || preset != "rect" || requireOnlyNativeAttrs(child, xml.Name{Local: "prst"}) != nil || requireOnlyNativeChildren(child, xml.Name{Space: dialect.drawing, Local: "avLst"}) != nil || !onlyNativeXMLSpace(child.Text) {
				return nil, unsupportedNativePlaceholder("placeholder geometry must be an unadjusted rectangle")
			}
			adjustments, err := nativeSingleton(child, dialect.drawing, "avLst", false)
			if err != nil {
				return nil, err
			}
			if adjustments != nil && requireEmptyNativeElement(adjustments) != nil {
				return nil, unsupportedNativePlaceholder("placeholder geometry adjustments are unsupported")
			}
		case "noFill", "solidFill", "gradFill", "pattFill", "blipFill":
			if fills > 1 {
				continue
			}
			if err := extractor.resolveNativePlaceholderPreviewFill(child, ancestor, dialect, preview); err != nil {
				return nil, err
			}
		case "ln":
			if err := extractor.resolveNativePlaceholderPreviewOutline(child, ancestor, dialect, preview); err != nil {
				return nil, err
			}
		default:
			preview.omit("ancestor a:" + child.Name.Local)
		}
	}
	return transform, nil
}

// resolveNativePlaceholderPreviewFill applies one layer's fill declaration.
// Only an exact sRGB or documented theme solid color is painted; every other
// declaration clears the inherited fill so a nearer unmodeled paint is never
// replaced by a farther solid one.
func (extractor *nativeExtractor) resolveNativePlaceholderPreviewFill(child *nativeXMLNode, ancestor bool, dialect nativeExtractDialect, preview *nativePlaceholderPreview) error {
	switch child.Name.Local {
	case "noFill":
		if requireEmptyNativeElement(child) != nil {
			return unsupportedNativePlaceholder("placeholder no-fill markup is not exact")
		}
		preview.fill = nil
	case "solidFill":
		color, err := exactNativeSolidColor(child, dialect, extractor.theme)
		if err != nil {
			if !ancestor {
				return unsupportedNativePlaceholder("placeholder fill color is outside the inherited frame preview")
			}
			preview.fill = nil
			preview.omit("ancestor a:solidFill (unresolved color)")
			return nil
		}
		preview.fill = &color
	default:
		preview.fill = nil
		preview.omit("ancestor a:" + child.Name.Local)
	}
	return nil
}

// resolveNativePlaceholderPreviewOutline applies one layer's a:ln. A solid
// sRGB/theme line fill paints with its declared width, or a hairline when the
// source declares none; a:noFill paints nothing; dashes, compound lines, caps,
// joins and unmodeled markup are omitted with disclosure on ancestors and
// refused on the slide placeholder itself.
func (extractor *nativeExtractor) resolveNativePlaceholderPreviewOutline(line *nativeXMLNode, ancestor bool, dialect nativeExtractDialect, preview *nativePlaceholderPreview) error {
	name := "p:spPr/a:ln"
	if ancestor {
		name = "ancestor a:ln"
	}
	unmodeled := func(reason string) error {
		if !ancestor {
			return unsupportedNativePlaceholder("placeholder outline " + reason + " is outside the inherited frame preview")
		}
		preview.stroke, preview.hairline = nil, false
		preview.omit(name + " (" + reason + ")")
		return nil
	}
	if requireOnlyNativeAttrs(line, xml.Name{Local: "w"}, xml.Name{Local: "cap"}, xml.Name{Local: "cmpd"}, xml.Name{Local: "algn"}) != nil || !onlyNativeXMLSpace(line.Text) {
		return unmodeled("attributes")
	}
	children := []xml.Name{{Space: dialect.drawing, Local: "noFill"}, {Space: dialect.drawing, Local: "solidFill"}, {Space: dialect.drawing, Local: "prstDash"}, {Space: dialect.drawing, Local: "round"}, {Space: dialect.drawing, Local: "bevel"}, {Space: dialect.drawing, Local: "miter"}}
	if requireOnlyNativeChildren(line, children...) != nil {
		return unmodeled("markup")
	}
	for _, child := range children {
		if _, err := nativeSingleton(line, child.Space, child.Local, false); err != nil {
			return unmodeled("markup")
		}
	}
	noFill := nativeChild(line, dialect.drawing, "noFill")
	solidFill := nativeChild(line, dialect.drawing, "solidFill")
	if noFill != nil && solidFill != nil {
		return unmodeled("fill")
	}
	if noFill != nil {
		if requireEmptyNativeElement(noFill) != nil {
			return unmodeled("no-fill markup")
		}
		preview.stroke, preview.hairline = nil, false
		return nil
	}
	if solidFill == nil {
		// No line fill declared at this layer; the chain keeps the inherited one.
		return nil
	}
	if dash := nativeChild(line, dialect.drawing, "prstDash"); dash != nil {
		value, ok := exactNativeAttr(dash, "", "val")
		if !ok || value != "solid" || requireOnlyNativeAttrs(dash, xml.Name{Local: "val"}) != nil || requireOnlyNativeChildren(dash) != nil {
			return unmodeled("dash")
		}
	}
	color, err := exactNativeSolidColor(solidFill, dialect, extractor.theme)
	if err != nil {
		return unmodeled("color")
	}
	width, hairline := int64(nativePlaceholderPreviewHairlineWidthEmu), true
	if value, ok := exactNativeAttr(line, "", "w"); ok {
		parsed, err := parseCanonicalNativeInt(value, 0, nativeMaxLineWidthEmu)
		if err != nil {
			return unmodeled("width")
		}
		if parsed > 0 {
			width, hairline = parsed, false
		}
	}
	for _, attr := range []string{"cap", "cmpd", "algn"} {
		if _, ok := exactNativeAttr(line, "", attr); ok {
			preview.omit(name + "@" + attr)
		}
	}
	for _, join := range []string{"round", "bevel", "miter"} {
		if nativeChild(line, dialect.drawing, join) != nil {
			preview.omit(name + "/a:" + join)
		}
	}
	preview.stroke, preview.hairline = &NativeStroke{Color: color, WidthEMU: int64Pointer(width)}, hairline
	return nil
}

// Ancestor body properties are structurally validated here; the merged result
// is validated by the text-body layout policy that owns autofit and columns.
func validateNativePlaceholderPreviewBodyProperties(bodyPr *nativeXMLNode, dialect nativeExtractDialect) error {
	if !onlyNativeXMLSpace(bodyPr.Text) {
		return unsupportedNativePlaceholder("placeholder body properties contain text")
	}
	if err := requireOnlyNativeAttrs(bodyPr,
		xml.Name{Local: "rot"}, xml.Name{Local: "spcFirstLastPara"},
		xml.Name{Local: "vertOverflow"}, xml.Name{Local: "horzOverflow"},
		xml.Name{Local: "vert"}, xml.Name{Local: "wrap"},
		xml.Name{Local: "lIns"}, xml.Name{Local: "tIns"}, xml.Name{Local: "rIns"}, xml.Name{Local: "bIns"},
		xml.Name{Local: "numCol"}, xml.Name{Local: "spcCol"}, xml.Name{Local: "rtlCol"},
		xml.Name{Local: "fromWordArt"}, xml.Name{Local: "anchor"}, xml.Name{Local: "anchorCtr"},
		xml.Name{Local: "forceAA"}, xml.Name{Local: "upright"}, xml.Name{Local: "compatLnSpc"}); err != nil {
		return unsupportedNativePlaceholder("placeholder body properties contain an unsupported attribute")
	}
	if duplicateNativeAttrs(bodyPr.Attrs) {
		return unsupportedNativePlaceholder("placeholder body properties repeat an attribute")
	}
	if err := requireOnlyNativeChildren(bodyPr,
		xml.Name{Space: dialect.drawing, Local: "noAutofit"},
		xml.Name{Space: dialect.drawing, Local: "normAutofit"},
		xml.Name{Space: dialect.drawing, Local: "spAutoFit"}); err != nil {
		return unsupportedNativePlaceholder("placeholder body properties contain an unsupported child")
	}
	if len(bodyPr.Children) > 1 {
		return unsupportedNativePlaceholder("placeholder body properties declare conflicting autofit")
	}
	return nil
}

// mergeNativePlaceholderBodyProperties applies override attributes over base
// attributes and treats the three autofit children as one replaceable slot.
func mergeNativePlaceholderBodyProperties(base, override *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	result := &nativeXMLNode{Name: xml.Name{Space: dialect.drawing, Local: "bodyPr"}}
	if base != nil {
		result.Attrs = append(result.Attrs, base.Attrs...)
		result.Children = append(result.Children, base.Children...)
	}
	for _, attr := range override.Attrs {
		replaced := false
		for i := range result.Attrs {
			if result.Attrs[i].Name == attr.Name {
				result.Attrs[i] = attr
				replaced = true
				break
			}
		}
		if !replaced {
			result.Attrs = append(result.Attrs, attr)
		}
	}
	if len(override.Children) != 0 {
		result.Children = append([]*nativeXMLNode(nil), override.Children...)
	}
	return result
}

// nativeProjectPlaceholderPreviewNonVisual owns a text-box view of the
// nonvisual block: placeholder binding, locks, extensions and accessibility
// text are removed from the view only; the source bytes are untouched.
func nativeProjectPlaceholderPreviewNonVisual(nv *nativeXMLNode, dialect nativeExtractDialect) *nativeXMLNode {
	projected := *nv
	projected.Children = make([]*nativeXMLNode, 0, len(nv.Children))
	for _, properties := range nv.Children {
		switch properties.Name {
		case xml.Name{Space: dialect.presentation, Local: "nvPr"}:
			projected.Children = append(projected.Children, &nativeXMLNode{Name: properties.Name})
		case xml.Name{Space: dialect.presentation, Local: "cNvSpPr"}:
			copy := *properties
			copy.Children = nil
			copy.Attrs = []xml.Attr{{Name: xml.Name{Local: "txBox"}, Value: "1"}}
			projected.Children = append(projected.Children, &copy)
		case xml.Name{Space: dialect.presentation, Local: "cNvPr"}:
			copy := *properties
			copy.Children = nil
			copy.Attrs = nil
			for _, attr := range properties.Attrs {
				if attr.Name == (xml.Name{Local: "id"}) || attr.Name == (xml.Name{Local: "name"}) {
					copy.Attrs = append(copy.Attrs, attr)
				}
			}
			projected.Children = append(projected.Children, &copy)
		default:
			projected.Children = append(projected.Children, properties)
		}
	}
	return &projected
}

func nativeMarkPlaceholderPreview(element *NativeElement, preview *nativePlaceholderPreview, slideID, elementID, part string) {
	if element == nil || preview == nil {
		return
	}
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	family := "body"
	if preview.kind == NativePlaceholderTypeTitle || preview.kind == NativePlaceholderTypeCtrTitle {
		family = "title"
	}
	message := fmt.Sprintf("Read-only approximate placeholder inheritance: %s placeholder resolved through the layout-by-index and master-by-%s-family chain with master %sStyle, ancestor list styles and body properties as declared preview layers.", preview.sourceKind, family, family)
	if !preview.hasTextBody {
		message += " The slide placeholder has no text body; PowerPoint paints no prompt, so no text is invented."
	}
	if preview.fill != nil || preview.stroke != nil {
		// The painted frame reuses the AutoShape fill/stroke contract, so the
		// preview element is a rect preset shape; it stays preserve-only and
		// every placeholder mutation refusal keys off this diagnostic.
		preset := NativeShapePresetRect
		element.Kind, element.Preset, element.Fill, element.Stroke = NativeElementKindShape, &preset, preview.fill, preview.stroke
		var painted []string
		if preview.fill != nil {
			painted = append(painted, "solid fill "+*preview.fill)
		}
		if preview.stroke != nil {
			outline := fmt.Sprintf("solid outline %s (%d EMU", preview.stroke.Color, *preview.stroke.WidthEMU)
			if preview.hairline {
				outline += ", a hairline default because the source declares no width"
			}
			painted = append(painted, outline+")")
		}
		message += " Frame " + strings.Join(painted, " and ") + " inherited from the slide/layout/master placeholder chain (nearest wins) and painted read-only as a rect preset."
		if !nativeParagraphsHaveText(element.Paragraphs) {
			message += " PowerPoint slideshow and export hide placeholders without text, so this painted frame is an approximation."
		}
	}
	if names := preview.omissions(); len(names) != 0 {
		message += " Not painted: " + strings.Join(names, ", ") + "."
	}
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
		Severity: NativeDiagnosticSeverityWarning,
		Code:     nativePlaceholderPreviewCode,
		Message:  message,
		Scope:    &NativeDiagnosticScope{SlideID: &slideID, ElementID: &elementID, PartName: &part},
	})
}

func nativeParagraphsHaveText(paragraphs *[]NativeParagraph) bool {
	if paragraphs == nil {
		return false
	}
	for _, paragraph := range *paragraphs {
		for _, run := range paragraph.Runs {
			if run.Text != nil && strings.TrimSpace(*run.Text) != "" {
				return true
			}
		}
	}
	return false
}
