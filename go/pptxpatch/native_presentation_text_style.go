package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

const nativePresentationTextStylePreviewCode = "pptx.presentation-text-style-preview"

const nativePresentationTextStylePreviewMessage = "explicit presentation level styles are projected before local level, paragraph and run properties; target remains read-only"

// extractNativeShapeParagraphs projects the explicit matching levels of the
// presentation defaultTextStyle beneath a non-placeholder shape's local list
// style before the exact paragraph extractor runs. Only explicit matching
// level styles are qualified here. DefPPr fallback semantics differ in
// PowerPoint and remain refused, not merged by guesswork. Table cells keep
// their separate table-style path; no font defaults are invented. The
// projection is an owned view; raw source nodes and offsets stay untouched.
func (extractor *nativeExtractor) extractNativeShapeParagraphs(body *nativeXMLNode, dialect nativeExtractDialect) ([]NativeParagraph, error) {
	defaults := extractor.presentationTextPreviewStyle
	if defaults == nil || body == nil {
		return extractor.extractNativeParagraphs(body, dialect)
	}
	local, err := nativeSingleton(body, dialect.drawing, "lstStyle", true)
	if err != nil {
		return nil, err
	}
	levels := func(node *nativeXMLNode) (map[string]*nativeXMLNode, error) {
		if requireOnlyNativeAttrs(node) != nil || !onlyNativeXMLSpace(node.Text) {
			return nil, unsupportedNativeTextContent("unmodeled presentation/list text style metadata")
		}
		result := map[string]*nativeXMLNode{}
		for _, child := range node.Children {
			name := child.Name.Local
			if child.Name.Space != dialect.drawing || (name != "defPPr" && (len(name) != 7 || name[:3] != "lvl" || name[3] < '1' || name[3] > '9' || name[4:] != "pPr")) {
				return nil, unsupportedNativeTextContent("unsupported presentation text style level")
			}
			if result[name] != nil {
				return nil, nativeDuplicateSingletonError{space: dialect.drawing, local: name}
			}
			if err := validateNativeTextStyleProperties(child, dialect, true, extractor.theme); err != nil {
				return nil, err
			}
			result[name] = child
		}
		return result, nil
	}
	base, err := levels(defaults)
	if err != nil {
		return nil, err
	}
	override, err := levels(local)
	if err != nil {
		return nil, err
	}
	if base["defPPr"] != nil || override["defPPr"] != nil {
		return nil, unsupportedNativeTextContent("defPPr fallback alongside presentation levels is not qualified")
	}
	merged := &nativeXMLNode{Name: local.Name}
	for level := 1; level <= 9; level++ {
		name := fmt.Sprintf("lvl%dpPr", level)
		if base[name] == nil && override[name] == nil {
			continue
		}
		properties := mergeNativeStyleNodes(base[name], override[name], dialect)
		properties.Name = xml.Name{Space: dialect.drawing, Local: name}
		merged.Children = append(merged.Children, properties)
	}
	projected := *body
	projected.Children = append([]*nativeXMLNode(nil), body.Children...)
	for i, child := range projected.Children {
		if child == local {
			projected.Children[i] = merged
		}
	}
	return extractor.extractNativeParagraphs(&projected, dialect)
}

func nativeMarkPresentationTextStylePreview(element *NativeElement) {
	element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: nativePresentationTextStylePreviewCode, Message: nativePresentationTextStylePreviewMessage})
}
