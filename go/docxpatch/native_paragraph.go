package docxpatch

import (
	"encoding/json"
	"fmt"
	"strconv"
	"strings"
)

var nativeParagraphLayoutAttrs = map[string][2]string{
	"spacing_before_twips": {"spacing", "before"}, "spacing_after_twips": {"spacing", "after"},
	"line_spacing": {"spacing", "line"}, "line_rule": {"spacing", "lineRule"},
	"indent_left_twips": {"ind", "left"}, "indent_right_twips": {"ind", "right"},
	"first_line_twips": {"ind", "firstLine"}, "hanging_twips": {"ind", "hanging"},
}

func decodeNativeParagraphLayout(members map[string]json.RawMessage) (*NativeDOCXParagraphPropertyPatchV1, bool, error) {
	found := false
	for key := range members {
		if _, ok := nativeParagraphLayoutAttrs[key]; ok {
			found = true
		}
	}
	if !found {
		return nil, false, nil
	}
	patch := &NativeDOCXParagraphPropertyPatchV1{Layout: map[string]*string{}}
	for key, raw := range members {
		if key == "alignment" {
			value, err := decodeNativeMutationJSONString(raw)
			if err != nil {
				return nil, true, err
			}
			patch.Alignment = &value
			continue
		}
		if _, ok := nativeParagraphLayoutAttrs[key]; !ok {
			return nil, true, fmt.Errorf("unknown paragraph property %q", key)
		}
		if string(raw) == "null" {
			patch.Layout[key] = nil
			continue
		}
		value := string(raw)
		if key == "line_rule" {
			var err error
			value, err = decodeNativeMutationJSONString(raw)
			if err != nil {
				return nil, true, err
			}
		} else if _, ok := nativeJSONInt(raw); !ok {
			return nil, true, fmt.Errorf("%s must be a whole number or null", key)
		}
		patch.Layout[key] = &value
	}
	return patch, true, validateNativeParagraphLayout(patch)
}

func validateNativeParagraphLayout(patch *NativeDOCXParagraphPropertyPatchV1) error {
	if patch.Alignment != nil && !nativeAlignmentValues[*patch.Alignment] {
		return fmt.Errorf("invalid alignment")
	}
	for key, value := range patch.Layout {
		if _, ok := nativeParagraphLayoutAttrs[key]; !ok {
			return fmt.Errorf("unknown paragraph property %q", key)
		}
		if value == nil {
			continue
		}
		if key == "line_rule" {
			if *value != "auto" && *value != "exact" && *value != "atLeast" {
				return fmt.Errorf("line_rule must be auto, exact or atLeast")
			}
			continue
		}
		n, err := strconv.Atoi(*value)
		minimum := 0
		if strings.HasPrefix(key, "indent_") {
			minimum = -31680
		}
		if err != nil || n < minimum || n > 31680 {
			return fmt.Errorf("%s must be %d..31680", key, minimum)
		}
	}
	if patch.Layout["first_line_twips"] != nil && patch.Layout["hanging_twips"] != nil {
		return fmt.Errorf("first line and hanging indentation are mutually exclusive")
	}
	return nil
}

func nativeWritableParagraphIndent(node *nativeXMLNode, ns string) bool {
	if !nativeExactResolvedParagraphIndent(node, ns) {
		return false
	}
	for _, attr := range node.Attrs {
		if attr.Name.Local != "left" && attr.Name.Local != "right" && attr.Name.Local != "firstLine" && attr.Name.Local != "hanging" {
			return false
		}
	}
	return true
}

func nativeExtractParagraphLayout(p *NativeParagraphPropertiesV1, node *nativeXMLNode, ns string) {
	for key, pair := range nativeParagraphLayoutAttrs {
		if pair[0] != node.Name.Local {
			continue
		}
		value, ok := nativeAttr(node, ns, pair[1])
		if !ok {
			continue
		}
		if key == "line_rule" {
			p.LineRule = &value
			continue
		}
		n, err := strconv.ParseInt(value, 10, 64)
		if err != nil {
			continue
		}
		switch key {
		case "spacing_before_twips":
			p.SpacingBeforeTwips = &n
		case "spacing_after_twips":
			p.SpacingAfterTwips = &n
		case "line_spacing":
			p.LineSpacing = &n
		case "indent_left_twips":
			p.IndentLeftTwips = &n
		case "indent_right_twips":
			p.IndentRightTwips = &n
		case "first_line_twips":
			p.FirstLineTwips = &n
		case "hanging_twips":
			p.HangingTwips = &n
		}
	}
}

func nativeParagraphLayoutSplice(part []byte, paragraph *nativeXMLNode, patch *NativeDOCXParagraphPropertyPatchV1) (nativeTextSplice, error) {
	written := map[string]map[string]*string{}
	for key, value := range patch.Layout {
		pair := nativeParagraphLayoutAttrs[key]
		if written[pair[0]] == nil {
			written[pair[0]] = map[string]*string{}
		}
		written[pair[0]][pair[1]] = value
	}
	if patch.Layout["first_line_twips"] != nil {
		written["ind"]["hanging"] = nil
	}
	if patch.Layout["hanging_twips"] != nil {
		written["ind"]["firstLine"] = nil
	}
	if patch.Alignment != nil {
		written["jc"] = map[string]*string{"val": patch.Alignment}
	}
	return nativeMergePropertyContainer(part, paragraph, "pPr", nativeParagraphPropertyOrder, written)
}

func proveNativeParagraphLayout(properties *NativeParagraphPropertiesV1, patch *NativeDOCXParagraphPropertyPatchV1) error {
	raw, _ := json.Marshal(properties)
	values := map[string]json.RawMessage{}
	_ = json.Unmarshal(raw, &values)
	for key, want := range patch.Layout {
		got, ok := values[key]
		if want == nil {
			if ok {
				return fmt.Errorf("%s was not cleared", key)
			}
			continue
		}
		expected := *want
		if key == "line_rule" {
			encoded, _ := json.Marshal(expected)
			expected = string(encoded)
		}
		if string(got) != expected {
			return fmt.Errorf("%s did not round-trip", key)
		}
	}
	if patch.Alignment != nil && (properties.Alignment == nil || *properties.Alignment != *patch.Alignment) {
		return fmt.Errorf("alignment did not round-trip")
	}
	return nil
}
