package pptxpatch

import "encoding/xml"

// nativeLineChoiceGroups lists the mutually exclusive CT_LineProperties
// children (ECMA-376 Part 1 §20.1.2.2.24). A local <a:ln> that overrides a
// theme style-matrix outline replaces the inherited member of each group as a
// whole; it never blends two fills, dashes, or joins.
var nativeLineChoiceGroups = [][]string{
	{"noFill", "solidFill", "gradFill", "pattFill"},
	{"prstDash", "custDash"},
	{"round", "bevel", "miter"},
	{"headEnd"},
	{"tailEnd"},
	{"extLst"},
}

// mergeNativeInheritedLine projects the effective outline of a local <a:ln>
// over the theme matrix outline selected by a style lnRef: local attributes
// override inherited attributes and each local choice-group child replaces the
// inherited group. Unknown local children are appended unchanged so the exact
// outline validator still sees and refuses them. Neither input is modified and
// the result aliases neither slice. Every value comes from the source package;
// nothing is defaulted or invented.
func mergeNativeInheritedLine(base, local *nativeXMLNode, ns string) *nativeXMLNode {
	if base == nil {
		return local
	}
	if local == nil {
		return base
	}
	result := &nativeXMLNode{Name: local.Name, Text: base.Text, RawStart: local.RawStart, RawEnd: local.RawEnd}
	if !onlyNativeXMLSpace(local.Text) {
		result.Text = local.Text
	}
	result.Attrs = append([]xml.Attr(nil), base.Attrs...)
	for _, attr := range local.Attrs {
		replaced := false
		for index := range result.Attrs {
			if result.Attrs[index].Name == attr.Name {
				result.Attrs[index] = attr
				replaced = true
				break
			}
		}
		if !replaced {
			result.Attrs = append(result.Attrs, attr)
		}
	}
	groupOf := func(name xml.Name) int {
		if name.Space != ns {
			return -1
		}
		for index, group := range nativeLineChoiceGroups {
			for _, member := range group {
				if member == name.Local {
					return index
				}
			}
		}
		return -1
	}
	result.Children = append([]*nativeXMLNode(nil), base.Children...)
	for _, child := range local.Children {
		if group := groupOf(child.Name); group >= 0 {
			kept := result.Children[:0:0]
			for _, existing := range result.Children {
				if groupOf(existing.Name) != group {
					kept = append(kept, existing)
				}
			}
			result.Children = kept
		}
		result.Children = append(result.Children, child)
	}
	return result
}
