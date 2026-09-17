package docxpatch

import (
	"encoding/xml"
	"strings"
)

// Markup Compatibility and Extensibility (ECMA-376 Part 3, clause 10).
//
// mc:AlternateContent is not content in its own right: it is a selector. It
// holds one or more mc:Choice elements, each naming in its Requires attribute
// the namespace prefixes a consumer must understand to take that branch, plus
// an optional trailing mc:Fallback. A consumer takes the first mc:Choice whose
// required namespaces it all understands, and otherwise the mc:Fallback.
//
// Two properties this file insists on, because both are easy to get wrong:
//
//   - Requires holds namespace PREFIXES, not namespace names. A prefix means
//     only what the xmlns declarations in scope say it means, so every token is
//     resolved against those declarations before it is compared. Matching the
//     literal text "wps" would let a package bind that prefix to any namespace
//     at all and still have its branch selected and unwrapped.
//   - A selected branch that contributes no element, and an alternate with no
//     branch to select, are omissions. They are reported so the caller can
//     disclose dropped source content instead of painting a silent gap.
//
// Nothing here changes source bytes, native extraction or editing authority.

const nativeMCMaxAlternateDepth = 8

const (
	// nativeMCSelectedChoice reports the branch of an understood mc:Choice.
	nativeMCSelectedChoice = "choice"
	// nativeMCSelectedFallback reports the mc:Fallback branch, taken when no
	// mc:Choice named a namespace the consumer understands.
	nativeMCSelectedFallback = "fallback"
	// nativeMCUnselected reports an alternate with neither an understood
	// mc:Choice nor an mc:Fallback: its content is dropped.
	nativeMCUnselected = "unselected"
	// nativeMCInvalid reports an alternate outside the Part 3 content model.
	nativeMCInvalid = "invalid"
)

// Omission reasons a caller reports verbatim, so the disclosure names markup
// compatibility rather than the missing element it happened to notice.
const (
	nativeMCReasonInvalid     = "markup-compatibility-invalid"
	nativeMCReasonUnselected  = "markup-compatibility-unselected"
	nativeMCReasonEmptyBranch = "markup-compatibility-empty-branch"
	nativeMCReasonTooDeep     = "markup-compatibility-nesting-too-deep"
)

func nativeMCIsAlternate(node *nativeXMLNode) bool {
	return node != nil && node.Name == (xml.Name{Space: nativeMarkupCompatibilityNS, Local: "AlternateContent"})
}

// nativeMCPrefixNamespace resolves an XML namespace prefix against the xmlns
// declarations in scope at node, innermost first. An unbound prefix, or one
// bound to the empty string, resolves to nothing.
func nativeMCPrefixNamespace(node *nativeXMLNode, prefix string) (string, bool) {
	if prefix == "" || prefix == "xmlns" {
		return "", false
	}
	for current := node; current != nil; current = current.parent {
		for _, attr := range current.Attrs {
			if attr.Name == (xml.Name{Space: "xmlns", Local: prefix}) {
				if attr.Value == "" {
					return "", false
				}
				return attr.Value, true
			}
		}
	}
	return "", false
}

// nativeMCUnderstandsRequires reports whether every prefix the Requires
// attribute names resolves, in scope, to a namespace the consumer understands.
// An absent or empty Requires is not a valid Choice and is never understood.
func nativeMCUnderstandsRequires(choice *nativeXMLNode, understood map[string]bool) bool {
	requires, present := nativeUnqualifiedAttr(choice, "Requires")
	if !present {
		return false
	}
	prefixes := strings.Fields(requires)
	if len(prefixes) == 0 {
		return false
	}
	for _, prefix := range prefixes {
		namespace, bound := nativeMCPrefixNamespace(choice, prefix)
		if !bound || !understood[namespace] {
			return false
		}
	}
	return true
}

// nativeMCOnlyNamespaceAttrs reports whether node carries nothing but xmlns
// declarations and the allowed unqualified attributes.
func nativeMCOnlyNamespaceAttrs(node *nativeXMLNode, allowed ...string) bool {
	for _, attr := range node.Attrs {
		if attr.Name.Space == "xmlns" || (attr.Name.Space == "" && attr.Name.Local == "xmlns") {
			continue
		}
		if attr.Name.Space != "" {
			return false
		}
		if !nativeMCContains(allowed, attr.Name.Local) {
			return false
		}
	}
	return true
}

func nativeMCContains(values []string, want string) bool {
	for _, value := range values {
		if value == want {
			return true
		}
	}
	return false
}

// nativeMCSelectAlternate applies the Part 3 selection rule to one
// mc:AlternateContent and returns the selected branch with the outcome. The
// alternate's content model is Choice+ Fallback?; anything else is invalid and
// is never guessed at.
func nativeMCSelectAlternate(alternate *nativeXMLNode, understood map[string]bool) (*nativeXMLNode, string) {
	if !nativeMCIsAlternate(alternate) || !nativeMCOnlyNamespaceAttrs(alternate) {
		return nil, nativeMCInvalid
	}
	choices, fallback := []*nativeXMLNode{}, (*nativeXMLNode)(nil)
	for _, child := range alternate.Children {
		switch child.Name {
		case xml.Name{Space: nativeMarkupCompatibilityNS, Local: "Choice"}:
			if fallback != nil || !nativeMCOnlyNamespaceAttrs(child, "Requires") {
				return nil, nativeMCInvalid
			}
			if _, present := nativeUnqualifiedAttr(child, "Requires"); !present {
				return nil, nativeMCInvalid
			}
			choices = append(choices, child)
		case xml.Name{Space: nativeMarkupCompatibilityNS, Local: "Fallback"}:
			if fallback != nil || !nativeMCOnlyNamespaceAttrs(child) {
				return nil, nativeMCInvalid
			}
			fallback = child
		default:
			return nil, nativeMCInvalid
		}
	}
	if len(choices) == 0 {
		return nil, nativeMCInvalid
	}
	for _, choice := range choices {
		if nativeMCUnderstandsRequires(choice, understood) {
			return choice, nativeMCSelectedChoice
		}
	}
	if fallback != nil {
		return fallback, nativeMCSelectedFallback
	}
	return nil, nativeMCUnselected
}

// nativeMCResolvedChildren returns node's element children in document order
// with every mc:AlternateContent replaced, in place, by the children of the
// branch the selection rule picks. Alternates nested inside a selected branch
// are resolved the same way, to a bounded depth.
//
// The returned reason is "" only when every alternate contributed its branch's
// content. Otherwise it names the first alternate that contributed nothing, so
// the caller discloses the omission rather than painting a gap.
func nativeMCResolvedChildren(node *nativeXMLNode, understood map[string]bool) ([]*nativeXMLNode, string) {
	return nativeMCResolve(node.Children, understood, 0)
}

func nativeMCResolve(children []*nativeXMLNode, understood map[string]bool, depth int) ([]*nativeXMLNode, string) {
	resolved, reason := make([]*nativeXMLNode, 0, len(children)), ""
	note := func(value string) {
		if reason == "" {
			reason = value
		}
	}
	for _, child := range children {
		if !nativeMCIsAlternate(child) {
			resolved = append(resolved, child)
			continue
		}
		if depth >= nativeMCMaxAlternateDepth {
			note(nativeMCReasonTooDeep)
			continue
		}
		branch, outcome := nativeMCSelectAlternate(child, understood)
		switch outcome {
		case nativeMCInvalid:
			note(nativeMCReasonInvalid)
			continue
		case nativeMCUnselected:
			note(nativeMCReasonUnselected)
			continue
		}
		// An empty Choice or Fallback is a real instruction to drop content.
		// It is disclosed, never treated as "nothing to draw".
		inner, innerReason := nativeMCResolve(branch.Children, understood, depth+1)
		if innerReason != "" {
			note(innerReason)
		}
		if len(inner) == 0 {
			note(nativeMCReasonEmptyBranch)
			continue
		}
		resolved = append(resolved, inner...)
	}
	return resolved, reason
}

// nativeMCFirstChild returns the first resolved child with the given expanded
// name, mirroring firstDirectNativeChild over a resolved child list.
func nativeMCFirstChild(children []*nativeXMLNode, namespace, local string) *nativeXMLNode {
	for _, child := range children {
		if child.Name.Space == namespace && child.Name.Local == local {
			return child
		}
	}
	return nil
}
