package docxpatch

import (
	"encoding/xml"
	"strings"
)

// The result is already restricted to one uniformly formatted text run. Thus
// MERGEFORMAT preserves its exact result properties; Arabic selects the same
// decimal digits as the qualified PAGE/NUMPAGES display profile.
func nativePageFieldInstruction(raw string) (string, bool) {
	if len(raw) > 256 {
		return "", false
	}
	tokens := strings.FieldsFunc(raw, func(r rune) bool { return r == ' ' || r == '\t' || r == '\r' || r == '\n' })
	if len(tokens) == 0 || (tokens[0] != "PAGE" && tokens[0] != "NUMPAGES") {
		return "", false
	}
	seen := map[string]bool{}
	for i := 1; i < len(tokens); i++ {
		token := tokens[i]
		if !strings.HasPrefix(token, `\*`) {
			return "", false
		}
		name := strings.TrimPrefix(token, `\*`)
		if name == "" {
			i++
			if i >= len(tokens) {
				return "", false
			}
			name = tokens[i]
		}
		if (name != "MERGEFORMAT" && name != "Arabic") || seen[name] {
			return "", false
		}
		seen[name] = true
	}
	return tokens[0], true
}

func hasNativeFieldBegin(run *nativeXMLNode, wordNS string) bool {
	for _, child := range directNativeChildren(run, wordNS, "fldChar") {
		if value, _ := nativeAttr(child, wordNS, "fldCharType"); value == "begin" {
			return true
		}
	}
	return false
}

// Only one paragraph-local field with a single result run. Instruction and
// boundary XML remain bound by the paragraph/story source digests; only the
// result run enters the native text model and its cached text is discarded.
func (extractor *nativeExtractor) extractFlatPageField(partName, paragraphID string, sequence []*nativeXMLNode) (NativeRunV1, bool, error) {
	refuse := func() (NativeRunV1, bool, error) { return NativeRunV1{}, false, nil }
	if len(sequence) < 5 {
		return refuse()
	}
	wordNS := extractor.wordNS
	metadata := func(run *nativeXMLNode, kind, value string) bool {
		if run.Name != (xml.Name{Space: wordNS, Local: "r"}) || !nativeExactRevisionContainer(run, wordNS, "rsidR", "rsidRPr", "rsidDel") {
			return false
		}
		var content *nativeXMLNode
		for _, child := range run.Children {
			if child.Name == (xml.Name{Space: wordNS, Local: "rPr"}) {
				if !nativeExactPageFieldProperties(child, wordNS) {
					return false
				}
				continue
			}
			if content != nil {
				return false
			}
			content = child
		}
		if len(directNativeChildren(run, wordNS, "rPr")) > 1 || content == nil || content.Name != (xml.Name{Space: wordNS, Local: kind}) {
			return false
		}
		if kind == "fldChar" {
			actual, present := nativeAttr(content, wordNS, "fldCharType")
			return present && actual == value && nativeExactLeaf(content, xml.Name{Space: wordNS, Local: "fldCharType"})
		}
		if len(content.Children) != 0 {
			return false
		}
		for _, attr := range content.Attrs {
			if attr.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || attr.Value != "preserve" {
				return false
			}
		}
		instruction, ok := nativePageFieldInstruction(content.Text)
		return ok && instruction == value
	}
	if !metadata(sequence[0], "fldChar", "begin") || !metadata(sequence[2], "fldChar", "separate") || !metadata(sequence[4], "fldChar", "end") {
		return refuse()
	}
	instruction := "PAGE"
	if !metadata(sequence[1], "instrText", instruction) {
		instruction = "NUMPAGES"
		if !metadata(sequence[1], "instrText", instruction) {
			return refuse()
		}
	}
	result := sequence[3]
	if result.Name != (xml.Name{Space: wordNS, Local: "r"}) || !nativeExactRevisionContainer(result, wordNS, "rsidR", "rsidRPr", "rsidDel") || len(directNativeChildren(result, wordNS, "rPr")) > 1 {
		return refuse()
	}
	// Only the visible result's formatting enters layout. Keep original XML
	// anchors and result identity while excluding exact proofing metadata and
	// the deferred complex-script size from the conservative editable projection.
	// The resolver still reads their original source and refuses active script
	// semantics. No source bytes or editing permissions are changed.
	projection := *result
	projection.Children = nil
	for _, child := range result.Children {
		if child.Name == (xml.Name{Space: wordNS, Local: "rPr"}) {
			if len(directNativeChildren(child, wordNS, "noProof")) == 0 && len(directNativeChildren(child, wordNS, "szCs")) == 0 {
				projection.Children = append(projection.Children, child)
				continue
			}
			if !nativeExactPageFieldProperties(child, wordNS) {
				return refuse()
			}
			properties := *child
			properties.Children = nil
			for _, property := range child.Children {
				if property.Name.Local != "noProof" && property.Name.Local != "szCs" {
					properties.Children = append(properties.Children, property)
				}
			}
			projection.Children = append(projection.Children, &properties)
		} else {
			if child.Name != (xml.Name{Space: wordNS, Local: "t"}) || len(child.Children) != 0 {
				return refuse()
			}
			for _, attr := range child.Attrs {
				if attr.Name != (xml.Name{Space: "http://www.w3.org/XML/1998/namespace", Local: "space"}) || (attr.Value != "preserve" && attr.Value != "default") {
					return refuse()
				}
			}
			projection.Children = append(projection.Children, child)
		}
	}
	runs, unsafe, err := extractor.extractRunNode(partName, paragraphID, &projection)
	if err != nil {
		return NativeRunV1{}, false, err
	}
	if unsafe || len(runs) != 1 || runs[0].Kind != "text" {
		return refuse()
	}
	runs[0].PageField = instruction
	runs[0].Text = nativeString("")
	return runs[0], true, nil
}

// Character style and proofing metadata on non-visible field instruction runs
// are retained, not used to override the visible result run's style.
func nativeExactPageFieldProperties(node *nativeXMLNode, wordNS string) bool {
	if !nativeExactContainer(node) {
		return false
	}
	projection := *node
	projection.Children = nil
	seen := map[string]bool{}
	for _, child := range node.Children {
		if child.Name.Space != wordNS || seen[child.Name.Local] {
			return false
		}
		seen[child.Name.Local] = true
		switch child.Name.Local {
		case "rStyle":
			value, ok := nativeAttr(child, wordNS, "val")
			if !ok || !nativeIDPattern.MatchString(value) || !nativeExactLeaf(child, xml.Name{Space: wordNS, Local: "val"}) {
				return false
			}
		case "noProof":
			if !nativeNeutralSourceProperty(child, node, wordNS) {
				return false
			}
		default:
			projection.Children = append(projection.Children, child)
		}
	}
	return nativeExactParagraphMarkProperties(&projection, wordNS)
}
