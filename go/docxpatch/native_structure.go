package docxpatch

import (
	"fmt"
	"strings"
)

// Splitting duplicates only plain paragraph properties and single-text runs.
// Section boundaries, annotations, fields and relationship-bearing content
// must not acquire a new scope as a side effect of Enter.
func nativeSplittableParagraph(node *nativeXMLNode, ns string) bool {
	for _, child := range node.Children {
		if child.Name.Space != ns {
			return false
		}
		if child.Name.Local == "pPr" {
			if firstDirectNativeChild(child, ns, "sectPr") != nil {
				return false
			}
			continue
		}
		if child.Name.Local != "r" || child.Text != "" {
			return false
		}
		texts := 0
		for _, item := range child.Children {
			if item.Name.Space != ns || (item.Name.Local != "t" && item.Name.Local != "rPr") {
				return false
			}
			if item.Name.Local == "t" {
				texts++
			}
		}
		if texts != 1 {
			return false
		}
	}
	return true
}

func applyNativeParagraphStructure(source []byte, revision string, mutation nativeDOCXMutationV1) (*NativeDOCXMutationResultV1, error) {
	refuse := func(code, message string) (*NativeDOCXMutationResultV1, error) {
		return nil, nativeMutationError(code, mutation.TargetID, message)
	}
	if nativeSHA(source) != revision {
		return refuse("STALE_REVISION", "package revision changed")
	}
	doc, err := ExtractNativeDocumentV1(source)
	if err != nil {
		return nil, err
	}
	pkg, err := openNativeDOCXPackage(source)
	if err != nil {
		return nil, err
	}
	if nativeDOCXPackageHasDigitalSignature(pkg) {
		return refuse("SIGNED_PACKAGE", "signed packages cannot be changed")
	}
	index := -1
	for i, block := range doc.Body.Blocks {
		if block.Paragraph != nil && block.Paragraph.ID == mutation.TargetID {
			index = i
			break
		}
	}
	if mutation.TargetKind != "paragraph" || index < 0 {
		return refuse("UNSUPPORTED_TARGET_KIND", "choose a body paragraph")
	}
	paragraph := doc.Body.Blocks[index].Paragraph
	if paragraph.Anchor.XMLSHA256 != mutation.ExpectedXMLSHA256 {
		return refuse("STALE_TARGET", "paragraph anchor changed")
	}
	if !nativePolicyAllows(paragraph.EditPolicy, mutation.Operation) {
		return refuse("UNSUPPORTED_CONSTRUCT", "paragraph does not support this structural operation")
	}
	partName := paragraph.Anchor.PartName
	part := pkg.files[partName]
	root, err := parseNativeXML(partName, part)
	if err != nil {
		return nil, err
	}
	node := nativeNodeByPath(root, paragraph.Anchor.Path)
	if node == nil {
		return refuse("STALE_TARGET", "paragraph source is missing")
	}
	// Declare the word namespace locally; never assume a producer's prefix.
	start := `<w:p xmlns:w="` + node.Name.Space + `">`
	empty := `<w:r><w:t></w:t></w:r>`
	inserted := start + empty + `</w:p>`
	splice := nativeTextSplice{start: node.End, end: node.End, text: []byte(inserted)}
	expectedLeft, expectedRight := nativeStructureText(paragraph), ""
	if mutation.Operation == "paragraph.split" {
		if !nativeSplittableParagraph(node, node.Name.Space) {
			return refuse("UNSUPPORTED_CONSTRUCT", "paragraph cannot be split safely")
		}
		var run *NativeRunV1
		offset := 0
		for i := range paragraph.Runs {
			candidate := &paragraph.Runs[i]
			if candidate.ID == mutation.SplitRunID {
				run = candidate
				break
			}
			if candidate.Text != nil {
				offset += nativeMutationUTF16CodeUnits(*candidate.Text)
			}
		}
		if run == nil || run.Text == nil {
			return refuse("STALE_TARGET", "split run is absent")
		}
		at, ok := nativeUTF16ByteOffset(*run.Text, mutation.SplitOffset)
		if !ok {
			return refuse("INVALID_RANGE", "split must be a UTF-16 character boundary")
		}
		textNode := nativeNodeByPath(node, run.Anchor.Path)
		if textNode == nil || textNode.parent == nil {
			return refuse("STALE_TARGET", "split run source is missing")
		}
		owner := textNode.parent
		tagEnd := int64(nativeStartTagEnd(part[node.Start:node.End]))
		runTagEnd := int64(nativeStartTagEnd(part[owner.Start:owner.End]))
		leftText, err := nativeFormatTextElement(part[textNode.Start:textNode.End], (*run.Text)[:at])
		if err != nil {
			return nil, err
		}
		rightText, err := nativeFormatTextElement(part[textNode.Start:textNode.End], (*run.Text)[at:])
		if err != nil {
			return nil, err
		}
		runStart := string(part[owner.Start : owner.Start+runTagEnd])
		props := ""
		if p := firstDirectNativeChild(owner, node.Name.Space, "rPr"); p != nil {
			props = string(part[p.Start:p.End])
		}
		runEnd := "</" + nativeFormatQName(part, owner) + ">"
		pEnd := "</" + nativeFormatQName(part, node) + ">"
		pProps := ""
		if p := firstDirectNativeChild(node, node.Name.Space, "pPr"); p != nil {
			pProps = string(part[p.Start:p.End])
		}
		// New paragraph uses the source prefix so copied property/run bytes remain in scope.
		pStart := "<" + nativeFormatQName(part, node)
		for _, attr := range node.Attrs {
			if attr.Name.Space == "xmlns" {
				pStart += ` xmlns:` + attr.Name.Local + `="` + nativeMutationEscapeAttribute(attr.Value) + `"`
			}
			if attr.Name.Space == "" && attr.Name.Local == "xmlns" {
				pStart += ` xmlns="` + nativeMutationEscapeAttribute(attr.Value) + `"`
			}
		}
		pStart += ">"
		left := string(part[node.Start:owner.Start]) + runStart + props + string(leftText) + runEnd + pEnd
		right := pStart + pProps + runStart + props + string(rightText) + runEnd + string(part[owner.End:node.End-int64(len(pEnd))]) + pEnd
		if tagEnd <= 0 {
			return refuse("UNSUPPORTED_LEXICAL_FORM", "invalid paragraph start tag")
		}
		splice = nativeTextSplice{start: node.Start, end: node.End, text: []byte(left + right)}
		boundary, ok := nativeUTF16ByteOffset(expectedLeft, offset+mutation.SplitOffset)
		if !ok {
			return refuse("INVALID_RANGE", "invalid paragraph split offset")
		}
		expectedLeft, expectedRight = expectedLeft[:boundary], expectedLeft[boundary:]
	}
	return writeNativeMutationSplices(source, pkg, doc, revision, map[string][]nativeTextSplice{partName: {splice}}, func(after *NativeDocumentV1) error {
		if len(after.Body.Blocks) != len(doc.Body.Blocks)+1 {
			return fmt.Errorf("paragraph insertion changed unexpected blocks")
		}
		if nativeStructureText(after.Body.Blocks[index].Paragraph) != expectedLeft || nativeStructureText(after.Body.Blocks[index+1].Paragraph) != expectedRight {
			return nativeMutationError("POST_WRITE_MISMATCH", mutation.TargetID, "paragraph text did not round-trip")
		}
		for i := range doc.Body.Blocks {
			if i == index {
				continue
			}
			j := i
			if i > index {
				j++
			}
			if doc.Body.Blocks[i].Paragraph != nil && nativeStructureText(doc.Body.Blocks[i].Paragraph) != nativeStructureText(after.Body.Blocks[j].Paragraph) {
				return nativeMutationError("POST_WRITE_MISMATCH", mutation.TargetID, "neighboring paragraph changed")
			}
		}
		return nil
	})
}

func nativeStructureText(p *NativeParagraphV1) string {
	if p == nil {
		return "\x00"
	}
	var text strings.Builder
	for _, run := range p.Runs {
		if run.Text != nil {
			text.WriteString(*run.Text)
		}
	}
	return text.String()
}
