package docxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"net/url"
	"path"
	"strings"
)

type nativeHyperlinkPatch struct {
	URL      *string
	HasURL   bool
	Expected string
}

func nativeHyperlinkURL(value string) bool {
	if len(value) == 0 || len(value) > 2048 || strings.ContainsAny(value, "\r\n\t \\<>") || !nativeMutationXMLTextValid(value) {
		return false
	}
	parsed, err := url.Parse(value)
	if err != nil {
		return false
	}
	switch parsed.Scheme {
	case "https", "http":
		return parsed.Hostname() != "" && parsed.User == nil
	case "mailto":
		return parsed.Opaque != ""
	}
	return false
}

func nativeHyperlinkRun(node *nativeXMLNode, ns string) bool {
	if node.Name != (xml.Name{Space: ns, Local: "r"}) {
		return false
	}
	count := 0
	for _, child := range node.Children {
		if child.Name.Space != ns || (child.Name.Local != "t" && child.Name.Local != "rPr") {
			return false
		}
		if child.Name.Local == "t" {
			count++
		}
	}
	return count == 1
}

// The first written tier models ordinary external links containing one plain
// text run. Bookmarks, fields, tooltip/target-frame semantics and complex link
// wrappers remain preserve-only, with the existing explicit diagnostic.
func (extractor *nativeExtractor) nativeEditableHyperlink(partName string, node *nativeXMLNode) (*NativeHyperlinkV1, bool) {
	if !nativeExactContainer(node, xml.Name{Space: extractor.relNS, Local: "id"}) || len(node.Children) != 1 || !nativeHyperlinkRun(node.Children[0], extractor.wordNS) {
		return nil, false
	}
	id, ok := nativeAttr(node, extractor.relNS, "id")
	if !ok {
		return nil, false
	}
	for _, rel := range extractor.pkg.rels[partName] {
		if rel.ID == id && rel.Type == extractor.relNS+"/hyperlink" && rel.External && nativeHyperlinkURL(rel.Target) {
			return &NativeHyperlinkV1{URL: rel.Target, Anchor: extractor.anchor(partName, node)}, true
		}
	}
	return nil, false
}

func applyNativeHyperlink(source []byte, revision string, mutation nativeDOCXMutationV1) (*NativeDOCXMutationResultV1, error) {
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
	var p *NativeParagraphV1
	start, end := 0, 0
	for _, block := range doc.Body.Blocks {
		candidate := block.Paragraph
		if candidate == nil {
			continue
		}
		offset := 0
		for _, run := range candidate.Runs {
			length := 0
			if run.Text != nil {
				length = nativeMutationUTF16CodeUnits(*run.Text)
			}
			if mutation.TargetKind == "run" && run.ID == mutation.TargetID {
				if run.Anchor.XMLSHA256 != mutation.ExpectedXMLSHA256 {
					return refuse("STALE_TARGET", "run anchor changed")
				}
				if run.Hyperlink != nil && mutation.Hyperlink.Expected != run.Hyperlink.Anchor.XMLSHA256 {
					return refuse("STALE_TARGET", "hyperlink anchor is required and must match")
				}
				if run.Hyperlink == nil && mutation.Hyperlink.Expected != "" {
					return refuse("STALE_TARGET", "target has no hyperlink")
				}
				p, start, end = candidate, offset, offset+length
			}
			offset += length
		}
		if mutation.TargetKind == "paragraph" && candidate.ID == mutation.TargetID {
			if candidate.Anchor.XMLSHA256 != mutation.ExpectedXMLSHA256 {
				return refuse("STALE_TARGET", "paragraph anchor changed")
			}
			if mutation.Hyperlink.Expected != "" {
				return refuse("INVALID_PAYLOAD", "paragraph links use the paragraph anchor")
			}
			p, start, end = candidate, 0, offset
		}
	}
	if p == nil || !nativePolicyAllows(p.EditPolicy, "hyperlink.set") {
		return refuse("UNSUPPORTED_CONSTRUCT", "choose supported body text")
	}
	if mutation.Range != nil {
		from, to := mutation.Range.StartUTF16, mutation.Range.EndUTF16
		if from < 0 || to <= from || to > end-start {
			return refuse("INVALID_RANGE", "hyperlink range must select text")
		}
		start, end = start+from, start+to
	}
	if start == end {
		return refuse("INVALID_RANGE", "select nonempty text to link")
	}
	partName := p.Anchor.PartName
	part := pkg.files[partName]
	root, err := parseNativeXML(partName, part)
	if err != nil {
		return nil, err
	}
	relNS := relNSTransitional
	if root.Name.Space == wordMLStrict {
		relNS = relNSStrict
	}
	relID := ""
	splices := map[string][]nativeTextSplice{}
	if mutation.Hyperlink.URL != nil {
		relID, err = nativeAddHyperlinkRelationship(pkg, partName, relNS, *mutation.Hyperlink.URL, splices)
		if err != nil {
			return nil, err
		}
	}
	prefix := "injLink"
	for strings.Contains(string(part), prefix) {
		prefix += "x"
	}
	wrap := func(run []byte, id string) []byte {
		if id == "" {
			return run
		}
		return []byte(`<` + prefix + `:hyperlink xmlns:` + prefix + `="` + root.Name.Space + `" xmlns:` + prefix + `R="` + relNS + `" ` + prefix + `R:id="` + nativeMutationEscapeAttribute(id) + `">` + string(run) + `</` + prefix + `:hyperlink>`)
	}
	offset := 0
	expected := []nativeLinkText{}
	for _, run := range p.Runs {
		if run.Text == nil {
			if mutation.TargetKind == "paragraph" {
				return refuse("UNSUPPORTED_CONSTRUCT", "hyperlink ranges require plain text runs")
			}
			continue
		}
		length := nativeMutationUTF16CodeUnits(*run.Text)
		from, to := max(0, start-offset), min(length, end-offset)
		oldURL := ""
		if run.Hyperlink != nil {
			oldURL = run.Hyperlink.URL
		}
		if to <= from {
			expected = append(expected, nativeLinkText{*run.Text, oldURL})
			offset += length
			continue
		}
		if !run.CanEditHyperlink {
			return refuse("UNSUPPORTED_CONSTRUCT", "selected run cannot be linked safely")
		}
		a, okA := nativeUTF16ByteOffset(*run.Text, from)
		b, okB := nativeUTF16ByteOffset(*run.Text, to)
		if !okA || !okB {
			return refuse("INVALID_RANGE", "hyperlink range splits a surrogate pair")
		}
		textNode := nativeNodeByPath(root, run.Anchor.Path)
		if textNode == nil {
			return refuse("STALE_TARGET", "run source is missing")
		}
		owner := textNode.parent
		replace := owner
		oldID := ""
		var oldLink *nativeXMLNode
		if run.Hyperlink != nil {
			oldLink = owner.parent
			replace = oldLink
			oldID, _ = nativeAttr(oldLink, relNS, "id")
		}
		replacement := []byte{}
		segments := []struct{ text, id, url string }{{(*run.Text)[:a], oldID, oldURL}, {(*run.Text)[a:b], relID, ""}, {(*run.Text)[b:], oldID, oldURL}}
		if mutation.Hyperlink.URL != nil {
			segments[1].url = *mutation.Hyperlink.URL
		}
		for _, segment := range segments {
			if segment.text == "" {
				continue
			}
			text, err := nativeFormatTextElement(part[textNode.Start:textNode.End], segment.text)
			if err != nil {
				return nil, err
			}
			raw := append([]byte{}, part[owner.Start:textNode.Start]...)
			raw = append(raw, text...)
			raw = append(raw, part[textNode.End:owner.End]...)
			// Moving a run out of its old wrapper must retain wrapper-local bindings.
			if oldLink != nil {
				raw = nativeCarryLinkNamespaces(raw, owner, oldLink)
			}
			replacement = append(replacement, wrap(raw, segment.id)...)
			expected = append(expected, nativeLinkText{segment.text, segment.url})
		}
		splices[partName] = append(splices[partName], nativeTextSplice{start: replace.Start, end: replace.End, text: replacement})
		offset += length
	}
	return writeNativeMutationSplices(source, pkg, doc, revision, splices, func(after *NativeDocumentV1) error {
		for _, block := range after.Body.Blocks {
			if block.Paragraph != nil && block.Paragraph.Anchor.Path == p.Anchor.Path {
				actual := []nativeLinkText{}
				for _, run := range block.Paragraph.Runs {
					if run.Text != nil {
						target := ""
						if run.Hyperlink != nil {
							target = run.Hyperlink.URL
						}
						actual = append(actual, nativeLinkText{*run.Text, target})
					}
				}
				if nativeLinkProjection(actual) != nativeLinkProjection(expected) {
					return nativeMutationError("POST_WRITE_MISMATCH", p.ID, "hyperlink text or targets did not round-trip")
				}
				return nil
			}
		}
		return nativeMutationError("POST_WRITE_MISMATCH", p.ID, "linked paragraph disappeared")
	})
}

type nativeLinkText struct{ text, url string }

func nativeLinkProjection(runs []nativeLinkText) string {
	var out, chunk strings.Builder
	target := ""
	flush := func() {
		if chunk.Len() > 0 {
			fmt.Fprintf(&out, "%d:%s:%d:%s", len(target), target, chunk.Len(), chunk.String())
			chunk.Reset()
		}
	}
	for _, run := range runs {
		if run.url != target {
			flush()
			target = run.url
		}
		chunk.WriteString(run.text)
	}
	flush()
	return out.String()
}

func nativeCarryLinkNamespaces(raw []byte, run, link *nativeXMLNode) []byte {
	declarations := ""
	for _, attr := range link.Attrs {
		if !nativeSettingsNamespaceDeclaration(attr) {
			continue
		}
		shadowed := false
		for _, own := range run.Attrs {
			if own.Name == attr.Name {
				shadowed = true
			}
		}
		if shadowed {
			continue
		}
		name := "xmlns"
		if attr.Name.Space == "xmlns" {
			name += ":" + attr.Name.Local
		}
		declarations += ` ` + name + `="` + nativeMutationEscapeAttribute(attr.Value) + `"`
	}
	at := nativeStartTagEnd(raw) - 1
	return []byte(string(raw[:at]) + declarations + string(raw[at:]))
}

func nativeAddHyperlinkRelationship(pkg *nativePackage, partName, relNS, target string, splices map[string][]nativeTextSplice) (string, error) {
	used := map[string]bool{}
	for _, rel := range pkg.rels[partName] {
		used[rel.ID] = true
		if rel.External && rel.Type == relNS+"/hyperlink" && rel.Target == target {
			return rel.ID, nil
		}
	}
	id := ""
	for n := 1; ; n++ {
		id = fmt.Sprintf("rIdLink%d", n)
		if !used[id] {
			break
		}
	}
	entry := `<Relationship xmlns="` + opcRelationshipsNS + `" Id="` + id + `" Type="` + relNS + `/hyperlink" Target="` + nativeMutationEscapeAttribute(target) + `" TargetMode="External"/>`
	name := pkg.relsPart[partName]
	if name == "" {
		name = path.Join(path.Dir(partName), "_rels", path.Base(partName)+".rels")
		if _, exists := pkg.files[name]; exists {
			return "", nativeMutationError("UNSUPPORTED_CONSTRUCT", "", "relationship part name is occupied")
		}
		data := []byte(`<Relationships xmlns="` + opcRelationshipsNS + `">` + entry + `</Relationships>`)
		// Refuse a new part unless the existing OPC content type defaults cover it.
		files := map[string][]byte{}
		for key, value := range pkg.files {
			files[key] = value
		}
		files[name] = data
		types, err := parseNativeContentTypes(pkg.files["[Content_Types].xml"], files)
		if err != nil || types[name] != "application/vnd.openxmlformats-package.relationships+xml" {
			return "", nativeMutationError("UNSUPPORTED_CONSTRUCT", "", "package does not declare relationship content types")
		}
		splices[name] = []nativeTextSplice{{start: 0, end: 0, text: data}}
		return id, nil
	}
	data := pkg.files[name]
	root, err := parseNativeXML(name, data)
	if err != nil {
		return "", err
	}
	close := bytes.LastIndex(data[root.Start:root.End], []byte("</"))
	if close >= 0 {
		at := root.Start + int64(close)
		splices[name] = []nativeTextSplice{{start: at, end: at, text: []byte(entry)}}
	} else {
		at := nativeStartTagEnd(data[root.Start:root.End]) - 2
		if at < 0 || data[root.Start+int64(at)] != '/' {
			return "", fmt.Errorf("unsupported relationship root")
		}
		replacement := string(data[root.Start:root.Start+int64(at)]) + ">" + entry + "</" + nativeFormatQName(data, root) + ">"
		splices[name] = []nativeTextSplice{{start: root.Start, end: root.End, text: []byte(replacement)}}
	}
	return id, nil
}
