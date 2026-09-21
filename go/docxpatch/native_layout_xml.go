package docxpatch

import (
	"fmt"
	"sort"
	"strings"
)

// Merge only the requested attributes and children. All untouched lexical bytes,
// including container attributes, whitespace, comments and extension nodes, survive.
func nativeMergePropertyContainer(part []byte, owner *nativeXMLNode, local string, order []string, written map[string]map[string]*string) (nativeTextSplice, error) {
	container := firstDirectNativeChild(owner, owner.Name.Space, local)
	prefix := nativeFormatQName(part, owner)
	at := strings.Index(prefix, ":")
	if at < 0 {
		return nativeTextSplice{}, fmt.Errorf("layout edits require an explicit WordprocessingML prefix")
	}
	prefix = prefix[:at+1]
	if container == nil {
		result := []byte("<" + prefix + local + ">")
		for _, key := range order {
			if changes, ok := written[key]; ok {
				child, err := nativeMergeLayoutAttributes(nil, nil, prefix, key, changes)
				if err != nil {
					return nativeTextSplice{}, err
				}
				result = append(result, child...)
			}
		}
		result = append(result, []byte("</"+prefix+local+">")...)
		tagEnd := nativeStartTagEnd(part[owner.Start:owner.End])
		if tagEnd < 1 || part[owner.Start+int64(tagEnd)-2] == '/' {
			return nativeTextSplice{}, fmt.Errorf("self-closing owner requires expansion")
		}
		pos := owner.Start + int64(tagEnd)
		return nativeTextSplice{start: pos, end: pos, text: result}, nil
	}
	if local == "pPr" && owner.Children[0] != container {
		return nativeTextSplice{}, fmt.Errorf("pPr must be the first child")
	}
	raw, err := nativeMergeLayoutChildren(part, container, order, written, prefix)
	return nativeTextSplice{start: container.Start, end: container.End, text: raw}, err
}

func nativeMergeLayoutChildren(part []byte, container *nativeXMLNode, order []string, written map[string]map[string]*string, prefix string) ([]byte, error) {
	rank := func(local string) int {
		for i, name := range order {
			if name == local {
				return i
			}
		}
		return len(order)
	}
	keys := []string{}
	for key := range written {
		keys = append(keys, key)
	}
	sort.Slice(keys, func(i, j int) bool { return rank(keys[i]) < rank(keys[j]) })
	raw := part[container.Start:container.End]
	tagEnd := nativeStartTagEnd(raw)
	if tagEnd < 1 {
		return nil, fmt.Errorf("unterminated properties")
	}
	if raw[tagEnd-2] == '/' {
		expanded := string(raw[:tagEnd-2]) + ">"
		for _, key := range keys {
			child, err := nativeMergeLayoutAttributes(nil, nil, prefix, key, written[key])
			if err != nil {
				return nil, err
			}
			expanded += string(child)
		}
		return []byte(expanded + "</" + nativeFormatQName(part, container) + ">"), nil
	}
	splices := []nativeTextSplice{}
	insertions := map[int64][]byte{}
	for _, key := range keys {
		var existing *nativeXMLNode
		for _, child := range container.Children {
			if child.Name.Space == container.Name.Space && child.Name.Local == key {
				if existing != nil {
					return nil, fmt.Errorf("duplicate %s", key)
				}
				existing = child
			}
		}
		child, err := nativeMergeLayoutAttributes(part, existing, prefix, key, written[key])
		if err != nil {
			return nil, err
		}
		if existing != nil {
			splices = append(splices, nativeTextSplice{start: existing.Start - container.Start, end: existing.End - container.Start, text: child})
			continue
		}
		pos := container.End - int64(len("</"+nativeFormatQName(part, container)+">"))
		for _, node := range container.Children {
			if rank(node.Name.Local) > rank(key) {
				pos = node.Start
				break
			}
		}
		pos -= container.Start
		insertions[pos] = append(insertions[pos], child...)
	}
	for pos, text := range insertions {
		splices = append(splices, nativeTextSplice{start: pos, end: pos, text: text})
	}
	sort.SliceStable(splices, func(i, j int) bool {
		if splices[i].start == splices[j].start {
			return splices[i].end < splices[j].end
		}
		return splices[i].start < splices[j].start
	})
	return applyNativeTextSplices(raw, splices)
}

func nativeMergeLayoutAttributes(part []byte, node *nativeXMLNode, prefix, local string, changes map[string]*string) ([]byte, error) {
	raw := []byte("<" + prefix + local + "/>")
	if node != nil {
		raw = part[node.Start:node.End]
		if len(node.Children) > 0 || strings.TrimSpace(node.Text) != "" {
			return nil, fmt.Errorf("%s must be a leaf", local)
		}
	}
	end := nativeStartTagEnd(raw)
	if end < 1 {
		return nil, fmt.Errorf("unterminated %s", local)
	}
	// Scan the source start tag without normalizing quote style or unrelated attributes.
	i := 1
	for i < end && !nativeMutationXMLSpace(raw[i]) && raw[i] != '/' && raw[i] != '>' {
		i++
	}
	output := append([]byte(nil), raw[:i]...)
	seen := map[string]bool{}
	for i < end-1 && raw[i] != '/' {
		start := i
		for i < end && nativeMutationXMLSpace(raw[i]) {
			i++
		}
		if raw[i] == '/' || raw[i] == '>' {
			output = append(output, raw[start:i]...)
			break
		}
		nameStart := i
		for i < end && raw[i] != '=' && !nativeMutationXMLSpace(raw[i]) {
			i++
		}
		name := string(raw[nameStart:i])
		for i < end && raw[i] != '=' {
			i++
		}
		i++
		for i < end && nativeMutationXMLSpace(raw[i]) {
			i++
		}
		quote := raw[i]
		i++
		for i < end && raw[i] != quote {
			i++
		}
		i++
		attr := strings.TrimPrefix(name, prefix)
		value, change := changes[attr]
		change = change && name == prefix+attr
		if change {
			seen[attr] = true
			if value != nil {
				output = append(output, []byte(" "+name+`="`+nativeMutationEscapeAttribute(*value)+`"`)...)
			}
		} else {
			output = append(output, raw[start:i]...)
		}
	}
	keys := []string{}
	for key := range changes {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	for _, key := range keys {
		if !seen[key] && changes[key] != nil {
			output = append(output, []byte(" "+prefix+key+`="`+nativeMutationEscapeAttribute(*changes[key])+`"`)...)
		}
	}
	return append(output, raw[i:]...), nil
}
