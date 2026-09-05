package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"regexp"
)

const (
	NativeXLSXMaxXMLTokens   = 4_000_000
	NativeXLSXMaxXMLElements = 2_000_000
)

var nativeXMLDeclarationPattern = regexp.MustCompile(`^[\x20\x09\x0D\x0A]*version[\x20\x09\x0D\x0A]*=[\x20\x09\x0D\x0A]*(?:"1\.0"|'1\.0')(?:[\x20\x09\x0D\x0A]+encoding[\x20\x09\x0D\x0A]*=[\x20\x09\x0D\x0A]*(?:"[Uu][Tt][Ff]-8"|'[Uu][Tt][Ff]-8'))?(?:[\x20\x09\x0D\x0A]+standalone[\x20\x09\x0D\x0A]*=[\x20\x09\x0D\x0A]*(?:"(?:yes|no)"|'(?:yes|no)'))?[\x20\x09\x0D\x0A]*$`)

func validNativeXMLDeclaration(instruction xml.ProcInst) bool {
	return instruction.Target == "xml" && nativeXMLDeclarationPattern.Match(instruction.Inst)
}

func isNamespaceDeclaration(attribute xml.Attr) bool {
	return (attribute.Name.Space == "" && attribute.Name.Local == "xmlns") || attribute.Name.Space == "xmlns"
}

func unexpectedSemanticXMLAttributes(start xml.StartElement, allowed ...xml.Name) []xml.Name {
	allowedSet := make(map[xml.Name]bool, len(allowed))
	for _, name := range allowed {
		allowedSet[name] = true
	}
	var unexpected []xml.Name
	for _, attribute := range start.Attr {
		if isNamespaceDeclaration(attribute) || allowedSet[attribute.Name] {
			continue
		}
		unexpected = append(unexpected, attribute.Name)
	}
	return unexpected
}

func requireOnlySemanticXMLAttributes(start xml.StartElement, allowed ...xml.Name) error {
	if unexpected := unexpectedSemanticXMLAttributes(start, allowed...); len(unexpected) != 0 {
		return fmt.Errorf("element {%s}%s has unexpected semantic attribute {%s}%s", start.Name.Space, start.Name.Local, unexpected[0].Space, unexpected[0].Local)
	}
	return nil
}

func asciiEqualFold(left, right string) bool {
	if len(left) != len(right) {
		return false
	}
	for index := range left {
		a, b := left[index], right[index]
		if a >= 'A' && a <= 'Z' {
			a += 'a' - 'A'
		}
		if b >= 'A' && b <= 'Z' {
			b += 'a' - 'A'
		}
		if a != b {
			return false
		}
	}
	return true
}

func preflightNativeCoreXML(data []byte) (tokens, elements int, err error) {
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootSeen, rootClosed := false, false
	declarationSeen, prefixSeen := false, false
	for {
		token, tokenErr := decoder.Token()
		if tokenErr == io.EOF {
			break
		}
		if tokenErr != nil {
			return 0, 0, tokenErr
		}
		tokens++
		if tokens > NativeXLSXMaxXMLTokens {
			return 0, 0, fmt.Errorf("XML exceeds %d tokens", NativeXLSXMaxXMLTokens)
		}
		switch token := token.(type) {
		case xml.StartElement:
			seenAttributes := map[xml.Name]bool{}
			for _, attribute := range token.Attr {
				if seenAttributes[attribute.Name] {
					return 0, 0, fmt.Errorf("element {%s}%s has duplicate attribute {%s}%s", token.Name.Space, token.Name.Local, attribute.Name.Space, attribute.Name.Local)
				}
				seenAttributes[attribute.Name] = true
			}
			if depth == 0 {
				if rootSeen || rootClosed {
					return 0, 0, fmt.Errorf("XML has multiple root elements")
				}
				rootSeen = true
			}
			prefixSeen = true
			depth++
			elements++
			if depth > NativeXLSXMaxXMLDepth {
				return 0, 0, fmt.Errorf("XML nesting exceeds %d", NativeXLSXMaxXMLDepth)
			}
			if elements > NativeXLSXMaxXMLElements {
				return 0, 0, fmt.Errorf("XML exceeds %d elements", NativeXLSXMaxXMLElements)
			}
		case xml.EndElement:
			depth--
			if depth < 0 {
				return 0, 0, fmt.Errorf("XML has an unmatched closing element")
			}
			if depth == 0 {
				rootClosed = true
			}
		case xml.CharData:
			if depth == 0 {
				if len(bytes.TrimSpace(token)) != 0 {
					return 0, 0, fmt.Errorf("XML has text outside its root")
				}
				if !rootSeen && len(token) != 0 {
					prefixSeen = true
				}
			}
		case xml.Comment:
			if !rootSeen {
				prefixSeen = true
			}
		case xml.ProcInst:
			if declarationSeen || prefixSeen || rootSeen || !validNativeXMLDeclaration(token) {
				return 0, 0, fmt.Errorf("XML has an invalid, repeated, or misplaced declaration/processing instruction %q", token.Target)
			}
			declarationSeen, prefixSeen = true, true
		case xml.Directive:
			return 0, 0, fmt.Errorf("XML directives/DOCTYPE are unsupported")
		}
	}
	if depth != 0 || !rootSeen || !rootClosed {
		return 0, 0, fmt.Errorf("XML requires exactly one complete root")
	}
	return tokens, elements, nil
}
