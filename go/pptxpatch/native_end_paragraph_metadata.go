package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strings"
)

// Only non-layout checking flags and a language identical to the final visible
// run are qualified. End-mark font metrics and empty paragraphs remain strict.
func qualifyNativeEndParagraphMetadata(node *nativeXMLNode, paragraph NativeParagraph) error {
	if node == nil {
		return nil
	}
	refuse := func() error {
		return fmt.Errorf("pptxpatch: native extract: active or unqualified end-paragraph formatting is not representable in v1")
	}
	if requireOnlyNativeAttrs(node, xml.Name{Local: "dirty"}, xml.Name{Local: "smtClean"}, xml.Name{Local: "lang"}) != nil || requireOnlyNativeChildren(node) != nil || len(paragraph.Runs) == 0 {
		return refuse()
	}
	last := paragraph.Runs[len(paragraph.Runs)-1]
	if last.Text == nil || strings.TrimSpace(*last.Text) == "" {
		return refuse()
	}
	for _, attr := range node.Attrs {
		switch attr.Name.Local {
		case "dirty", "smtClean":
			if _, err := nativeBool(attr.Value); err != nil {
				return refuse()
			}
		case "lang":
			if !validNativeLanguage(attr.Value) || last.Language == nil || *last.Language != attr.Value {
				return refuse()
			}
		}
	}
	return nil
}

func nativeHasEndParagraphMetadata(node *nativeXMLNode, dialect nativeExtractDialect) bool {
	if node == nil {
		return false
	}
	if node.Name == (xml.Name{Space: dialect.drawing, Local: "endParaRPr"}) {
		return true
	}
	for _, child := range node.Children {
		if nativeHasEndParagraphMetadata(child, dialect) {
			return true
		}
	}
	return false
}
