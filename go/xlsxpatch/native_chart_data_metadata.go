package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// This classifier never projects metadata into cell values or visibility. It
// retains an unsupported record with a distinct code only for the closed ECMA
// CT_BookViews/CT_BookView and CT_SheetDimension subsets below. Extensions,
// foreign attributes, unknown children and malformed values retain generic codes.
func consumeNativeChartDataMetadata(decoder *xml.Decoder, root xml.StartElement, namespace string) (bool, error) {
	qualified := root.Name.Space == namespace
	book := root.Name.Local == "bookViews"
	if book {
		qualified = qualified && len(unexpectedSemanticXMLAttributes(root)) == 0
	} else {
		qualified = qualified && root.Name.Local == "dimension" && nativeChartDimensionAttributes(root)
	}
	level, views := 1, 0
	for level > 0 {
		token, err := decoder.Token()
		if err != nil {
			return false, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			level++
			if level > NativeXLSXMaxXMLDepth {
				return false, fmt.Errorf("XML nesting exceeds %d", NativeXLSXMaxXMLDepth)
			}
			if book && level == 2 && token.Name == (xml.Name{Space: namespace, Local: "workbookView"}) {
				views++
				qualified = qualified && views <= 256 && nativeChartWorkbookViewAttributes(token)
			} else {
				qualified = false
			}
		case xml.EndElement:
			level--
		case xml.CharData:
			if strings.Trim(string(token), " \t\r\n") != "" {
				qualified = false
			}
		case xml.ProcInst:
			return false, fmt.Errorf("XML contains unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return false, fmt.Errorf("XML contains unsupported directive")
		}
	}
	return qualified && (!book || views > 0), nil
}

func nativeChartDimensionAttributes(root xml.StartElement) bool {
	seen := false
	for _, attribute := range root.Attr {
		if isNamespaceDeclaration(attribute) {
			continue
		}
		if seen || attribute.Name != (xml.Name{Local: "ref"}) {
			return false
		}
		seen = true
		canonical, err := canonicalNativeRangeReferenceV2(attribute.Value)
		if err != nil || canonical != attribute.Value {
			return false
		}
	}
	return seen
}

func nativeChartWorkbookViewAttributes(root xml.StartElement) bool {
	seen := map[string]bool{}
	for _, attribute := range root.Attr {
		if isNamespaceDeclaration(attribute) {
			continue
		}
		if attribute.Name.Space != "" || seen[attribute.Name.Local] {
			return false
		}
		seen[attribute.Name.Local] = true
		value := attribute.Value
		switch attribute.Name.Local {
		case "visibility":
			if value != "visible" && value != "hidden" && value != "veryHidden" {
				return false
			}
		case "minimized", "showHorizontalScroll", "showVerticalScroll", "showSheetTabs", "autoFilterDateGrouping":
			if value != "0" && value != "1" && value != "false" && value != "true" {
				return false
			}
		case "xWindow", "yWindow":
			if _, err := strconv.ParseInt(value, 10, 32); err != nil {
				return false
			}
		case "windowWidth", "windowHeight", "tabRatio", "firstSheet", "activeTab":
			if _, err := strconv.ParseUint(value, 10, 32); err != nil {
				return false
			}
		default:
			return false
		}
	}
	return true
}
