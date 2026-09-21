package xlsxpatch

import (
	"encoding/xml"
	"strings"
)

// NativeWorkbookAutoFilterV1 is the worksheet AutoFilter range as authored.
// Criteria stay in the source XML; the desktop Filter Mode note reads this ref.
type NativeWorkbookAutoFilterV1 struct {
	Ref string `json:"ref"`
}

// NativeWorkbookAutoFilterV2 mirrors NativeWorkbookAutoFilterV1 on the v2 wire.
type NativeWorkbookAutoFilterV2 struct {
	Ref string `json:"ref"`
}

func parseNativeAutoFilter(decoder *xml.Decoder, root xml.StartElement) (*NativeWorkbookAutoFilterV1, error) {
	ref, found, err := unqualifiedXMLAttribute(root, "ref")
	if err != nil {
		return nil, err
	}
	if err := skipNativeXMLElement(decoder, root, 1); err != nil {
		return nil, err
	}
	if !found || parseCanonicalA1Range(ref) == nil || len(unexpectedSemanticXMLAttributes(root, xml.Name{Local: "ref"})) != 0 {
		return nil, nil
	}
	return &NativeWorkbookAutoFilterV1{Ref: ref}, nil
}

func parseCanonicalA1Range(ref string) *struct{ row, column, endRow, endColumn int } {
	if ref == "" || strings.ContainsAny(ref, "$!") {
		return nil
	}
	row, column, endRow, endColumn, err := parseDimensionReference(ref)
	if err != nil {
		return nil
	}
	start := cellReference(row, column)
	canonical := start
	if row != endRow || column != endColumn {
		canonical = start + ":" + cellReference(endRow, endColumn)
	}
	if canonical != ref {
		return nil
	}
	return &struct{ row, column, endRow, endColumn int }{row, column, endRow, endColumn}
}

func nativeAutoFilterIssues(filter *NativeWorkbookAutoFilterV1, path string) []nativeSheetViewIssue {
	if filter == nil {
		return nil
	}
	if parseCanonicalA1Range(filter.Ref) == nil {
		return []nativeSheetViewIssue{{code: "INVALID_VALUE", path: path + "/auto_filter/ref", message: "AutoFilter ref must be a canonical bounded A1 range"}}
	}
	return nil
}
