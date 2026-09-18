package xlsxpatch

import (
	"encoding/xml"
	"sort"
	"strings"
)

// nativeDimensionNeutralityPolicy names the only policy a consumer may act on.
const nativeDimensionNeutralityPolicy = "non-dimensional-worksheet-markup-v1"

// NativeSheetDimensionNeutralityV1 records, for one worksheet part, which
// bounded-projection disclosure codes were raised only by markup that ECMA-376
// gives no role in a row height, a column width or a merged rectangle.
//
// The bounded dimension projection reads sheetFormatPr, cols/col, row and
// mergeCells. A worksheet may still carry markup outside that vocabulary, and
// the extractor discloses it (SHEET_FORMAT_EXTRAS, ROW_DIMENSION_EXTRAS,
// COLUMN_DIMENSION_EXTRAS, COLS_ATTRIBUTES, SHEET_VIEW_GEOMETRY,
// WORKSHEET_ATTRIBUTES). Most of that markup is display state, outline
// bookkeeping, border flags or text-descent metadata: it is not projected, and
// it also cannot move a row band, a column band or a merge. This preview
// separates the two cases so that "we did not model it" stops being read as
// "the dimensions may be wrong".
//
// FOREIGN_WORKSHEET_MARKUP is the one code here that names markup outside
// SpreadsheetML entirely, and it is cleared only for anchored form controls
// (see nativeNeutralForeignMarkup).
//
// Nothing here reproduces the markup it clears. Frozen panes, zoom, gridline
// visibility, outline levels, thick-edge flags, descent metadata and the
// controls themselves stay omitted and stay disclosed. A code that is absent
// from Codes keeps refusing.
type NativeSheetDimensionNeutralityV1 struct {
	SheetPart string   `json:"sheet_part"`
	Policy    string   `json:"policy"`
	Codes     []string `json:"codes"`
	Warnings  []string `json:"warnings"`
}

const nativeDimensionNeutralityUnavailable = "No unmodeled worksheet markup was qualified as non-dimensional for this sheet. Every dimension refusal stands."

const nativeDimensionNeutralityAvailable = "The listed disclosures were raised only by worksheet markup that ECMA-376 gives no role in a row height, a column width or a merged rectangle, so the bounded dimension projection is unchanged. The markup itself stays omitted: view state, outline levels, thick-edge flags and text-descent metadata are not reproduced."

// Attributes the bounded projection reads plus attributes proven not to change
// a dimension. Everything else keeps its code refusing.
//
//	sheetFormatPr (ECMA-376 §18.3.1.81): baseColWidth, defaultColWidth,
//	  defaultRowHeight and zeroHeight carry the dimensions. customHeight only
//	  records that defaultRowHeight was authored rather than derived;
//	  thickTop/thickBottom are border flags; outlineLevelRow/outlineLevelCol are
//	  outline summaries.
//	col (§18.3.1.13): width, hidden, customWidth, bestFit and style carry the
//	  width; collapsed and outlineLevel are outline state. phonetic is NOT
//	  listed: phonetic guides participate in automatic sizing.
//	row (§18.3.1.73): ht, hidden, customHeight, s and spans carry the height;
//	  collapsed, customFormat, outlineLevel, thickTop and thickBot are outline,
//	  style and border flags. ph is NOT listed, for the same reason as phonetic.
var (
	nativeNeutralSheetFormatAttributes = map[string]bool{
		"baseColWidth": true, "defaultColWidth": true, "defaultRowHeight": true, "zeroHeight": true,
		"customHeight": true, "thickTop": true, "thickBottom": true,
		"outlineLevelRow": true, "outlineLevelCol": true,
	}
	nativeNeutralColumnAttributes = map[string]bool{
		"min": true, "max": true, "width": true, "hidden": true, "customWidth": true, "bestFit": true, "style": true,
		"collapsed": true, "outlineLevel": true,
	}
	nativeNeutralRowAttributes = map[string]bool{
		"r": true, "ht": true, "hidden": true, "customHeight": true, "s": true, "spans": true,
		"collapsed": true, "customFormat": true, "outlineLevel": true, "thickTop": true, "thickBot": true,
	}
	// CT_SheetView (§18.3.1.87) is window state. rightToLeft is deliberately
	// absent: it mirrors the sheet's layout direction.
	nativeNeutralSheetViewAttributes = map[string]bool{
		"windowProtection": true, "showFormulas": true, "showGridLines": true, "showRowColHeaders": true,
		"showZeros": true, "tabSelected": true, "showRuler": true, "showOutlineSymbols": true,
		"defaultGridColor": true, "showWhiteSpace": true, "view": true, "topLeftCell": true, "colorId": true,
		"zoomScale": true, "zoomScaleNormal": true, "zoomScaleSheetLayoutView": true, "zoomScalePageLayoutView": true,
		"workbookViewId": true,
	}
	nativeNeutralPaneAttributes      = map[string]bool{"xSplit": true, "ySplit": true, "topLeftCell": true, "activePane": true, "state": true}
	nativeNeutralSelectionAttributes = map[string]bool{"pane": true, "activeCell": true, "activeCellId": true, "sqref": true}
)

// previewNativeDimensionNeutrality re-reads one worksheet part within the
// shared bounded preview limits and reports the disclosure codes it can clear.
// An unreadable, oversized or unexpected part clears nothing.
func previewNativeDimensionNeutrality(raw []byte, part string) NativeSheetDimensionNeutralityV1 {
	result := NativeSheetDimensionNeutralityV1{
		SheetPart: part, Policy: nativeDimensionNeutralityPolicy,
		Codes: []string{}, Warnings: []string{nativeDimensionNeutralityUnavailable},
	}
	root, err := parsePreviewXML(raw)
	if err != nil || root.name.Local != "worksheet" ||
		(root.name.Space != spreadsheetMLTransitional && root.name.Space != spreadsheetMLStrict) {
		return result
	}
	codes := map[string]bool{}
	// CT_Worksheet (ECMA-376 §18.3.1.99) declares no attributes at all, so any
	// attribute here is markup-compatibility or extension metadata. Ignorable
	// content markers (ISO/IEC 29500-3) cannot carry a cell dimension.
	codes["WORKSHEET_ATTRIBUTES"] = true
	if nativeNeutralSheetViews(root) {
		codes["SHEET_VIEW_GEOMETRY"] = true
	}
	if format := root.child("sheetFormatPr"); format != nil &&
		nativeNeutralPreviewLeaf(format, nativeNeutralSheetFormatAttributes) {
		codes["SHEET_FORMAT_EXTRAS"] = true
	}
	if cols := root.child("cols"); cols != nil && nativeNeutralColumns(root, cols) {
		// CT_Cols (§18.3.1.17) declares no attributes.
		codes["COLS_ATTRIBUTES"] = true
		codes["COLUMN_DIMENSION_EXTRAS"] = true
	}
	if data := root.child("sheetData"); data != nil && nativeNeutralRows(root, data) {
		codes["ROW_DIMENSION_EXTRAS"] = true
	}
	if nativeNeutralForeignMarkup(root) {
		codes["FOREIGN_WORKSHEET_MARKUP"] = true
	}
	ordered := make([]string, 0, len(codes))
	for code := range codes {
		ordered = append(ordered, code)
	}
	sort.Strings(ordered)
	result.Codes = ordered
	result.Warnings = []string{nativeDimensionNeutralityAvailable}
	return result
}

// nativeNeutralForeignMarkup accepts a worksheet whose only foreign-namespace
// direct children are markup-compatibility AlternateContent blocks holding
// nothing but a SpreadsheetML <controls> block.
//
// FOREIGN_WORKSHEET_MARKUP is raised by any direct child of CT_Worksheet
// outside the SpreadsheetML namespace, which is how Excel writes the form
// controls block: mc:AlternateContent wrapping mc:Choice Requires="x14". The
// controls inside are anchored objects. ECMA-376 gives CT_ObjectAnchor no role
// in a row height, a column width or a merged rectangle - the anchor is read
// from the grid, not written to it - so the bounded dimension projection is
// unchanged. An mc:Fallback is accepted only when it is empty, because a
// Fallback with content states an alternative this reader is not applying.
//
// Nothing here reproduces the markup it clears. A worksheet carrying any other
// foreign child, or an AlternateContent holding anything but controls, keeps
// the code refusing.
func nativeNeutralForeignMarkup(root *previewXML) bool {
	for _, child := range root.children {
		if child.name.Space == root.name.Space {
			continue
		}
		if child.name != (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "AlternateContent"}) {
			return false
		}
		if !nativeNeutralPreviewAttributesEmpty(child) || strings.TrimSpace(child.text) != "" {
			return false
		}
		for _, branch := range child.children {
			if branch.name == (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "Fallback"}) {
				if len(branch.children) != 0 || strings.TrimSpace(branch.text) != "" {
					return false
				}
				continue
			}
			if branch.name != (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "Choice"}) {
				return false
			}
			if strings.TrimSpace(branch.text) != "" {
				return false
			}
			for _, block := range branch.children {
				if block.name != (xml.Name{Space: root.name.Space, Local: "controls"}) {
					return false
				}
			}
		}
	}
	return true
}

// nativeNeutralPreviewLeaf accepts an attribute-only element whose attributes
// are all recognised, plus the x14ac descent hint.
func nativeNeutralPreviewLeaf(node *previewXML, allowed map[string]bool) bool {
	if len(node.children) != 0 || strings.TrimSpace(node.text) != "" {
		return false
	}
	for _, attr := range node.attrs {
		if isPreviewNamespaceDeclaration(attr) {
			continue
		}
		if attr.Name.Space == nativeRowDescentNamespace && attr.Name.Local == "dyDescent" {
			continue
		}
		if attr.Name.Space != "" || !allowed[attr.Name.Local] {
			return false
		}
	}
	return true
}

// nativeNeutralSheetViews accepts exactly one sheetViews holding exactly one
// sheetView whose attributes and pane/selection children are recognised window
// state. Multiple views, pivot selections, extension lists, foreign markup and
// an explicit rightToLeft keep SHEET_VIEW_GEOMETRY refusing.
func nativeNeutralSheetViews(root *previewXML) bool {
	seen := 0
	for _, child := range root.children {
		if child.name.Local == "sheetViews" {
			seen++
		}
	}
	if seen == 0 {
		// Absent sheetViews never raises the code; clearing it is harmless and
		// keeps the answer independent of what the extractor happened to emit.
		return true
	}
	views := root.child("sheetViews")
	if seen != 1 || views == nil {
		return false
	}
	// CT_SheetViews (§18.3.1.88) declares no attributes.
	if !nativeNeutralPreviewAttributesEmpty(views) {
		return false
	}
	if strings.TrimSpace(views.text) != "" || len(views.children) != 1 {
		return false
	}
	view := views.children[0]
	if view.name.Space != root.name.Space || view.name.Local != "sheetView" {
		return false
	}
	if strings.TrimSpace(view.text) != "" {
		return false
	}
	for _, attr := range view.attrs {
		if isPreviewNamespaceDeclaration(attr) {
			continue
		}
		if attr.Name.Space != "" {
			return false
		}
		if attr.Name.Local == "rightToLeft" {
			if attr.Value != "0" && attr.Value != "false" {
				return false
			}
			continue
		}
		if !nativeNeutralSheetViewAttributes[attr.Name.Local] {
			return false
		}
	}
	for _, child := range view.children {
		if child.name.Space != root.name.Space {
			return false
		}
		switch child.name.Local {
		case "pane":
			if !nativeNeutralPreviewLeaf(child, nativeNeutralPaneAttributes) {
				return false
			}
		case "selection":
			if !nativeNeutralPreviewLeaf(child, nativeNeutralSelectionAttributes) {
				return false
			}
		default:
			return false
		}
	}
	return true
}

func nativeNeutralPreviewAttributesEmpty(node *previewXML) bool {
	for _, attr := range node.attrs {
		if !isPreviewNamespaceDeclaration(attr) {
			return false
		}
	}
	return true
}

func nativeNeutralColumns(root, cols *previewXML) bool {
	if strings.TrimSpace(cols.text) != "" {
		return false
	}
	for _, col := range cols.children {
		if col.name.Space != root.name.Space || col.name.Local != "col" {
			return false
		}
		if !nativeNeutralPreviewLeaf(col, nativeNeutralColumnAttributes) {
			return false
		}
	}
	return true
}

func nativeNeutralRows(root, data *previewXML) bool {
	for _, row := range data.children {
		if row.name.Space != root.name.Space || row.name.Local != "row" {
			return false
		}
		for _, attr := range row.attrs {
			if isPreviewNamespaceDeclaration(attr) {
				continue
			}
			if attr.Name.Space == nativeRowDescentNamespace && attr.Name.Local == "dyDescent" {
				continue
			}
			if attr.Name.Space != "" || !nativeNeutralRowAttributes[attr.Name.Local] {
				return false
			}
		}
	}
	return true
}
