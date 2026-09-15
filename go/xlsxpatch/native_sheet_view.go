package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"math"
	"strconv"
	"strings"
)

// NativeWorkbookSheetViewV1 is the read-only sheetView pane fact recorded as
// authored: frozen row/column counts (or split twips), the bottom-right pane's
// top-left cell and the active pane. Grid and print previews ignore view state;
// it is preview metadata, never a mutation target.
type NativeWorkbookSheetViewV1 struct {
	PaneState     string   `json:"pane_state"`
	FrozenRows    int      `json:"frozen_rows"`
	FrozenColumns int      `json:"frozen_columns"`
	SplitXTwips   *float64 `json:"split_x_twips,omitempty"`
	SplitYTwips   *float64 `json:"split_y_twips,omitempty"`
	TopLeftCell   *string  `json:"top_left_cell,omitempty"`
	ActivePane    *string  `json:"active_pane,omitempty"`
}

// NativeWorkbookSheetViewV2 mirrors NativeWorkbookSheetViewV1 on the v2 wire.
type NativeWorkbookSheetViewV2 struct {
	PaneState     string   `json:"pane_state"`
	FrozenRows    int      `json:"frozen_rows"`
	FrozenColumns int      `json:"frozen_columns"`
	SplitXTwips   *float64 `json:"split_x_twips,omitempty"`
	SplitYTwips   *float64 `json:"split_y_twips,omitempty"`
	TopLeftCell   *string  `json:"top_left_cell,omitempty"`
	ActivePane    *string  `json:"active_pane,omitempty"`
}

const (
	nativeSheetViewMaxSelections   = 4
	nativeSheetViewMaxSelectionRef = 16
	nativeSheetViewMaxSplitTwips   = 2147483647
)

var nativePaneStates = map[string]bool{"frozen": true, "frozenSplit": true, "split": true}
var nativeActivePanes = map[string]bool{"topLeft": true, "topRight": true, "bottomLeft": true, "bottomRight": true}

// nativeSheetViewsResult is the outcome of reading <sheetViews>. benign means
// the whole element carried only UI state that read-only previews may ignore;
// view is the typed pane fact when a benign <pane> was present.
type nativeSheetViewsResult struct {
	benign bool
	view   *NativeWorkbookSheetViewV1
}

// parseNativeSheetViews reads one fully parsed <sheetViews>. Exactly one
// sheetView whose attributes are the required workbookViewId, tabSelected, or
// explicit Excel-default values is benign; a <pane> and up to four
// <selection> children are UI state recorded as a typed fact. Any other
// option, value, or markup keeps the element source-authoritative
// (SHEET_VIEW_GEOMETRY) exactly as before.
func parseNativeSheetViews(decoder *xml.Decoder, root xml.StartElement) (nativeSheetViewsResult, error) {
	result := nativeSheetViewsResult{benign: len(unexpectedSemanticXMLAttributes(root)) == 0}
	depth, sheetViews, panes, selections := 0, 0, 0, 0
	for {
		token, err := decoder.Token()
		if err != nil {
			return nativeSheetViewsResult{}, err
		}
		switch token := token.(type) {
		case xml.StartElement:
			switch {
			case depth == 0 && token.Name == (xml.Name{Space: root.Name.Space, Local: "sheetView"}):
				sheetViews++
				if sheetViews != 1 || !nativeBenignSheetViewAttributes(token) {
					result.benign = false
				}
			case depth == 1 && token.Name == (xml.Name{Space: root.Name.Space, Local: "pane"}):
				panes++
				view, ok := parseNativeSheetViewPane(token)
				if panes != 1 || selections != 0 || !ok {
					result.benign = false
				} else {
					result.view = view
				}
			case depth == 1 && token.Name == (xml.Name{Space: root.Name.Space, Local: "selection"}):
				selections++
				if selections > nativeSheetViewMaxSelections || !nativeBenignSelectionAttributes(token) {
					result.benign = false
				}
			default:
				result.benign = false
			}
			depth++
		case xml.EndElement:
			if depth == 0 {
				if token.Name != root.Name {
					return nativeSheetViewsResult{}, fmt.Errorf("sheetViews has a mismatched closing element")
				}
				if !result.benign || sheetViews != 1 {
					return nativeSheetViewsResult{benign: false}, nil
				}
				return result, nil
			}
			depth--
		case xml.CharData:
			if len(bytes.TrimSpace(token)) != 0 {
				result.benign = false
			}
		case xml.ProcInst:
			return nativeSheetViewsResult{}, fmt.Errorf("sheetViews contains a processing instruction")
		case xml.Directive:
			return nativeSheetViewsResult{}, fmt.Errorf("sheetViews contains an XML directive")
		}
	}
}

// nativeBenignSheetViewAttributes accepts workbookViewId, tabSelected, and
// explicit Excel-default view options only. Excel default is gridlines on;
// native chrome already draws the grid, so only an explicit off (or any other
// non-default option such as zoom, scroll position, or a page view) is view
// geometry that we do not project.
func nativeBenignSheetViewAttributes(element xml.StartElement) bool {
	allowed := []xml.Name{
		{Local: "workbookViewId"}, {Local: "tabSelected"},
		{Local: "showGridLines"}, {Local: "showRowColHeaders"}, {Local: "showZeros"}, {Local: "showRuler"},
		{Local: "showOutlineSymbols"}, {Local: "defaultGridColor"}, {Local: "showWhiteSpace"},
		{Local: "showFormulas"}, {Local: "rightToLeft"}, {Local: "windowProtection"},
		{Local: "view"}, {Local: "topLeftCell"}, {Local: "colorId"},
		{Local: "zoomScale"}, {Local: "zoomScaleNormal"}, {Local: "zoomScaleSheetLayoutView"}, {Local: "zoomScalePageLayoutView"},
	}
	if len(unexpectedSemanticXMLAttributes(element, allowed...)) != 0 {
		return false
	}
	viewID, found, err := unqualifiedXMLAttribute(element, "workbookViewId")
	if err != nil || !found {
		return false
	}
	if _, err = strconv.ParseUint(viewID, 10, 32); err != nil {
		return false
	}
	selected, found, err := unqualifiedXMLAttribute(element, "tabSelected")
	if err != nil || (found && selected != "0" && selected != "1" && selected != "false" && selected != "true") {
		return false
	}
	for _, local := range []string{"showGridLines", "showRowColHeaders", "showZeros", "showRuler", "showOutlineSymbols", "defaultGridColor", "showWhiteSpace"} {
		if !nativeDefaultOnXMLFlag(element, local) {
			return false
		}
	}
	for _, local := range []string{"showFormulas", "rightToLeft", "windowProtection"} {
		if !nativeDefaultOffXMLFlag(element, local) {
			return false
		}
	}
	for _, option := range []struct {
		local  string
		values []string
	}{
		{"view", []string{"normal"}},
		{"topLeftCell", []string{"A1"}},
		{"colorId", []string{"64"}},
		{"zoomScale", []string{"100"}},
		{"zoomScaleNormal", []string{"0", "100"}},
		{"zoomScaleSheetLayoutView", []string{"0", "100"}},
		{"zoomScalePageLayoutView", []string{"0", "100"}},
	} {
		if !nativeDefaultValuedXMLAttribute(element, option.local, option.values...) {
			return false
		}
	}
	return true
}

func nativeDefaultOnXMLFlag(element xml.StartElement, local string) bool {
	value, found, err := unqualifiedXMLAttribute(element, local)
	return err == nil && (!found || value == "1" || value == "true")
}

func nativeDefaultValuedXMLAttribute(element xml.StartElement, local string, values ...string) bool {
	value, found, err := unqualifiedXMLAttribute(element, local)
	if err != nil {
		return false
	}
	if !found {
		return true
	}
	for _, candidate := range values {
		if value == candidate {
			return true
		}
	}
	return false
}

// parseNativeSheetViewPane records a fully understood <pane>. Frozen states
// carry integer row/column counts; a plain split carries twips. Anything else
// is not benign.
func parseNativeSheetViewPane(element xml.StartElement) (*NativeWorkbookSheetViewV1, bool) {
	if len(unexpectedSemanticXMLAttributes(element, xml.Name{Local: "xSplit"}, xml.Name{Local: "ySplit"}, xml.Name{Local: "topLeftCell"}, xml.Name{Local: "activePane"}, xml.Name{Local: "state"})) != 0 {
		return nil, false
	}
	state, found, err := unqualifiedXMLAttribute(element, "state")
	if err != nil {
		return nil, false
	}
	if !found {
		state = "split"
	}
	if !nativePaneStates[state] {
		return nil, false
	}
	view := &NativeWorkbookSheetViewV1{PaneState: state}
	frozen := state != "split"
	for _, axis := range []struct {
		local   string
		count   *int
		twips   **float64
		maximum uint64
	}{
		{"xSplit", &view.FrozenColumns, &view.SplitXTwips, uint64(excelMaxColumns)},
		{"ySplit", &view.FrozenRows, &view.SplitYTwips, uint64(excelMaxRows)},
	} {
		if frozen {
			value, _, err := optionalNativeUintAttribute(element, axis.local, axis.maximum)
			if err != nil {
				return nil, false
			}
			*axis.count = int(value)
			continue
		}
		raw, found, err := unqualifiedXMLAttribute(element, axis.local)
		if err != nil {
			return nil, false
		}
		if found {
			value, err := finiteNativeFloat(raw, 0, nativeSheetViewMaxSplitTwips)
			if err != nil {
				return nil, false
			}
			*axis.twips = nativeWorkbookFloat(value)
		}
	}
	if topLeft, found, err := unqualifiedXMLAttribute(element, "topLeftCell"); err != nil {
		return nil, false
	} else if found {
		if !nativeCanonicalCellReference(topLeft) {
			return nil, false
		}
		view.TopLeftCell = nativeWorkbookString(topLeft)
	}
	if active, found, err := unqualifiedXMLAttribute(element, "activePane"); err != nil {
		return nil, false
	} else if found {
		if !nativeActivePanes[active] {
			return nil, false
		}
		view.ActivePane = nativeWorkbookString(active)
	}
	return view, true
}

// nativeBenignSelectionAttributes accepts selection UI state: an optional pane
// name, a canonical active cell, its id, and a bounded canonical sqref list
// that contains the active cell.
func nativeBenignSelectionAttributes(element xml.StartElement) bool {
	if len(unexpectedSemanticXMLAttributes(element, xml.Name{Local: "pane"}, xml.Name{Local: "activeCell"}, xml.Name{Local: "activeCellId"}, xml.Name{Local: "sqref"})) != 0 {
		return false
	}
	pane, found, err := unqualifiedXMLAttribute(element, "pane")
	if err != nil || (found && !nativeActivePanes[pane]) {
		return false
	}
	if _, _, err := optionalNativeUintAttribute(element, "activeCellId", uint64(^uint32(0))); err != nil {
		return false
	}
	active, activeFound, err := unqualifiedXMLAttribute(element, "activeCell")
	if err != nil || (activeFound && !nativeCanonicalCellReference(active)) {
		return false
	}
	selection, selectionFound, err := unqualifiedXMLAttribute(element, "sqref")
	if err != nil {
		return false
	}
	if !selectionFound {
		return true
	}
	refs := strings.Split(selection, " ")
	if len(refs) == 0 || len(refs) > nativeSheetViewMaxSelectionRef {
		return false
	}
	activeRow, activeColumn := -1, -1
	if activeFound {
		activeRow, activeColumn, _ = parseCellReference(active)
	}
	contained := !activeFound
	for _, ref := range refs {
		canonical, err := canonicalNativeMergedRangeReference(ref)
		if err != nil || canonical != ref {
			return false
		}
		minimumRow, minimumColumn, maximumRow, maximumColumn, err := parseDimensionReference(canonical)
		if err != nil {
			return false
		}
		if activeFound && activeRow >= minimumRow && activeRow <= maximumRow && activeColumn >= minimumColumn && activeColumn <= maximumColumn {
			contained = true
		}
	}
	return contained
}

func nativeCanonicalCellReference(value string) bool {
	row, column, err := parseCellReference(value)
	return err == nil && cellReference(row, column) == value
}

// nativeSheetViewPaneMessage describes the recorded pane for the inventory.
func nativeSheetViewPaneMessage(view *NativeWorkbookSheetViewV1) string {
	detail := view.PaneState + " pane"
	if view.PaneState != "split" {
		detail += fmt.Sprintf(" (%d frozen rows, %d frozen columns)", view.FrozenRows, view.FrozenColumns)
	}
	if view.TopLeftCell != nil {
		detail += " with bottom-right top-left cell " + *view.TopLeftCell
	}
	return detail + " is recorded read-only; grid and print previews ignore view state and view-dependent mutation is refused"
}

type nativeSheetViewIssue struct{ code, path, message string }

// nativeSheetViewIssues validates a decoded sheet_view fact for both wire versions.
func nativeSheetViewIssues(view *NativeWorkbookSheetViewV1, path string) []nativeSheetViewIssue {
	if view == nil {
		return nil
	}
	var issues []nativeSheetViewIssue
	add := func(code, field, message string) {
		issues = append(issues, nativeSheetViewIssue{code: code, path: path + "/sheet_view/" + field, message: message})
	}
	if !nativePaneStates[view.PaneState] {
		add("INVALID_VALUE", "pane_state", "pane state must be frozen, frozenSplit, or split")
	}
	if view.FrozenRows < 0 || view.FrozenRows > excelMaxRows {
		add("INVALID_NUMBER", "frozen_rows", "frozen rows are outside Excel bounds")
	}
	if view.FrozenColumns < 0 || view.FrozenColumns > excelMaxColumns {
		add("INVALID_NUMBER", "frozen_columns", "frozen columns are outside Excel bounds")
	}
	frozen := view.PaneState != "split"
	if frozen && (view.SplitXTwips != nil || view.SplitYTwips != nil) {
		add("INVALID_UNION", "split_x_twips", "frozen panes carry counts, not split twips")
	}
	if !frozen && (view.FrozenRows != 0 || view.FrozenColumns != 0) {
		add("INVALID_UNION", "frozen_rows", "split panes carry twips, not frozen counts")
	}
	for _, twips := range []struct {
		field string
		value *float64
	}{{"split_x_twips", view.SplitXTwips}, {"split_y_twips", view.SplitYTwips}} {
		if twips.value != nil && (math.IsNaN(*twips.value) || math.IsInf(*twips.value, 0) || *twips.value < 0 || *twips.value > nativeSheetViewMaxSplitTwips || (*twips.value == 0 && math.Signbit(*twips.value))) {
			add("INVALID_NUMBER", twips.field, "split twips must be finite, non-negative, and canonical")
		}
	}
	if view.TopLeftCell != nil && !nativeCanonicalCellReference(*view.TopLeftCell) {
		add("INVALID_VALUE", "top_left_cell", "top-left cell must be a canonical bounded A1 reference")
	}
	if view.ActivePane != nil && !nativeActivePanes[*view.ActivePane] {
		add("INVALID_VALUE", "active_pane", "active pane must name one of the four panes")
	}
	return issues
}
