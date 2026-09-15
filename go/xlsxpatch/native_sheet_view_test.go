package xlsxpatch

import (
	"reflect"
	"strings"
	"testing"
)

// Calc-authored sheetView: explicit Excel-default options, a frozen pane, and
// one selection per pane.
const nativeSheetViewCalcFixture = `<sheetViews><sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="true" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0">` +
	`<pane xSplit="5" ySplit="10" topLeftCell="A1" activePane="bottomRight" state="frozen"/>` +
	`<selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/><selection pane="topRight" activeCell="F1" activeCellId="0" sqref="F1"/>` +
	`<selection pane="bottomLeft" activeCell="A11" activeCellId="0" sqref="A11"/><selection pane="bottomRight" activeCell="F11" activeCellId="0" sqref="F11"/>` +
	`</sheetView></sheetViews>`

func nativeSheetViewFixtureParts(sheetViews string) map[string]string {
	parts := nativeWorkbookFixture(false)
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<dimension ref="A1:K2"/>`, `<dimension ref="A1:K2"/>`+sheetViews, 1)
	return parts
}

func nativeUnsupportedMessage(workbook *NativeWorkbookV1, code string) string {
	for _, item := range workbook.Unsupported {
		if item.Code == code {
			return item.Message
		}
	}
	return ""
}

func TestExtractNativeWorkbookV1RecordsFrozenPaneAsReadOnlyFact(t *testing.T) {
	raw := buildZip(t, nativeSheetViewFixtureParts(nativeSheetViewCalcFixture))
	workbook, err := ExtractNativeWorkbookV1(raw)
	if err != nil {
		t.Fatal(err)
	}
	if hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") {
		t.Fatalf("benign frozen pane became geometry authority: %#v", workbook.Unsupported)
	}
	if !hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_PANE") {
		t.Fatalf("frozen pane was not inventoried: %#v", workbook.Unsupported)
	}
	if message := nativeUnsupportedMessage(workbook, "SHEET_VIEW_PANE"); !strings.Contains(message, "10 frozen rows, 5 frozen columns") || !strings.Contains(message, "top-left cell A1") {
		t.Fatalf("pane diagnostic lost its facts: %q", message)
	}
	view := workbook.Sheets[0].SheetView
	if view == nil || view.PaneState != "frozen" || view.FrozenRows != 10 || view.FrozenColumns != 5 || view.SplitXTwips != nil || view.SplitYTwips != nil ||
		view.TopLeftCell == nil || *view.TopLeftCell != "A1" || view.ActivePane == nil || *view.ActivePane != "bottomRight" {
		t.Fatalf("typed sheet-view fact was not recorded: %#v", view)
	}
	if workbook.Sheets[1].SheetView != nil {
		t.Fatalf("sheet without sheetViews gained a pane fact: %#v", workbook.Sheets[1].SheetView)
	}
	if issues := ValidateNativeWorkbookV1(workbook); len(issues) != 0 {
		t.Fatalf("v1 contract with sheet_view is invalid: %v", issues)
	}
	v2, err := ExtractNativeWorkbookV2(raw)
	if err != nil {
		t.Fatal(err)
	}
	if v2.Sheets[0].SheetView == nil || v2.Sheets[0].SheetView.FrozenRows != 10 || v2.Sheets[0].SheetView.FrozenColumns != 5 || !nativeGetCorpusHasCode(v2.Unsupported, "SHEET_VIEW_PANE") {
		t.Fatalf("v2 lost the sheet-view fact: %#v", v2.Sheets[0].SheetView)
	}
	encoded, err := EncodeNativeWorkbookV2(v2)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(string(encoded), `"sheet_view":{"pane_state":"frozen","frozen_rows":10,"frozen_columns":5,"top_left_cell":"A1","active_pane":"bottomRight"}`) {
		t.Fatalf("sheet_view wire shape drifted: %s", encoded)
	}
	decoded, err := DecodeNativeWorkbookV2(encoded)
	if err != nil {
		t.Fatal(err)
	}
	if !reflect.DeepEqual(decoded.Sheets[0].SheetView, v2.Sheets[0].SheetView) {
		t.Fatalf("sheet_view did not round-trip: %#v", decoded.Sheets[0].SheetView)
	}
}

func TestExtractNativeWorkbookV1RecordsSplitPaneTwips(t *testing.T) {
	workbook, err := ExtractNativeWorkbookV1(buildZip(t, nativeSheetViewFixtureParts(`<sheetViews><sheetView workbookViewId="0"><pane xSplit="2000" ySplit="1500.5" topLeftCell="C4" activePane="bottomRight"/><selection pane="bottomRight" activeCell="C4" sqref="C4:D6 F8"/></sheetView></sheetViews>`)))
	if err != nil {
		t.Fatal(err)
	}
	view := workbook.Sheets[0].SheetView
	if hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") || !hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_PANE") || view == nil ||
		view.PaneState != "split" || view.FrozenRows != 0 || view.FrozenColumns != 0 || view.SplitXTwips == nil || *view.SplitXTwips != 2000 || view.SplitYTwips == nil || *view.SplitYTwips != 1500.5 ||
		view.TopLeftCell == nil || *view.TopLeftCell != "C4" {
		t.Fatalf("split pane was not recorded as twips: view=%#v unsupported=%#v", view, workbook.Unsupported)
	}
}

func TestExtractNativeWorkbookV1KeepsUnmodeledSheetViewsAsGeometry(t *testing.T) {
	pane := `<pane xSplit="1" ySplit="1" topLeftCell="B2" activePane="bottomRight" state="frozen"/>`
	for _, test := range []struct{ name, sheetViews string }{
		{"zoom", `<sheetViews><sheetView workbookViewId="0" zoomScale="80">` + pane + `</sheetView></sheetViews>`},
		{"page break view", `<sheetViews><sheetView workbookViewId="0" view="pageBreakPreview">` + pane + `</sheetView></sheetViews>`},
		{"scrolled view", `<sheetViews><sheetView workbookViewId="0" topLeftCell="C3">` + pane + `</sheetView></sheetViews>`},
		{"gridlines off", `<sheetViews><sheetView workbookViewId="0" showGridLines="0">` + pane + `</sheetView></sheetViews>`},
		{"right to left", `<sheetViews><sheetView workbookViewId="0" rightToLeft="1">` + pane + `</sheetView></sheetViews>`},
		{"unknown view attribute", `<sheetViews><sheetView workbookViewId="0" future="1">` + pane + `</sheetView></sheetViews>`},
		{"missing workbookViewId", `<sheetViews><sheetView>` + pane + `</sheetView></sheetViews>`},
		{"two sheet views", `<sheetViews><sheetView workbookViewId="0">` + pane + `</sheetView><sheetView workbookViewId="1"/></sheetViews>`},
		{"bogus pane state", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" state="locked"/></sheetView></sheetViews>`},
		{"fractional frozen count", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1.5" ySplit="1" state="frozen"/></sheetView></sheetViews>`},
		{"frozen count beyond Excel rows", `<sheetViews><sheetView workbookViewId="0"><pane ySplit="1048577" state="frozen"/></sheetView></sheetViews>`},
		{"negative split twips", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="-5" ySplit="1"/></sheetView></sheetViews>`},
		{"non-canonical pane top-left", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="$B$2" state="frozen"/></sheetView></sheetViews>`},
		{"unknown active pane", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" activePane="middle" state="frozen"/></sheetView></sheetViews>`},
		{"unknown pane attribute", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" state="frozen" future="1"/></sheetView></sheetViews>`},
		{"pane child markup", `<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" state="frozen"><extLst/></pane></sheetView></sheetViews>`},
		{"two panes", `<sheetViews><sheetView workbookViewId="0">` + pane + pane + `</sheetView></sheetViews>`},
		{"pane after selection", `<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/>` + pane + `</sheetView></sheetViews>`},
		{"five selections", `<sheetViews><sheetView workbookViewId="0">` + pane + strings.Repeat(`<selection activeCell="A1" sqref="A1"/>`, 5) + `</sheetView></sheetViews>`},
		{"selection outside sqref", `<sheetViews><sheetView workbookViewId="0">` + pane + `<selection activeCell="A1" sqref="B2"/></sheetView></sheetViews>`},
		{"selection unknown attribute", `<sheetViews><sheetView workbookViewId="0">` + pane + `<selection activeCell="A1" sqref="A1" future="1"/></sheetView></sheetViews>`},
		{"selection unknown pane", `<sheetViews><sheetView workbookViewId="0">` + pane + `<selection pane="middle" activeCell="A1" sqref="A1"/></sheetView></sheetViews>`},
		{"selection sqref list too long", `<sheetViews><sheetView workbookViewId="0">` + pane + `<selection activeCell="A1" sqref="` + strings.TrimSpace(strings.Repeat("A1 ", 17)) + `"/></sheetView></sheetViews>`},
		{"selection non-canonical sqref", `<sheetViews><sheetView workbookViewId="0">` + pane + `<selection activeCell="A1" sqref="a1"/></sheetView></sheetViews>`},
		{"pivot selection", `<sheetViews><sheetView workbookViewId="0">` + pane + `<pivotSelection pane="bottomRight"/></sheetView></sheetViews>`},
		{"sheet view text", `<sheetViews><sheetView workbookViewId="0">` + pane + `text</sheetView></sheetViews>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			workbook, err := ExtractNativeWorkbookV1(buildZip(t, nativeSheetViewFixtureParts(test.sheetViews)))
			if err != nil {
				t.Fatal(err)
			}
			if !hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_GEOMETRY") || hasNativeWorkbookUnsupported(workbook, "SHEET_VIEW_PANE") || workbook.Sheets[0].SheetView != nil {
				t.Fatalf("unmodeled sheet view did not stay source-authoritative: view=%#v unsupported=%#v", workbook.Sheets[0].SheetView, workbook.Unsupported)
			}
		})
	}
}

func TestExtractNativeWorkbookV1SheetViewsStillRejectMalformedMarkup(t *testing.T) {
	for _, test := range []struct{ name, sheetViews string }{
		{"processing instruction", `<sheetViews><sheetView workbookViewId="0"><?pi?></sheetView></sheetViews>`},
		{"directive", `<sheetViews><sheetView workbookViewId="0"><!DOCTYPE x></sheetView></sheetViews>`},
	} {
		t.Run(test.name, func(t *testing.T) {
			if _, err := ExtractNativeWorkbookV1(buildZip(t, nativeSheetViewFixtureParts(test.sheetViews))); err == nil {
				t.Fatal("malformed sheetViews markup was accepted")
			}
		})
	}
}

func TestNativeSheetViewIssuesRejectContradictoryFacts(t *testing.T) {
	valid := NativeWorkbookSheetViewV1{PaneState: "frozen", FrozenRows: 1, FrozenColumns: 1, TopLeftCell: nativeWorkbookString("B2"), ActivePane: nativeWorkbookString("bottomRight")}
	if issues := nativeSheetViewIssues(&valid, "/sheets/0"); len(issues) != 0 {
		t.Fatalf("valid fact rejected: %#v", issues)
	}
	if issues := nativeSheetViewIssues(nil, "/sheets/0"); len(issues) != 0 {
		t.Fatalf("absent fact produced issues: %#v", issues)
	}
	for name, view := range map[string]NativeWorkbookSheetViewV1{
		"unknown state":            {PaneState: "locked"},
		"frozen with twips":        {PaneState: "frozen", FrozenRows: 1, SplitXTwips: nativeWorkbookFloat(10)},
		"split with counts":        {PaneState: "split", FrozenRows: 1},
		"negative twips":           {PaneState: "split", SplitYTwips: nativeWorkbookFloat(-1)},
		"rows beyond Excel bounds": {PaneState: "frozen", FrozenRows: excelMaxRows + 1},
		"non-canonical top-left":   {PaneState: "frozen", TopLeftCell: nativeWorkbookString("b2")},
		"unknown active pane":      {PaneState: "frozen", ActivePane: nativeWorkbookString("middle")},
	} {
		t.Run(name, func(t *testing.T) {
			if issues := nativeSheetViewIssues(&view, "/sheets/0"); len(issues) == 0 {
				t.Fatalf("contradictory fact accepted: %#v", view)
			}
			workbook, err := ExtractNativeWorkbookV1(buildZip(t, nativeSheetViewFixtureParts(nativeSheetViewCalcFixture)))
			if err != nil {
				t.Fatal(err)
			}
			tampered := view
			workbook.Sheets[0].SheetView = &tampered
			if issues := ValidateNativeWorkbookV1(workbook); len(issues) == 0 {
				t.Fatal("v1 validation accepted a contradictory sheet_view")
			}
		})
	}
}
