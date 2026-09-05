package pptxpatch

import (
	"bytes"
	"fmt"
	"strings"
	"testing"
)

func TestExtractNativePPTXExactTablesTransitionalAndStrict(t *testing.T) {
	t.Parallel()
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			t.Parallel()
			cells := [][]string{
				{nativeExactTableCellXML("Alpha", "l", "FFF2CC"), nativeExactTableCellXML("Beta", "ctr", "DDEBF7")},
				{nativeExactTableCellXML("Gamma", "r", "E2F0D9"), nativeExactTableCellXML("Delta", "l", "FCE4D6")},
			}
			table := nativeExactTableGraphicFrameXML(3, "Native table", []int64{600000, 400000}, []int64{300000, 300000}, cells, "")
			sentinel := nativeAutoShapeXML(4, "After table", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
			deck, err := ExtractNativePPTX(nativeTableFixture(t, strict, table+sentinel), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract exact table: %v", err)
			}
			if len(deck.Slides) != 1 || len(deck.Slides[0].Elements) != 3 {
				t.Fatalf("table z-order was not retained: %#v", deck.Slides)
			}
			if deck.Slides[0].Elements[0].Kind != NativeElementKindText || deck.Slides[0].Elements[1].Kind != NativeElementKindTable || deck.Slides[0].Elements[2].Kind != NativeElementKindShape {
				t.Fatalf("table source order changed: %#v", deck.Slides[0].Elements)
			}
			element := deck.Slides[0].Elements[1]
			if element.Source == nil || element.Source.ObjectID != "cNvPr-3" || element.Compatibility.Status != NativeCompatibilityStatusEditable || len(element.Passthrough) != 0 || element.Table == nil {
				t.Fatalf("exact table was not projected editable: %#v", element)
			}
			if *element.Transform.X != 300000 || *element.Transform.Y != 150000 || *element.Transform.Cx != 1000000 || *element.Transform.Cy != 600000 {
				t.Fatalf("table transform changed: %#v", element.Transform)
			}
			if fmt.Sprint(element.Table.ColumnWidths) != "[600000 400000]" || fmt.Sprint(element.Table.RowHeights) != "[300000 300000]" || len(element.Table.Rows) != 2 || len(element.Table.Rows[0]) != 2 {
				t.Fatalf("table grid changed: %#v", element.Table)
			}
			cell := element.Table.Rows[0][1]
			if cell.Text == nil || *cell.Text != "Beta" || cell.Paragraphs == nil || len(*cell.Paragraphs) != 1 || cell.TextBody == nil || cell.Align != nil || cell.Border != nil || cell.Fill == nil || *cell.Fill != "DDEBF7" {
				t.Fatalf("native cell authority was incomplete: %#v", cell)
			}
			paragraph := (*cell.Paragraphs)[0]
			if paragraph.Align == nil || *paragraph.Align != NativeTextAlignCenter || paragraph.Level == nil || *paragraph.Level != 0 || paragraph.Bullet == nil || *paragraph.Bullet || len(paragraph.Runs) != 1 {
				t.Fatalf("native cell paragraph changed: %#v", paragraph)
			}
			run := paragraph.Runs[0]
			if run.Text == nil || *run.Text != "Beta" || run.FontFamily == nil || *run.FontFamily != "Aptos" || run.FontSizeHundredthPt == nil || *run.FontSizeHundredthPt != 1200 || run.Color == nil || *run.Color != "112233" || run.Bold == nil || *run.Bold || run.Italic == nil || *run.Italic {
				t.Fatalf("native cell run changed: %#v", run)
			}
			if *cell.TextBody.LeftInsetEMU != 10000 || *cell.TextBody.RightInsetEMU != 10000 || *cell.TextBody.TopInsetEMU != 5000 || *cell.TextBody.BottomInsetEMU != 5000 || cell.TextBody.Wrap != NativeTextWrapSquare || cell.TextBody.VerticalAnchor != NativeTextVerticalAnchorTop {
				t.Fatalf("native cell layout changed: %#v", cell.TextBody)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid exact table deck: %#v", issues)
			}
			encoded, err := MarshalNativePPTXJSON(deck)
			if err != nil {
				t.Fatalf("marshal exact table: %v", err)
			}
			decoded, err := DecodeNativePPTXJSON(encoded)
			if err != nil || len(ValidateNativePPTX(decoded)) != 0 {
				t.Fatalf("native table JSON round trip: err=%v issues=%#v", err, ValidateNativePPTX(decoded))
			}
		})
	}
}

func TestExtractNativePPTXGroupedTableRetainsAffineAndOrder(t *testing.T) {
	t.Parallel()
	table := nativeExactTableGraphicFrameXML(4, "Grouped table", []int64{500000, 500000}, []int64{500000}, [][]string{{
		nativeExactTableCellXML("One", "l", "FFFFFF"), nativeExactTableCellXML("Two", "r", "EEEEEE"),
	}}, "")
	shape := nativeAutoShapeXML(5, "Grouped sentinel", "rect", `<a:noFill/>`, nativeAutoShapeNoLine("flat", `<a:round/>`), "")
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Table group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="600" cy="800"/><a:chOff x="100" y="200"/><a:chExt cx="300" cy="400"/></a:xfrm></p:grpSpPr>` + table + shape + `</p:grpSp>`
	deck, err := ExtractNativePPTX(nativeTableFixture(t, false, group), nativeAtomicTestExtractOptions())
	if err != nil {
		t.Fatalf("extract grouped table: %v", err)
	}
	var projected *NativeElement
	for index := range deck.Slides[0].Elements {
		if deck.Slides[0].Elements[index].Kind == NativeElementKindGroup {
			projected = &deck.Slides[0].Elements[index]
		}
	}
	if projected == nil || projected.ChildTransform == nil || len(projected.Children) != 2 || projected.Children[0].Kind != NativeElementKindTable || projected.Children[1].Kind != NativeElementKindShape {
		t.Fatalf("grouped table/order was not projected: %#v", projected)
	}
	if *projected.Transform.X != 1000 || *projected.Transform.Y != 2000 || *projected.Transform.Cx != 600 || *projected.Transform.Cy != 800 || *projected.ChildTransform.X != 100 || *projected.ChildTransform.Y != 200 || *projected.ChildTransform.Cx != 300 || *projected.ChildTransform.Cy != 400 {
		t.Fatalf("group affine changed: %#v", projected)
	}
	child := projected.Children[0]
	if *child.Transform.X != 400000 || *child.Transform.Y != 200000 || *child.Transform.Cx != 1000000 || *child.Transform.Cy != 500000 || child.Table == nil || child.Compatibility.Status != NativeCompatibilityStatusEditable {
		t.Fatalf("grouped table local geometry changed: %#v", child)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		t.Fatalf("invalid grouped table deck: %#v", issues)
	}
}

func TestExtractNativePPTXTableGapsAreOpaqueAndExactPreserved(t *testing.T) {
	t.Parallel()
	base := nativeExactTableGraphicFrameXML(3, "Refused table", []int64{500000, 500000}, []int64{500000}, [][]string{{
		nativeExactTableCellXML("One", "l", "FFFFFF"), nativeExactTableCellXML("Two", "r", "EEEEEE"),
	}}, "")
	cases := []struct {
		name     string
		mutate   func(string) string
		wantCode string
	}{
		{name: "merge", mutate: func(value string) string { return strings.Replace(value, `<a:tc>`, `<a:tc gridSpan="2">`, 1) }, wantCode: "pptx.table-merge-unavailable"},
		{name: "table-style", mutate: func(value string) string {
			return strings.Replace(value, `<a:tblPr/>`, `<a:tblPr firstRow="1"><a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId></a:tblPr>`, 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "omitted-table-properties", mutate: func(value string) string {
			return strings.Replace(value, `<a:tblPr/>`, ``, 1)
		}, wantCode: "pptx.table-style-unavailable"},
		{name: "omitted-transform-offset", mutate: func(value string) string {
			return strings.Replace(value, `<a:off x="300000" y="150000"/>`, ``, 1)
		}, wantCode: "pptx.table-transform-unavailable"},
		{name: "omitted-cell-properties", mutate: func(value string) string {
			return removeFirstNativeTableXMLSpan(value, `<a:tcPr`, `</a:tcPr>`)
		}, wantCode: "pptx.table-cell-layout-unavailable"},
		{name: "omitted-cell-text-body", mutate: func(value string) string {
			return removeFirstNativeTableXMLSpan(value, `<a:txBody`, `</a:txBody>`)
		}, wantCode: "pptx.table-text-unavailable"},
		{name: "omitted-paragraph-properties", mutate: func(value string) string {
			return strings.Replace(value, `<a:pPr algn="l" lvl="0"><a:buNone/></a:pPr>`, ``, 1)
		}, wantCode: "pptx.table-text-unavailable"},
		{name: "omitted-run-properties", mutate: func(value string) string {
			return removeFirstNativeTableXMLSpan(value, `<a:rPr`, `</a:rPr>`)
		}, wantCode: "pptx.table-text-unavailable"},
		{name: "omitted-border", mutate: func(value string) string {
			return strings.Replace(value, `<a:lnL><a:noFill/></a:lnL>`, ``, 1)
		}, wantCode: "pptx.table-border-unavailable"},
		{name: "empty-border", mutate: func(value string) string {
			return strings.Replace(value, `<a:lnL><a:noFill/></a:lnL>`, `<a:lnL/>`, 1)
		}, wantCode: "pptx.table-border-unavailable"},
		{name: "empty-grid", mutate: func(value string) string {
			return strings.ReplaceAll(value, `<a:gridCol w="500000"/>`, ``)
		}, wantCode: "pptx.table-grid-unavailable"},
		{name: "empty-table", mutate: func(value string) string {
			return removeFirstNativeTableXMLSpan(value, `<a:tr`, `</a:tr>`)
		}, wantCode: "pptx.table-row-unavailable"},
		{name: "empty-paragraph", mutate: func(value string) string {
			return removeFirstNativeTableXMLSpan(value, `<a:r>`, `</a:r>`)
		}, wantCode: "pptx.table-text-unavailable"},
		{name: "theme-fill", mutate: func(value string) string {
			return strings.Replace(value, `<a:srgbClr val="FFFFFF"/>`, `<a:schemeClr val="accent1"/>`, 1)
		}, wantCode: "pptx.table-fill-unavailable"},
		{name: "visible-border", mutate: func(value string) string {
			return strings.Replace(value, `<a:lnL><a:noFill/></a:lnL>`, `<a:lnL w="12700"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:lnL>`, 1)
		}, wantCode: "pptx.table-border-unavailable"},
		{name: "body-effect", mutate: func(value string) string {
			return strings.Replace(value, `<a:bodyPr/>`, `<a:bodyPr><a:scene3d/></a:bodyPr>`, 1)
		}, wantCode: "pptx.table-cell-layout-unavailable"},
		{name: "chart-frame", mutate: func(value string) string {
			return strings.Replace(value, nativeDrawingTableURI, "http://schemas.openxmlformats.org/drawingml/2006/chart", 1)
		}, wantCode: "pptx.chart-markup-unavailable"},
	}
	for _, test := range cases {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			raw := test.mutate(base)
			requests := []NativePassthroughTokenRequest{}
			options := NativePPTXExtractOptions{TokenFactory: NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
				requests = append(requests, request)
				return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
			})}
			deck, err := ExtractNativePPTX(nativeTableFixture(t, false, raw), options)
			if err != nil {
				t.Fatalf("extract refused table: %v", err)
			}
			for _, element := range deck.Slides[0].Elements {
				if element.Kind == NativeElementKindTable {
					t.Fatalf("unsupported table leaked a partial projection: %#v", element)
				}
			}
			if deck.Slides[0].Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatalf("unsupported table did not preserve the slide: %#v", deck.Slides[0].Compatibility)
			}
			foundDiagnostic := false
			for _, diagnostic := range deck.Slides[0].Compatibility.Diagnostics {
				if diagnostic.Code == test.wantCode {
					foundDiagnostic = true
				}
			}
			if !foundDiagnostic {
				t.Fatalf("missing refusal diagnostic %q: %#v", test.wantCode, deck.Slides[0].Compatibility.Diagnostics)
			}
			var captured *NativePassthroughTokenRequest
			for index := range requests {
				if requests[index].ObjectID == "cNvPr-3" && requests[index].Reason == test.wantCode {
					captured = &requests[index]
				}
			}
			if captured == nil || captured.OwnerPart != "relocated/slides/slide-a.xml" || captured.ByteLength != int64(len(raw)) || !bytes.Equal(captured.Payload, []byte(raw)) {
				t.Fatalf("opaque table capability did not bind exact raw bytes: %#v", captured)
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatalf("invalid preserved table deck: %#v", issues)
			}
		})
	}
}

func TestExtractNativePPTXGroupedLateTableRefusalPublishesOnlyGroupCapability(t *testing.T) {
	t.Parallel()
	line := nativeAutoShapeSolidLine("12700", "flat", `<a:round/>`, "123456")
	attached := nativeConnectorXML(4, "Attached connector", `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>`, line, "", `<a:stCxn id="2" idx="0"/>`, "", "")
	refused := nativeExactTableGraphicFrameXML(5, "Late styled table", []int64{1000000}, []int64{500000}, [][]string{{nativeExactTableCellXML("Late", "l", "FFFFFF")}}, "")
	refused = strings.Replace(refused, `<a:tblPr/>`, `<a:tblPr bandRow="1"><a:tableStyleId>{5940675A-B579-460E-94D1-54222C63F5DA}</a:tableStyleId></a:tblPr>`, 1)
	group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Atomic table group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>` + attached + refused + `</p:grpSp>`
	requests := []NativePassthroughTokenRequest{}
	issuer := nativeAtomicTestTokenFactory(NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
		requests = append(requests, request)
		return "token-" + nativeSHA256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))[:24], nil
	}))
	deck, err := ExtractNativePPTX(nativeTableFixture(t, false, group), NativePPTXExtractOptions{TokenFactory: issuer})
	if err != nil {
		t.Fatalf("late table refusal: %v", err)
	}
	groupRequests := []NativePassthroughTokenRequest{}
	for _, request := range requests {
		if request.OwnerPart == "relocated/slides/slide-a.xml" && strings.HasPrefix(request.ObjectID, "cNvPr-") {
			groupRequests = append(groupRequests, request)
		}
	}
	if len(groupRequests) != 1 || groupRequests[0].ObjectID != "cNvPr-3" || groupRequests[0].Reason != "pptx.group-table-unavailable" || !strings.HasPrefix(string(groupRequests[0].Payload), `<p:grpSp`) {
		t.Fatalf("external issuer observed an orphan child capability: %#v", groupRequests)
	}
	for _, request := range requests {
		if strings.HasPrefix(request.Reason, "pptx.connector-") || strings.HasPrefix(request.Reason, "pptx.table-") {
			t.Fatalf("nested capability escaped atomic table-group refusal: %#v", requests)
		}
	}
	for _, element := range deck.Slides[0].Elements {
		if element.Kind == NativeElementKindGroup || element.Kind == NativeElementKindTable || element.Kind == NativeElementKindConnector {
			t.Fatalf("opaque table group leaked a partial projection: %#v", deck.Slides[0].Elements)
		}
	}
}

func TestExtractNativePPTXTableMalformedOrMixedDialectFailsWithoutProjection(t *testing.T) {
	t.Parallel()
	base := nativeExactTableGraphicFrameXML(3, "Malformed table", []int64{1000000}, []int64{500000}, [][]string{{nativeExactTableCellXML("One", "l", "FFFFFF")}}, "")
	duplicateTransform := strings.Replace(base, `<p:xfrm>`, `<p:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></p:xfrm><p:xfrm>`, 1)
	zeroExtent := strings.Replace(base, `cx="1000000" cy="500000"`, `cx="0" cy="500000"`, 1)
	mixedDialect := strings.Replace(base, `<a:tbl>`, fmt.Sprintf(`<s:tbl xmlns:s="%s">`, nsDrawingStrict), 1)
	mixedDialect = strings.Replace(mixedDialect, `</a:tbl>`, `</s:tbl>`, 1)
	reorderedTransform := strings.Replace(base, `<a:off x="300000" y="150000"/><a:ext cx="1000000" cy="500000"/>`, `<a:ext cx="1000000" cy="500000"/><a:off x="300000" y="150000"/>`, 1)
	reorderedNonVisual := strings.Replace(base, `<p:cNvPr id="3" name="Malformed table"/><p:cNvGraphicFramePr/>`, `<p:cNvGraphicFramePr/><p:cNvPr id="3" name="Malformed table"/>`, 1)
	reorderedTextBody := strings.Replace(base, `<a:bodyPr/><a:lstStyle/>`, `<a:lstStyle/><a:bodyPr/>`, 1)
	reorderedParagraph := strings.Replace(base, `<a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r>`, `<a:r><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>One</a:t></a:r><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r>`, 1)
	reorderedRun := strings.Replace(base, `<a:r><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>One</a:t></a:r>`, `<a:r><a:t>One</a:t><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr></a:r>`, 1)
	reorderedRunProperties := strings.Replace(base, `<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/>`, `<a:latin typeface="Aptos"/><a:solidFill><a:srgbClr val="112233"/></a:solidFill>`, 1)
	reorderedCellProperties := strings.Replace(base, `<a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR>`, `<a:lnR><a:noFill/></a:lnR><a:lnL><a:noFill/></a:lnL>`, 1)
	for name, value := range map[string]string{
		"duplicate-transform":  duplicateTransform,
		"zero-extent":          zeroExtent,
		"mixed-dialect":        mixedDialect,
		"reordered-transform":  reorderedTransform,
		"reordered-nonvisual":  reorderedNonVisual,
		"reordered-text-body":  reorderedTextBody,
		"reordered-paragraph":  reorderedParagraph,
		"reordered-run":        reorderedRun,
		"reordered-run-props":  reorderedRunProperties,
		"reordered-cell-props": reorderedCellProperties,
	} {
		name, value := name, value
		t.Run(name, func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeTableFixture(t, false, value), nativeTestExtractOptions())
			if err == nil || len(deck.Slides) != 0 {
				t.Fatalf("malformed table produced a partial deck: err=%v deck=%#v", err, deck)
			}
		})
	}
}

func TestExtractNativePPTXHiddenTablePreservesFrameOrOwningGroup(t *testing.T) {
	t.Parallel()
	table := nativeExactTableGraphicFrameXML(4, "Hidden table", []int64{1000000}, []int64{500000}, [][]string{{nativeExactTableCellXML("Hidden", "l", "FFFFFF")}}, "")
	table = strings.Replace(table, `name="Hidden table"`, `name="Hidden table" hidden="1"`, 1)

	for _, test := range []struct {
		name, children, objectID, reason, rawPrefix string
	}{
		{name: "top-level", children: table, objectID: "cNvPr-4", reason: "pptx.table-nonvisual-unavailable", rawPrefix: `<p:graphicFrame`},
		{name: "grouped", children: `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Hidden table group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>` + table + `</p:grpSp>`, objectID: "cNvPr-3", reason: "pptx.group-hidden-child-unavailable", rawPrefix: `<p:grpSp`},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			requests := []NativePassthroughTokenRequest{}
			issuer := nativeAtomicTestTokenFactory(NativePassthroughTokenFactoryFunc(func(request NativePassthroughTokenRequest) (string, error) {
				requests = append(requests, request)
				return "token-" + nativeSHA256([]byte(request.ObjectID + "\x00" + request.Reason))[:24], nil
			}))
			deck, err := ExtractNativePPTX(nativeTableFixture(t, false, test.children), NativePPTXExtractOptions{TokenFactory: issuer})
			if err != nil {
				t.Fatalf("extract hidden table: %v", err)
			}
			for _, element := range deck.Slides[0].Elements {
				if element.Kind == NativeElementKindTable || element.Kind == NativeElementKindGroup {
					t.Fatalf("hidden table leaked native projection: %#v", element)
				}
			}
			found := false
			for _, request := range requests {
				if request.ObjectID == test.objectID && request.Reason == test.reason && strings.HasPrefix(string(request.Payload), test.rawPrefix) {
					found = true
				}
			}
			if !found {
				t.Fatalf("hidden table did not preserve exact owning subtree: %#v", requests)
			}
		})
	}
}

func TestReserveNativeTableOutputIsAtomicAtCumulativeLimits(t *testing.T) {
	t.Parallel()
	text := "x"
	paragraphs := []NativeParagraph{{Runs: []NativeTextRun{{Text: &text}}}}
	exact := nativeExactTable{
		table:       NativeTable{ColumnWidths: []int64{1}, RowHeights: []int64{1}, Rows: [][]NativeTableCell{{{Text: &text, Paragraphs: &paragraphs}}}},
		outputNodes: 32, textCodeUnits: 1,
	}
	for _, test := range []struct {
		name      string
		extractor nativeExtractor
	}{
		{name: "nodes", extractor: nativeExtractor{outputNodesEmitted: nativeMaxNodes - exact.outputNodes + 1, textCodeUnitsEmitted: 7, tableCellsEmitted: 11}},
		{name: "text", extractor: nativeExtractor{outputNodesEmitted: 7, textCodeUnitsEmitted: int64(nativeMaxTotalTextCodeUnits), tableCellsEmitted: 11}},
		{name: "cells", extractor: nativeExtractor{outputNodesEmitted: 7, textCodeUnitsEmitted: 11, tableCellsEmitted: nativeMaxTableCells}},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			extractor := test.extractor
			beforeNodes, beforeText, beforeCells := extractor.outputNodesEmitted, extractor.textCodeUnitsEmitted, extractor.tableCellsEmitted
			if err := extractor.reserveNativeTableOutput(exact); err == nil {
				t.Fatalf("cumulative table %s overflow was accepted", test.name)
			}
			if extractor.outputNodesEmitted != beforeNodes || extractor.textCodeUnitsEmitted != beforeText || extractor.tableCellsEmitted != beforeCells {
				t.Fatalf("failed table reserve changed extractor budgets: %#v", extractor)
			}
		})
	}
}

func TestNativeTableAuthorityRequiresCompleteExactFrameGeometry(t *testing.T) {
	t.Parallel()
	tableXML := nativeExactTableGraphicFrameXML(3, "Authority table", []int64{500000, 500000}, []int64{500000}, [][]string{{
		nativeExactTableCellXML("One", "l", "FFFFFF"), nativeExactTableCellXML("Two", "r", "EEEEEE"),
	}}, "")
	for _, test := range []struct {
		name, code string
		mutate     func(*NativeElement)
	}{
		{name: "missing-row-heights", code: "native.tableGeometry", mutate: func(element *NativeElement) { element.Table.RowHeights = []int64{} }},
		{name: "frame-track-mismatch", code: "native.tableGeometry", mutate: func(element *NativeElement) { *element.Transform.Cx++ }},
		{name: "mixed-authority", code: "native.tableTextAuthority", mutate: func(element *NativeElement) {
			element.Table.Rows[0][1].Paragraphs = nil
			element.Table.Rows[0][1].TextBody = nil
		}},
	} {
		test := test
		t.Run(test.name, func(t *testing.T) {
			t.Parallel()
			deck, err := ExtractNativePPTX(nativeTableFixture(t, false, tableXML), nativeTestExtractOptions())
			if err != nil {
				t.Fatalf("extract authority table: %v", err)
			}
			var element *NativeElement
			for index := range deck.Slides[0].Elements {
				if deck.Slides[0].Elements[index].Kind == NativeElementKindTable {
					element = &deck.Slides[0].Elements[index]
				}
			}
			if element == nil {
				t.Fatal("missing exact table")
			}
			test.mutate(element)
			found := false
			for _, issue := range ValidateNativePPTX(deck) {
				found = found || issue.Code == test.code
			}
			if !found {
				t.Fatalf("missing %s issue: %#v", test.code, ValidateNativePPTX(deck))
			}
		})
	}
}

func nativeTableFixture(t *testing.T, strict bool, children string) []byte {
	t.Helper()
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, mutate: func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, children+`</p:spTree>`, 1)
	}})
}

func nativeExactTableGraphicFrameXML(id int, name string, columns, heights []int64, rows [][]string, xfrmAttrs string) string {
	var grid strings.Builder
	var body strings.Builder
	var cx, cy int64
	for _, width := range columns {
		cx += width
		fmt.Fprintf(&grid, `<a:gridCol w="%d"/>`, width)
	}
	for rowIndex, height := range heights {
		cy += height
		fmt.Fprintf(&body, `<a:tr h="%d">%s</a:tr>`, height, strings.Join(rows[rowIndex], ""))
	}
	return fmt.Sprintf(`<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="%d" name="%s"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm%s><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></p:xfrm><a:graphic><a:graphicData uri="%s"><a:tbl><a:tblPr/><a:tblGrid>%s</a:tblGrid>%s</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`, id, name, xfrmAttrs, id*100000, id*50000, cx, cy, nativeDrawingTableURI, grid.String(), body.String())
}

func nativeExactTableCellXML(text, align, fill string) string {
	return fmt.Sprintf(`<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="%s" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="0" i="0" sz="1200"><a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>%s</a:t></a:r></a:p></a:txBody><a:tcPr marL="10000" marR="10000" marT="5000" marB="5000" anchor="t" anchorCtr="0" horzOverflow="overflow" vert="horz"><a:lnL><a:noFill/></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:noFill/></a:lnT><a:lnB><a:noFill/></a:lnB><a:solidFill><a:srgbClr val="%s"/></a:solidFill></a:tcPr></a:tc>`, align, text, fill)
}

func removeFirstNativeTableXMLSpan(value, startMarker, endMarker string) string {
	start := strings.Index(value, startMarker)
	if start < 0 {
		return value
	}
	relativeEnd := strings.Index(value[start:], endMarker)
	if relativeEnd < 0 {
		return value
	}
	end := start + relativeEnd + len(endMarker)
	return value[:start] + value[end:]
}
