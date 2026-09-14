package pptxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"io"
	"os"
	"path/filepath"
	"reflect"
	"regexp"
	"strings"
	"testing"
)

func TestNativeGraphicFrameSourceLayoutAndAuthority(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, mode := range []string{"ordinary", "rotation", "mismatch", "group"} {
			raw := nativeExactTableGraphicFrameXML(4, "Intrinsic", []int64{600000, 400000}, []int64{300000}, [][]string{{nativeExactTableCellXML("Left", "l", "FFFFFF"), nativeExactTableCellXML("Right", "r", "EEEEEE")}}, "")
			if mode == "rotation" {
				raw = strings.Replace(raw, "<p:xfrm>", `<p:xfrm rot="-1800000" flipH="1">`, 1)
			}
			if mode == "mismatch" {
				raw = strings.Replace(raw, `cx="1000000"`, `cx="1500000"`, 1)
			}
			if mode == "group" {
				raw = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Table group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000"><a:off x="1000" y="2000"/><a:ext cx="6000000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="4000000"/></a:xfrm></p:grpSpPr>` + raw + `</p:grpSp>`
			}
			payload := nativeTableFixture(t, strict, raw)
			before := append([]byte(nil), payload...)
			deck, err := ExtractNativePPTX(payload, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatal(mode, err)
			}
			var table *NativeElement
			var find func([]NativeElement)
			find = func(es []NativeElement) {
				for i := range es {
					if es[i].Kind == NativeElementKindTable {
						table = &es[i]
					}
					find(es[i].Children)
				}
			}
			find(deck.Slides[0].Elements)
			if table == nil {
				t.Fatalf("%s table refused: %#v", mode, deck.Slides[0].Compatibility)
			}
			if !bytes.Equal(payload, before) || table.Table.ColumnWidths[0] != 600000 || table.Table.ColumnWidths[1] != 400000 || table.Table.RowHeights[0] != 300000 {
				t.Fatal("source/intrinsic changed")
			}
			if issues := ValidateNativePPTX(deck); len(issues) != 0 {
				t.Fatal(mode, issues)
			}
			if mode == "ordinary" {
				if table.GraphicFrameLayout != nil || table.Compatibility.Status != NativeCompatibilityStatusEditable {
					t.Fatal("ordinary changed")
				}
				continue
			}
			if table.GraphicFrameLayout == nil || *table.GraphicFrameLayout != nativeSourceAnchoredGraphicFrame || table.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
				t.Fatal("profile missing")
			}
			if mode == "rotation" && (table.Transform.RotationAngle == nil || *table.Transform.RotationAngle != 19800000 || table.Transform.FlipH == nil || !*table.Transform.FlipH) {
				t.Fatal("orientation lost")
			}
			if mode == "mismatch" && *table.Transform.Cx != 1500000 {
				t.Fatal("frame resized")
			}
			replacement := []NativeParagraph{{Runs: []NativeTextRun{{Text: stringPointer("Changed")}}}}
			_, err = resolveNativePPTXMutations(deck, []NativePPTXMutation{{OperationID: "edit", Kind: NativePPTXReplaceText, ElementID: table.ID, ExpectedFingerprintSHA256: table.Source.FingerprintSHA256, Paragraphs: &replacement}})
			if err == nil || !strings.Contains(err.Error(), "affine") {
				t.Fatalf("mutation guard: %v", err)
			}
			table.Compatibility.Status = NativeCompatibilityStatusEditable
			if len(ValidateNativePPTX(deck)) == 0 {
				t.Fatal("editable profile")
			}
		}
	}
}

func TestNativeGraphicFrameExactAncestorMutationIsAtomic(t *testing.T) {
	for _, strict := range []bool{false, true} {
		frame := nativeExactTableGraphicFrameXML(4, "Table", []int64{1000000}, []int64{300000}, [][]string{{nativeExactTableCellXML("Text", "l", "FFFFFF")}}, "")
		group := `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Exact scaled ancestor"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="8000000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="4000000"/></a:xfrm></p:grpSpPr>` + frame + `</p:grpSp>`
		original := nativeTableFixture(t, strict, group)
		before := append([]byte(nil), original...)
		deck, err := ExtractNativePPTX(original, nativeMutationExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		var ancestor, text NativeElement
		for _, element := range deck.Slides[0].Elements {
			if element.Kind == NativeElementKindGroup {
				ancestor = element
			}
			if element.Kind == NativeElementKindText {
				text = element
			}
		}
		if ancestor.Source == nil || text.Source == nil || nativeComplexAffineGroup(ancestor) {
			t.Fatal("fixture must exercise an exact legacy ancestor and an independently mutable text")
		}
		replacement := nativeMutationParagraphs("Changed atomically")
		valid := NativePPTXMutation{OperationID: "first", Kind: NativePPTXReplaceText, ElementID: text.ID, ExpectedFingerprintSHA256: text.Source.FingerprintSHA256, Paragraphs: &replacement}
		if _, err := resolveNativePPTXMutations(deck, []NativePPTXMutation{valid}); err != nil {
			t.Fatal("first operation must be valid", err)
		}
		for _, kind := range []NativePPTXMutationKind{NativePPTXReplaceText, NativePPTXUpdateAutoShape} {
			invalid := NativePPTXMutation{OperationID: "second", Kind: kind, ElementID: ancestor.ID, ExpectedFingerprintSHA256: ancestor.Source.FingerprintSHA256}
			expected := "not text-bearing"
			if kind == NativePPTXReplaceText {
				invalid.Paragraphs = &replacement
			} else {
				invalid.AutoShape = &NativePPTXAutoShapeMutation{Transform: ancestor.Children[0].Transform, Preset: NativeShapePresetRect}
				expected = "not an exact native AutoShape"
			}
			for _, operations := range [][]NativePPTXMutation{{invalid}, {valid, invalid}} {
				output, err := ApplyNativePPTXMutations(original, NativePPTXMutationRequest{ExpectedSourceRevision: *deck.SourceRevision, Operations: operations})
				if err == nil || !strings.Contains(err.Error(), expected) || len(output) != 0 || !bytes.Equal(original, before) {
					t.Fatalf("ancestor mutation must refuse without any output: %s %v", kind, err)
				}
			}
		}
	}
}

func TestNativeGraphicFramePolicyDiagnosticAtCapacity(t *testing.T) {
	element := NativeElement{}
	element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	for i := 0; i < nativeMaxDiagnosticsPerScope; i++ {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Code: "pptx.prior", Severity: NativeDiagnosticSeverityWarning, Message: "Prior detail"})
	}
	before := append([]NativeDiagnostic(nil), element.Compatibility.Diagnostics...)
	err := nativeMarkGraphicFrameLayout(&element)
	if err == nil || !strings.Contains(err.Error(), "diagnostic scope") || element.GraphicFrameLayout != nil || !reflect.DeepEqual(before, element.Compatibility.Diagnostics) {
		t.Fatal("full scope must refuse without changing any existing diagnostic or attaching a profile", err)
	}
}

func TestNativeGraphicFrameBrowserFixtures(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_GRAPHIC_FRAME_FIXTURE_DIR")
	if dir == "" {
		t.Skip("external source/WASM/browser proof")
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	for _, name := range []string{"table-intrinsic", "table-rotated-group", "table-y-overflow", "table-above-slide", "table-unqualified-x-clip"} {
		cell := strings.ReplaceAll(nativeExactTableCellXML("fgj ABC ABC ABC ABC ABC", "l", "EEEEEE"), `typeface="Aptos"`, `typeface="DejaVu Sans"`)
		if name == "table-unqualified-x-clip" {
			cell = strings.Replace(cell, `horzOverflow="overflow"`, `horzOverflow="clip"`, 1)
		}
		width, height := int64(2200000), int64(500000)
		if name == "table-y-overflow" {
			height = 20000
		}
		frame := nativeExactTableGraphicFrameXML(4, name, []int64{width}, []int64{height}, [][]string{{cell}}, ` rot="1800000" flipH="1"`)
		frame = strings.Replace(frame, `<a:off x="400000" y="200000"/>`, `<a:off x="3000000" y="2200000"/>`, 1)
		if name == "table-above-slide" {
			frame = strings.Replace(frame, `y="2200000"`, `y="-1000000"`, 1)
		}
		if name == "table-rotated-group" {
			frame = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="3" name="Physical source group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000" flipH="1"><a:off x="0" y="0"/><a:ext cx="12000000" cy="7000000"/><a:chOff x="0" y="0"/><a:chExt cx="9000000" cy="7000000"/></a:xfrm></p:grpSpPr>` + frame + `</p:grpSp>`
		}
		input := nativeExtractFixture(t, nativeExtractFixtureOptions{mutate: func(parts map[string]string) {
			raw := parts["relocated/slides/slide-a.xml"]
			start, end := strings.Index(raw, "<p:sp>"), strings.Index(raw, "</p:sp>")+len("</p:sp>")
			parts["relocated/slides/slide-a.xml"] = raw[:start] + frame + raw[end:]
		}})
		deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
		if err != nil {
			t.Fatal(err)
		}
		if issues := ValidateNativePPTX(deck); len(issues) != 0 {
			t.Fatal(issues)
		}
		var tables int
		var check func([]NativeElement)
		check = func(elements []NativeElement) {
			for _, element := range elements {
				if element.Kind == NativeElementKindTable && element.GraphicFrameLayout != nil && *element.GraphicFrameLayout == nativeSourceAnchoredGraphicFrame {
					tables++
				}
				check(element.Children)
			}
		}
		check(deck.Slides[0].Elements)
		expectedTables := 1
		if name == "table-unqualified-x-clip" {
			expectedTables = 0
		}
		if tables != expectedTables {
			t.Fatalf("%s: unexpected source-profiled table count %d: %#v", name, tables, deck.Slides[0].Compatibility)
		}
		encoded, err := json.MarshalIndent(deck, "", "  ")
		if err != nil {
			t.Fatal(err)
		}
		for ext, data := range map[string][]byte{".pptx": input, "-go.json": encoded} {
			if err := os.WriteFile(filepath.Join(dir, name+ext), data, 0644); err != nil {
				t.Fatal(err)
			}
		}
	}
}

func TestNativeGraphicFrameChartBrowserFixtures(t *testing.T) {
	dir := os.Getenv("INJOFFICE_PPTX_GRAPHIC_FRAME_FIXTURE_DIR")
	if dir == "" {
		t.Skip("external source/WASM/browser chart proof")
	}
	if err := os.MkdirAll(dir, 0755); err != nil {
		t.Fatal(err)
	}
	for _, workbook := range []bool{false, true} {
		for _, grouped := range []bool{false, true} {
			name := "chart-literal"
			if workbook {
				name = "chart-workbook"
			}
			if grouped {
				name += "-group"
			} else {
				name += "-baseline"
			}
			source := nativeLabeledBarXML(false)
			if workbook {
				source = strings.Replace(source, nativeBarSeriesXML(false), nativeWorkbookSeriesXML(false, false, false), 1)
				source = strings.ReplaceAll(source, "Sheet1!", "Data!")
				source = strings.Replace(source, `</c:chart>`, `<c:plotVisOnly val="0"/><c:dispBlanksAs val="gap"/></c:chart>`, 1)
				source = strings.Replace(source, `</c:chartSpace>`, `<c:externalData xmlns:r="`+nsOfficeRelsTransitional+`" r:id="workbook"><c:autoUpdate val="0"/></c:externalData></c:chartSpace>`, 1)
			}
			// Series color differs from supplied-label glyph color, so the browser can
			// prove both actual data paths and unscaled axis glyphs independently.
			start, end := strings.Index(source, "<c:ser"), strings.Index(source, "</c:ser>")+len("</c:ser>")
			source = source[:start] + strings.NewReplacer("123456", "E53935", "ABCDEF", "E53935").Replace(source[start:end]) + source[end:]
			var input []byte
			if workbook {
				input = nativeWorkbookInspectionFixture(t, false, "column", func(parts map[string]string) {
					parts["relocated/charts/source.xml"] = source
					parts["relocated/embeddings/Data.xlsx"] = string(nativeStackedWorkbookBytes(t, "mixed"))
				})
			} else {
				input = nativeChartFixture(t, nativeChartFixtureOptions{omitPreview: true, chartXML: source})
			}
			frame := strings.NewReplacer(`x="1000000" y="2000000"`, `x="0" y="0"`, `cx="3000000" cy="2000000"`, `cx="6000001" cy="4000001"`).Replace(nativeChartGraphicFrameXML(false, 3, name, ""))
			frame = strings.Replace(frame, `<p:xfrm>`, `<p:xfrm rot="1800000" flipV="1">`, 1)
			if grouped {
				frame = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="8" name="Nonuniform reflected rotated chart"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm rot="1800000" flipH="1"><a:off x="2000000" y="1000000"/><a:ext cx="7500000" cy="4000000"/><a:chOff x="0" y="0"/><a:chExt cx="6000000" cy="4000000"/></a:xfrm></p:grpSpPr>` + frame + `</p:grpSp>`
			}
			archive, err := zip.NewReader(bytes.NewReader(input), int64(len(input)))
			if err != nil {
				t.Fatal(err)
			}
			parts := []nativeExtractZipPart{}
			for _, entry := range archive.File {
				reader, err := entry.Open()
				if err != nil {
					t.Fatal(err)
				}
				raw, err := io.ReadAll(reader)
				reader.Close()
				if err != nil {
					t.Fatal(err)
				}
				data := string(raw)
				if entry.Name == "relocated/slides/slide-a.xml" {
					data = regexp.MustCompile(`(?s)<p:sp>.*?</p:sp>`).ReplaceAllString(data, "")
					data = regexp.MustCompile(`(?s)<p:graphicFrame.*?</p:graphicFrame>`).ReplaceAllStringFunc(data, func(string) string { return frame })
				}
				parts = append(parts, nativeExtractZipPart{name: entry.Name, data: data})
			}
			input = writeNativeExtractZip(t, parts)
			before := append([]byte(nil), input...)
			deck, err := ExtractNativePPTX(input, nativeAtomicTestExtractOptions())
			if err != nil {
				t.Fatal(err)
			}
			var charts []NativeElement
			var visit func([]NativeElement)
			visit = func(es []NativeElement) {
				for _, e := range es {
					if e.Kind == NativeElementKindChart {
						charts = append(charts, e)
					}
					visit(e.Children)
				}
			}
			visit(deck.Slides[0].Elements)
			if len(charts) != 1 {
				t.Fatalf("%s: expected one chart, got %d: %#v", name, len(charts), deck.Slides[0].Compatibility)
			}
			chart := charts[0]
			if chart.GraphicFrameLayout == nil || *chart.GraphicFrameLayout != nativeSourceAnchoredGraphicFrame || chart.Compatibility.Status != NativeCompatibilityStatusPreserveOnly || *chart.Transform.Cx != 6000001 || *chart.Transform.Cy != 4000001 || *chart.Transform.RotationAngle != 1800000 || !bytes.Equal(input, before) {
				t.Fatal(name, "source frame/profile mutated or absent")
			}
			if !workbook && chart.Chart.LiteralBar == nil {
				t.Fatal(name, "literal source chart unavailable")
			}
			if workbook {
				inspection, err := InspectNativePPTXChartWorkbooks(input)
				if err != nil || len(inspection.Charts) != 1 || len(inspection.Workbooks) != 1 || len(inspection.Omissions) != 0 {
					t.Fatalf("%s: workbook closure unavailable: %v %#v", name, err, inspection)
				}
			}
			encoded, err := json.MarshalIndent(deck, "", "  ")
			if err != nil {
				t.Fatal(err)
			}
			for ext, data := range map[string][]byte{".pptx": input, "-go.json": encoded} {
				if err := os.WriteFile(filepath.Join(dir, name+ext), data, 0644); err != nil {
					t.Fatal(err)
				}
			}
		}
	}
}
