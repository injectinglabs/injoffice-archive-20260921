package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"io"
	"os"
	"slices"
	"strconv"
	"strings"
	"testing"
)

func nativeMutationFixture(strict bool) map[string]string {
	entries := nativeWorkbookFixture(strict)
	ssNS := spreadsheetMLTransitional
	if strict {
		ssNS = spreadsheetMLStrict
	}
	entries["Meta/Styles.style"] = strings.Replace(entries["Meta/Styles.style"], `<color theme="1"/>`, `<color rgb="FF000000"/>`, 1)
	entries["Sheets/s1.xml"] = `<worksheet xmlns="` + ssNS + `"><dimension ref="A1:E1"/><cols><col min="1" max="3" width="12.25" customWidth="1"/></cols><sheetData>` +
		`<row r="1" ht="20" customHeight="1"><c r="A1"><v>001.2300</v></c><c r="B1" t="inlineStr"><is><t>old</t></is></c><c r="C1" s="1"><v>5</v></c><c r="D1"><f>OLD()</f><v>9</v></c>` +
		`<c r="E1" t="inlineStr"><is><t>authority-bound</t></is><extLst><ext uri="keep"><opaque xmlns="urn:test"/></ext></extLst></c></row>` +
		`</sheetData></worksheet>`
	return entries
}

func nativeRawBordersForTest(t *testing.T, packageBytes []byte) string {
	t.Helper()
	styles := readEntry(t, packageBytes, "Meta/Styles.style")
	start, end := strings.Index(styles, "<borders"), strings.Index(styles, "</borders>")
	if start < 0 || end < start {
		t.Fatalf("styles part has no borders table: %s", styles)
	}
	return styles[start : end+len("</borders>")]
}

func TestApplyNativeWorkbookMutationTransactionV1AtomicRoundTrip(t *testing.T) {
	for _, strict := range []bool{false, true} {
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			entries := nativeMutationFixture(strict)
			original := buildZip(t, entries)
			before, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			setText := mutation("set-text", "7", CellSetValue, 0, 1)
			setText.Value = "  exact _x0041_ text  "
			setFormula := mutation("set-formula", "7", CellSetFormula, 0, 3)
			setFormula.Formula = "=SUM(A1,C1)"
			style := StylePatchMutation{
				OperationID: "set-style", SheetID: "7", Kind: StylePatch,
				Range: StyleRange{Row: 0, Column: 2, EndRow: 0, EndColumn: 2},
				Style: StyleDelta{
					NumberFormat: ClearStyleProperty[string](), FontColor: ClearStyleProperty[string](), Bold: ClearStyleProperty[bool](),
					FillColor: SetStyleProperty("#DDEEFF"),
				},
			}
			row := LayoutMutation{OperationID: "set-row", SheetID: "7", Kind: RowSetHeight, Row: 0, HeightPoints: 24.5}
			column := LayoutMutation{OperationID: "set-column", SheetID: "7", Kind: ColumnSetWidth, Column: 1, Width: 14.25}
			result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
				ExpectedRevision: before.Revision,
				Cells:            []CellMutation{setText, setFormula}, Styles: []StylePatchMutation{style}, Layout: []LayoutMutation{row, column},
			})
			if err != nil {
				t.Fatal(err)
			}
			if result == nil || len(result.Package) == 0 || result.Workbook == nil || result.Workbook.DocumentID != before.DocumentID || result.Workbook.Revision == before.Revision {
				t.Fatalf("invalid mutation result: %#v", result)
			}
			lexical := findNativeCell(t, result.Workbook, "7", "A1")
			if lexical.Value == nil || lexical.Value.Lexical == nil || *lexical.Value.Lexical != "001.2300" {
				t.Fatalf("untouched numeric lexical spelling changed: %#v", lexical)
			}
			if sheetXML := readEntry(t, result.Package, "Sheets/s1.xml"); !strings.Contains(sheetXML, `<c r="A1"><v>001.2300</v></c>`) {
				t.Fatalf("untouched numeric OOXML lexical bytes changed: %s", sheetXML)
			}
			text := findNativeCell(t, result.Workbook, "7", "B1")
			if text.Value == nil || text.Value.Text == nil || *text.Value.Text != "  exact _x0041_ text  " {
				t.Fatalf("SpreadsheetML string did not round trip: %#v", text)
			}
			formula := findNativeCell(t, result.Workbook, "7", "D1")
			if formula.Formula == nil || formula.Formula.Text != "SUM(A1,C1)" || formula.Formula.Cached != nil {
				t.Fatalf("formula post-save extraction mismatch: %#v", formula)
			}
			styled := findNativeCell(t, result.Workbook, "7", "C1")
			effective := result.Workbook.Styles[styled.StyleID].Effective
			if effective.NumberFormat == nil || *effective.NumberFormat != "General" || effective.FontColor == nil || *effective.FontColor != "#000000" || effective.Bold == nil || *effective.Bold || effective.FillColor == nil || *effective.FillColor != "#DDEEFF" {
				t.Fatalf("clear-to-inherited style did not round trip: %#v", effective)
			}
			beforeBorder, _ := json.Marshal(before.Styles[1].Effective.Border)
			afterBorder, _ := json.Marshal(effective.Border)
			if !bytes.Equal(beforeBorder, afterBorder) || nativeRawBordersForTest(t, original) != nativeRawBordersForTest(t, result.Package) {
				t.Fatalf("immutable border authority changed across fill mutation: before=%s after=%s", beforeBorder, afterBorder)
			}
			for _, part := range []string{"Charts/chart1.xml", "Custom/data.bin", "Meta/Strings.xml", "Sheets/s2.xml"} {
				if got := readEntry(t, result.Package, part); got != entries[part] {
					t.Fatalf("untouched OPC part %q changed", part)
				}
			}
			if err := compareUntouchedRawZipEntries(original, result.Package, map[string]bool{
				"Book/Workbook.xml": true, "Meta/Styles.style": true, "Sheets/s1.xml": true,
			}); err != nil {
				t.Fatalf("no-calc-chain transaction changed a non-target raw OPC entry: %v", err)
			}
			reopened, err := ExtractNativeWorkbookV1WithOptions(result.Package, NativeWorkbookExtractionOptions{Previous: result.Workbook})
			if err != nil || reopened.Revision != result.Workbook.Revision || len(ValidateNativeWorkbookV1(reopened)) != 0 {
				t.Fatalf("saved package did not reopen deterministically: revision=%v err=%v", reopened, err)
			}
			assertIndependentXLSXMutationPackageConsumer(t, original, result.Package)
		})
	}
}

func TestApplyNativeWorkbookMutationTransactionV1RejectsStaleAndUnsupportedAtomically(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	valid := mutation("valid-first", "7", CellSetValue, 0, 1)
	valid.Value = "must-not-commit"
	unsupported := mutation("unsupported-second", "7", CellSetValue, 0, 4)
	unsupported.Value = "would-drop-extension"

	result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: before.Revision,
		Cells:            []CellMutation{valid, unsupported},
	})
	if err == nil || result != nil || !strings.Contains(err.Error(), "mutation-refused cell") {
		t.Fatalf("unsupported target was not refused atomically: result=%#v err=%v", result, err)
	}
	stale := "rev:" + strings.Repeat("0", 64)
	result, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
		ExpectedRevision: stale,
		Cells:            []CellMutation{valid},
	})
	if err == nil || result != nil || !strings.Contains(err.Error(), "stale revision") {
		t.Fatalf("stale source was accepted: result=%#v err=%v", result, err)
	}
}

func TestApplyNativeWorkbookMutationTransactionV1AdversarialContractRefusals(t *testing.T) {
	t.Run("formula group refuses whole sheet", func(t *testing.T) {
		original := buildZip(t, nativeWorkbookFixture(false))
		before, err := ExtractNativeWorkbookV1(original)
		if err != nil {
			t.Fatal(err)
		}
		operation := mutation("outside-group", "7", CellSetValue, 0, 2)
		operation.Value = 7
		result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
			ExpectedRevision: before.Revision, Cells: []CellMutation{operation},
		})
		if err == nil || result != nil || !strings.Contains(err.Error(), "FORMULA_GROUPS") {
			t.Fatalf("formula-group sheet was accepted: result=%#v err=%v", result, err)
		}
	})

	t.Run("partial source style refuses target", func(t *testing.T) {
		original := buildZip(t, nativeWorkbookFixture(false))
		before, err := ExtractNativeWorkbookV1(original)
		if err != nil {
			t.Fatal(err)
		}
		operation := StylePatchMutation{
			OperationID: "partial-style", SheetID: "9", Kind: StylePatch,
			Range: StyleRange{Row: 1, Column: 0, EndRow: 1, EndColumn: 0},
			Style: StyleDelta{Bold: SetStyleProperty(true)},
		}
		result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
			ExpectedRevision: before.Revision, Styles: []StylePatchMutation{operation},
		})
		if err == nil || result != nil || !strings.Contains(err.Error(), "unsupported style 0") {
			t.Fatalf("partial source style was accepted: result=%#v err=%v", result, err)
		}
	})
}

func TestVerifyNativeUnsupportedInventoryRequiresExactIdentitySet(t *testing.T) {
	part := "Sheets/s1.xml"
	before := []NativeWorkbookUnsupportedV1{{
		ID: "unsupported:a", Code: "CELL_EXTENSIONS", Capability: "extensions", ScopeID: "sheet:7",
		PartName: &part, Preservation: "preserve-exact", Message: "source message",
	}, {
		ID: "unsupported:b", Code: "OPAQUE_CELL_MARKUP", Capability: "cell-markup", ScopeID: "sheet:7",
		PartName: &part, Preservation: "preserve-exact", Message: "second source message",
	}}
	introduced := append([]NativeWorkbookUnsupportedV1{}, before...)
	introduced = append(introduced, NativeWorkbookUnsupportedV1{
		ID: "unsupported:b", Code: "OPAQUE_CELL_MARKUP", Capability: "cell-markup", ScopeID: "sheet:7",
		PartName: &part, Preservation: "preserve-exact", Message: "new item",
	})
	if err := verifyNativeUnsupportedInventoryPreserved(before, introduced); err == nil {
		t.Fatal("new unsupported item was accepted")
	}
	duplicate := append([]NativeWorkbookUnsupportedV1{}, before...)
	duplicate[1] = duplicate[0]
	if err := verifyNativeUnsupportedInventoryPreserved(before, duplicate); err == nil || !strings.Contains(err.Error(), "duplicates id") {
		t.Fatalf("equal-length duplicate result inventory was accepted: %v", err)
	}
	remapped := append([]NativeWorkbookUnsupportedV1{}, before...)
	remapped[0].ScopeID = "sheet:9"
	if err := verifyNativeUnsupportedInventoryPreserved(before, remapped); err == nil {
		t.Fatal("remapped unsupported item was accepted")
	}
	messageChanged := append([]NativeWorkbookUnsupportedV1{}, before...)
	messageChanged[0].Message = "rewritten explanation"
	if err := verifyNativeUnsupportedInventoryPreserved(before, messageChanged); err == nil {
		t.Fatal("non-exact unsupported message was accepted")
	}
}

func TestDecodeNativeWorkbookMutationTransactionV1StrictBoundedAndCASJoined(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	set := mutation("decoded-set", "7", CellSetValue, 0, 1)
	set.Value = "decoded"
	valid, err := json.Marshal(NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Cells: []CellMutation{set}})
	if err != nil {
		t.Fatal(err)
	}
	decoded, err := DecodeNativeWorkbookMutationTransactionV1(valid, before.Source.PackageSHA256)
	if err != nil || decoded.ExpectedRevision != before.Revision || len(decoded.Cells) != 1 {
		t.Fatalf("valid joined payload was refused: decoded=%#v err=%v", decoded, err)
	}
	result, err := ApplyNativeWorkbookMutationPayloadV1(original, valid, before.Source.PackageSHA256)
	if err != nil || result == nil || findNativeCell(t, result.Workbook, "7", "B1").Value == nil {
		t.Fatalf("strict raw payload did not apply: result=%#v err=%v", result, err)
	}
	exactNumberPayload := []byte(strings.Replace(string(valid), `"value":"decoded"`, `"value":9007199254740993`, 1))
	exactNumber, err := DecodeNativeWorkbookMutationTransactionV1(exactNumberPayload, before.Source.PackageSHA256)
	if err != nil {
		t.Fatal(err)
	}
	if number, ok := exactNumber.Cells[0].Value.(json.Number); !ok || number.String() != "9007199254740993" {
		t.Fatalf("raw JSON number was coerced before native mutation: %#v", exactNumber.Cells[0].Value)
	}

	zeroRev := "rev:" + strings.Repeat("0", 64)
	validText := string(valid)
	cases := map[string][]byte{
		"unknown root":        []byte(strings.Replace(validText, `{"expected_revision":`, `{"unknown":1,"expected_revision":`, 1)),
		"duplicate root":      []byte(strings.Replace(validText, `{"expected_revision":`, `{"expected_revision":"`+before.Revision+`","expected_revision":`, 1)),
		"unknown nested":      []byte(strings.Replace(validText, `"operation_id":"decoded-set"`, `"mystery":true,"operation_id":"decoded-set"`, 1)),
		"duplicate nested":    []byte(strings.Replace(validText, `"operation_id":"decoded-set"`, `"operation_id":"other","operation_id":"decoded-set"`, 1)),
		"trailing object":     append(append([]byte{}, valid...), []byte(` {}`)...),
		"native CAS mismatch": []byte(strings.Replace(validText, before.Revision, zeroRev, 1)),
	}
	for name, payload := range cases {
		t.Run(name, func(t *testing.T) {
			if _, err := DecodeNativeWorkbookMutationTransactionV1(payload, before.Source.PackageSHA256); err == nil {
				t.Fatal("hostile raw payload was accepted")
			}
		})
	}
	if _, err := DecodeNativeWorkbookMutationTransactionV1(valid, before.Revision); err == nil {
		t.Fatal("native rev was ambiguously accepted as the outer CAS")
	}
	oversized := bytes.Repeat([]byte{' '}, MaxNativeWorkbookMutationPayloadBytes+1)
	if _, err := DecodeNativeWorkbookMutationTransactionV1(oversized, before.Source.PackageSHA256); err == nil {
		t.Fatal("oversized raw payload was accepted")
	}
	deep := []byte(`{"unknown":` + strings.Repeat("[", nativeWorkbookMaxJSONDepth+1) + `0` + strings.Repeat("]", nativeWorkbookMaxJSONDepth+1) + `}`)
	if _, err := DecodeNativeWorkbookMutationTransactionV1(deep, before.Source.PackageSHA256); err == nil || !strings.Contains(err.Error(), "nesting exceeds") {
		t.Fatalf("depth-bomb payload did not hit the bounded duplicate-key scanner: %v", err)
	}
	staleOuter := "sha256:" + strings.Repeat("0", 64)
	if result, err := ApplyNativeWorkbookMutationPayloadV1(original, valid, staleOuter); err == nil || result != nil || !strings.Contains(err.Error(), "stale outer revision") {
		t.Fatalf("stale outer source CAS was accepted: result=%#v err=%v", result, err)
	}
}

func TestNativeWorkbookMutationVerificationRefusesTopologyAndAfterOnlyCells(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	t.Run("sheet topology", func(t *testing.T) {
		after := *before
		after.Sheets = append([]NativeWorkbookSheetV1(nil), before.Sheets...)
		after.Sheets[0].Name = "Renamed behind transaction"
		if err := verifyNativeWorkbookSheetTopologyPreserved(before, &after); err == nil {
			t.Fatal("sheet rename escaped topology verification")
		}
		after.Sheets = after.Sheets[:1]
		if err := verifyNativeWorkbookSheetTopologyPreserved(before, &after); err == nil {
			t.Fatal("sheet removal escaped topology verification")
		}
	})
	t.Run("after-only cell", func(t *testing.T) {
		after := *before
		after.Sheets = append([]NativeWorkbookSheetV1(nil), before.Sheets...)
		after.Sheets[0].Cells = append([]NativeWorkbookCellV1(nil), before.Sheets[0].Cells...)
		after.Sheets[0].Cells = append(after.Sheets[0].Cells, NativeWorkbookCellV1{Row: 5, Column: 5, Ref: "F6", Editable: true})
		if err := verifyNativeWorkbookMutationResult(before, &after, NativeWorkbookMutationTransactionV1{}, nativeExpectedStyleProjections{}); err == nil || !strings.Contains(err.Error(), "after-only cell") {
			t.Fatalf("after-only cell escaped bidirectional verification: %v", err)
		}
	})
}

func TestNativeWorkbookMutationVerificationRefusesCollateralProjectionDrift(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	verify := func(t *testing.T, mutate func(*NativeWorkbookV1), want string) {
		t.Helper()
		after := *before
		after.Styles = slices.Clone(before.Styles)
		after.Sheets = slices.Clone(before.Sheets)
		for index := range after.Sheets {
			after.Sheets[index].Rows = slices.Clone(before.Sheets[index].Rows)
			after.Sheets[index].Columns = slices.Clone(before.Sheets[index].Columns)
			after.Sheets[index].Cells = slices.Clone(before.Sheets[index].Cells)
			after.Sheets[index].MergedRanges = slices.Clone(before.Sheets[index].MergedRanges)
		}
		mutate(&after)
		err := verifyNativeWorkbookMutationResult(before, &after, NativeWorkbookMutationTransactionV1{}, nativeExpectedStyleProjections{})
		if err == nil || !strings.Contains(err.Error(), want) {
			t.Fatalf("collateral projection drift escaped verification: %v", err)
		}
	}

	t.Run("untouched cell style", func(t *testing.T) {
		verify(t, func(after *NativeWorkbookV1) { after.Sheets[0].Cells[0].StyleID = 1 }, "style changed outside style mutation")
	})
	t.Run("untouched cell authority", func(t *testing.T) {
		verify(t, func(after *NativeWorkbookV1) { after.Sheets[0].Cells[0].Editable = false }, "cell authority changed outside mutation")
	})
	t.Run("merged range", func(t *testing.T) {
		verify(t, func(after *NativeWorkbookV1) {
			after.Sheets[0].MergedRanges = append(after.Sheets[0].MergedRanges, NativeWorkbookMergedRangeV1{
				Ref: "F1:G1", Row: 0, Column: 5, EndRow: 0, EndColumn: 6, Editable: false,
			})
		}, "merged-range projection changed")
	})
	t.Run("cleared cell style", func(t *testing.T) {
		after := *before
		after.Sheets = append([]NativeWorkbookSheetV1(nil), before.Sheets...)
		after.Sheets[0].Cells = append([]NativeWorkbookCellV1(nil), before.Sheets[0].Cells...)
		after.Sheets[0].Cells = append(after.Sheets[0].Cells[:2], after.Sheets[0].Cells[3:]...)
		clear := mutation("clear-styled", "7", CellClearValue, 0, 2)
		err := verifyNativeWorkbookMutationResult(before, &after, NativeWorkbookMutationTransactionV1{Cells: []CellMutation{clear}}, nativeExpectedStyleProjections{})
		if err == nil || !strings.Contains(err.Error(), "style changed outside style mutation") {
			t.Fatalf("cleared cell lost its existing style without verification: %v", err)
		}
	})
	t.Run("existing style record", func(t *testing.T) {
		verify(t, func(after *NativeWorkbookV1) {
			changed := true
			after.Styles[0].Effective.Bold = &changed
		}, "existing style 0 changed")
	})
	t.Run("row dimension", func(t *testing.T) {
		verify(t, func(after *NativeWorkbookV1) {
			changed := 21.0
			after.Sheets[0].Rows[0].HeightPoints = &changed
		}, "row dimensions changed without a row mutation")
	})
	t.Run("column dimension", func(t *testing.T) {
		verify(t, func(after *NativeWorkbookV1) {
			changed := 13.0
			after.Sheets[0].Columns[0].Width = &changed
		}, "column dimensions changed without a column mutation")
	})
}

func TestApplyNativeWorkbookMutationTransactionV1RefusesBlankMergedTargets(t *testing.T) {
	for _, strict := range []bool{false, true} {
		strict := strict
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			entries := nativeMutationFixture(strict)
			entries["Sheets/s1.xml"] = strings.Replace(entries["Sheets/s1.xml"], `</sheetData></worksheet>`, `</sheetData><mergeCells count="1"><mergeCell ref="F1:G1"/></mergeCells></worksheet>`, 1)
			original := buildZip(t, entries)
			before, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			cell := mutation("blank-merge-cell", "7", CellSetValue, 0, 5)
			cell.Value = "refused"
			style := StylePatchMutation{
				OperationID: "blank-merge-style", SheetID: "7", Kind: StylePatch,
				Range: StyleRange{Row: 0, Column: 6, EndRow: 0, EndColumn: 6},
				Style: StyleDelta{Bold: SetStyleProperty(true)},
			}
			for _, test := range []struct {
				name        string
				transaction NativeWorkbookMutationTransactionV1
			}{
				{name: "cell", transaction: NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Cells: []CellMutation{cell}}},
				{name: "style", transaction: NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Styles: []StylePatchMutation{style}}},
			} {
				t.Run(test.name, func(t *testing.T) {
					result, err := ApplyNativeWorkbookMutationTransactionV1(original, test.transaction)
					if err == nil || result != nil || !strings.Contains(err.Error(), "mutation-refused merged cell") {
						t.Fatalf("blank merged target was not refused atomically: result=%#v err=%v", result, err)
					}
				})
			}
		})
	}
}

func TestApplyNativeWorkbookMutationTransactionV1ClearStylePreservesBorderAuthority(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	clear := StylePatchMutation{
		OperationID: "clear-direct-style", SheetID: "7", Kind: StylePatch,
		Range: StyleRange{Row: 0, Column: 2, EndRow: 0, EndColumn: 2},
		Style: StyleDelta{
			NumberFormat: ClearStyleProperty[string](), FontName: ClearStyleProperty[string](), FontSizePoints: ClearStyleProperty[float64](),
			Bold: ClearStyleProperty[bool](), Italic: ClearStyleProperty[bool](), FontColor: ClearStyleProperty[string](),
			FillColor: ClearStyleProperty[string](), HorizontalAlignment: ClearStyleProperty[HorizontalAlignment](),
			VerticalAlignment: ClearStyleProperty[VerticalAlignment](), WrapText: ClearStyleProperty[bool](),
		},
	}
	result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Styles: []StylePatchMutation{clear}})
	if err != nil {
		t.Fatal(err)
	}
	cell := findNativeCell(t, result.Workbook, "7", "C1")
	if cell.StyleID == 0 {
		t.Fatalf("clearable style fields incorrectly discarded immutable border authority: %#v", cell)
	}
	beforeBorder, _ := json.Marshal(before.Styles[1].Effective.Border)
	afterBorder, _ := json.Marshal(result.Workbook.Styles[cell.StyleID].Effective.Border)
	if !bytes.Equal(beforeBorder, afterBorder) || nativeRawBordersForTest(t, original) != nativeRawBordersForTest(t, result.Package) {
		t.Fatalf("border record changed across full mutable-style clear: before=%s after=%s", beforeBorder, afterBorder)
	}
}

func TestApplyNativeWorkbookMutationTransactionV1RefusesSemanticNoop(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	alreadySatisfied := LayoutMutation{OperationID: "same-row", SheetID: "7", Kind: RowSetHeight, Row: 0, HeightPoints: 20}
	result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Layout: []LayoutMutation{alreadySatisfied}})
	if err == nil || result != nil || !strings.Contains(err.Error(), "semantic no-op") {
		t.Fatalf("semantic no-op transaction was accepted: result=%#v err=%v", result, err)
	}
}

func TestApplyNativeWorkbookMutationTransactionV1ResourceAndCrossBatchBounds(t *testing.T) {
	original := buildZip(t, nativeMutationFixture(false))
	before, err := ExtractNativeWorkbookV1(original)
	if err != nil {
		t.Fatal(err)
	}
	operations := make([]LayoutMutation, maxNativeWorkbookTransactionOperations+1)
	for index := range operations {
		operations[index] = LayoutMutation{OperationID: "op-" + strconv.Itoa(index), SheetID: "7", Kind: RowSetHeight, Row: index, HeightPoints: 12}
	}
	result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Layout: operations})
	if err == nil || result != nil || !strings.Contains(err.Error(), "transaction exceeds") {
		t.Fatalf("oversized transaction was accepted: result=%#v err=%v", result, err)
	}

	cell := mutation("same-id", "7", CellSetValue, 0, 1)
	cell.Value = "x"
	layout := LayoutMutation{OperationID: "same-id", SheetID: "7", Kind: RowSetHeight, Row: 0, HeightPoints: 20}
	result, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Cells: []CellMutation{cell}, Layout: []LayoutMutation{layout}})
	if err == nil || result != nil || !strings.Contains(err.Error(), "duplicate operation_id") {
		t.Fatalf("cross-family duplicate id was accepted: result=%#v err=%v", result, err)
	}

	styleTargets := []StylePatchMutation{
		{
			OperationID: "range-a", SheetID: "7", Kind: StylePatch,
			Range: StyleRange{Row: 0, Column: 0, EndRow: 5_999, EndColumn: 0},
			Style: StyleDelta{Bold: SetStyleProperty(true)},
		},
		{
			OperationID: "range-b", SheetID: "7", Kind: StylePatch,
			Range: StyleRange{Row: 6_000, Column: 0, EndRow: 11_999, EndColumn: 0},
			Style: StyleDelta{Italic: SetStyleProperty(true)},
		},
	}
	result, err = ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{ExpectedRevision: before.Revision, Styles: styleTargets})
	if err == nil || result != nil || !strings.Contains(err.Error(), "style target expansion exceeds 10000") {
		t.Fatalf("aggregate style expansion was accepted: result=%#v err=%v", result, err)
	}
}

func TestApplyNativeWorkbookMutationTransactionV1IndependentPackageConsumer(t *testing.T) {
	for _, strict := range []bool{false, true} {
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			entries := nativeMutationFixture(strict)
			original := buildZip(t, entries)
			before, err := ExtractNativeWorkbookV1(original)
			if err != nil {
				t.Fatal(err)
			}
			setText := mutation("set-text", "7", CellSetValue, 0, 1)
			setText.Value = "  exact _x0041_ text  "
			setFormula := mutation("set-formula", "7", CellSetFormula, 0, 3)
			setFormula.Formula = "=SUM(A1,C1)"
			row := LayoutMutation{OperationID: "set-row", SheetID: "7", Kind: RowSetHeight, Row: 0, HeightPoints: 24.5}
			column := LayoutMutation{OperationID: "set-column", SheetID: "7", Kind: ColumnSetWidth, Column: 1, Width: 14.25}
			result, err := ApplyNativeWorkbookMutationTransactionV1(original, NativeWorkbookMutationTransactionV1{
				ExpectedRevision: before.Revision,
				Cells:            []CellMutation{setText, setFormula},
				Layout:           []LayoutMutation{row, column},
			})
			if err != nil {
				t.Fatal(err)
			}
			assertIndependentXLSXMutationPackageConsumer(t, original, result.Package)
		})
	}
}

func TestNativeWorkbookMutationPathHasNoLegacyReconstructionDependencies(t *testing.T) {
	source, err := os.ReadFile("native_mutate.go")
	if err != nil {
		t.Fatal(err)
	}
	lower := strings.ToLower(string(source))
	for _, forbidden := range []string{"luckyexcel", "mammoth", "document.createelement", "domparser", "jszip", "json.marshal"} {
		if strings.Contains(lower, forbidden) {
			t.Errorf("native mutation authority contains forbidden dependency %q", forbidden)
		}
	}
}

func assertIndependentXLSXMutationPackageConsumer(t *testing.T, original, produced []byte) {
	t.Helper()
	before := independentOPCEntries(t, original)
	after := independentOPCEntries(t, produced)
	assertIndependentContentTypes(t, after["[Content_Types].xml"])
	for _, part := range []string{"Charts/chart1.xml", "Custom/data.bin", "Meta/Strings.xml", "Sheets/s2.xml"} {
		if !bytes.Equal(before[part], after[part]) {
			t.Fatalf("independent consumer: untouched OPC part %q changed", part)
		}
	}
	sheet, ok := after["Sheets/s1.xml"]
	if !ok {
		t.Fatal("independent consumer: mutated worksheet part is missing")
	}
	cells, rowHeight, columnWidth := independentSpreadsheetCells(t, sheet)
	if cells["A1"].value != "001.2300" || cells["A1"].inlineText != "" || cells["A1"].formula != "" {
		t.Fatalf("independent consumer: untouched A1 drifted: %#v", cells["A1"])
	}
	if cells["B1"].inlineText != "  exact _x005F_x0041_ text  " {
		t.Fatalf("independent consumer: mutated B1 inline text = %q", cells["B1"].inlineText)
	}
	if cells["D1"].formula != "SUM(A1,C1)" || cells["D1"].value != "" {
		t.Fatalf("independent consumer: mutated D1 formula = %#v", cells["D1"])
	}
	if rowHeight != 24.5 {
		t.Fatalf("independent consumer: row 1 height = %v", rowHeight)
	}
	if columnWidth != 14.25 {
		t.Fatalf("independent consumer: column B width = %v", columnWidth)
	}
}

func independentOPCEntries(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	reader, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatalf("independent consumer could not reopen OPC zip: %v", err)
	}
	entries := make(map[string][]byte, len(reader.File))
	for _, file := range reader.File {
		opened, openErr := file.Open()
		if openErr != nil {
			t.Fatalf("independent consumer could not open %q: %v", file.Name, openErr)
		}
		payload, readErr := io.ReadAll(opened)
		opened.Close()
		if readErr != nil {
			t.Fatalf("independent consumer could not read %q: %v", file.Name, readErr)
		}
		if _, exists := entries[file.Name]; exists {
			t.Fatalf("independent consumer found duplicate zip name %q", file.Name)
		}
		entries[file.Name] = payload
	}
	if _, ok := entries["[Content_Types].xml"]; !ok {
		t.Fatal("independent consumer missing [Content_Types].xml")
	}
	return entries
}

func assertIndependentContentTypes(t *testing.T, data []byte) {
	t.Helper()
	decoder := xml.NewDecoder(bytes.NewReader(data))
	root := false
	children := 0
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("independent consumer: [Content_Types].xml is not well-formed XML: %v", err)
		}
		start, ok := token.(xml.StartElement)
		if !ok {
			continue
		}
		if start.Name.Space != nativeContentTypesNamespace {
			t.Fatalf("independent consumer: unexpected content-types namespace %q", start.Name.Space)
		}
		if !root {
			if start.Name.Local != "Types" {
				t.Fatalf("independent consumer: content-types root is %q", start.Name.Local)
			}
			root = true
			continue
		}
		if start.Name.Local != "Default" && start.Name.Local != "Override" {
			t.Fatalf("independent consumer: unexpected content-types child %q", start.Name.Local)
		}
		children++
	}
	if !root || children == 0 {
		t.Fatalf("independent consumer: [Content_Types].xml had root=%v children=%d", root, children)
	}
}

type independentSpreadsheetCell struct {
	value      string
	inlineText string
	formula    string
}

func independentSpreadsheetCells(t *testing.T, data []byte) (map[string]independentSpreadsheetCell, float64, float64) {
	t.Helper()
	decoder := xml.NewDecoder(bytes.NewReader(data))
	cells := map[string]independentSpreadsheetCell{}
	rowHeight := 0.0
	columnWidth := 0.0
	var path []string
	var currentRef string
	var current independentSpreadsheetCell
	var text strings.Builder
	collect := ""
	for {
		token, err := decoder.Token()
		if err == io.EOF {
			break
		}
		if err != nil {
			t.Fatalf("independent consumer: worksheet XML is not well-formed: %v", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			path = append(path, token.Name.Local)
			if token.Name.Space != spreadsheetMLTransitional && token.Name.Space != spreadsheetMLStrict {
				continue
			}
			switch token.Name.Local {
			case "col":
				min, minErr := strconv.Atoi(unqualifiedXMLAttr(token, "min"))
				max, maxErr := strconv.Atoi(unqualifiedXMLAttr(token, "max"))
				width, widthErr := strconv.ParseFloat(unqualifiedXMLAttr(token, "width"), 64)
				if minErr == nil && maxErr == nil && widthErr == nil && min <= 2 && max >= 2 && unqualifiedXMLAttr(token, "customWidth") == "1" {
					columnWidth = width
				}
			case "row":
				if unqualifiedXMLAttr(token, "r") == "1" {
					height, heightErr := strconv.ParseFloat(unqualifiedXMLAttr(token, "ht"), 64)
					if heightErr != nil || unqualifiedXMLAttr(token, "customHeight") != "1" {
						t.Fatalf("independent consumer: row 1 sizing attrs=%v", token.Attr)
					}
					rowHeight = height
				}
			case "c":
				currentRef = unqualifiedXMLAttr(token, "r")
				current = independentSpreadsheetCell{}
				text.Reset()
				collect = ""
			case "v", "f", "t":
				text.Reset()
				collect = token.Name.Local
			}
		case xml.CharData:
			if collect != "" {
				text.Write(token)
			}
		case xml.EndElement:
			if collect == token.Name.Local {
				switch collect {
				case "v":
					current.value = text.String()
				case "f":
					current.formula = text.String()
				case "t":
					if len(path) >= 2 && path[len(path)-2] == "is" {
						current.inlineText = text.String()
					}
				}
				collect = ""
			}
			if token.Name.Local == "c" && currentRef != "" {
				cells[currentRef] = current
				currentRef = ""
			}
			if len(path) == 0 || path[len(path)-1] != token.Name.Local {
				t.Fatalf("independent consumer: worksheet element stack drifted at %q", token.Name.Local)
			}
			path = path[:len(path)-1]
		}
	}
	return cells, rowHeight, columnWidth
}

func unqualifiedXMLAttr(start xml.StartElement, local string) string {
	for _, attr := range start.Attr {
		if attr.Name.Space == "" && attr.Name.Local == local {
			return attr.Value
		}
	}
	return ""
}
