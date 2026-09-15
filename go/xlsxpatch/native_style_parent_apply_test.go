package xlsxpatch

import (
	"slices"
	"strings"
	"testing"
)

func TestExtractNativeWorkbookRecordsStyleParentApplyMismatch(t *testing.T) {
	raw, _, err := BuildNativeGetCorpusFile("refuse-apply-flags")
	if err != nil {
		t.Fatal(err)
	}
	workbook, err := ExtractNativeWorkbookV2(raw)
	if err != nil {
		t.Fatalf("apply-flag mismatch still refuses the whole workbook: %v", err)
	}
	var diagnostic *NativeWorkbookUnsupportedV2
	for index := range workbook.Unsupported {
		if workbook.Unsupported[index].Code == "STYLE_PARENT_APPLY_MISMATCH" {
			if diagnostic != nil {
				t.Fatalf("mismatch diagnostic duplicated: %#v", workbook.Unsupported)
			}
			diagnostic = &workbook.Unsupported[index]
		}
	}
	if diagnostic == nil || diagnostic.ScopeID != "workbook" || diagnostic.Capability != "styles" || diagnostic.PartName == nil || *diagnostic.PartName != "xl/styles.xml" {
		t.Fatalf("mismatch diagnostic missing or misscoped: %#v", diagnostic)
	}
	for _, want := range []string{
		"cellXf 1 fill id 1 differs from cellStyleXf id 0 without applyFill",
		"cellXf 1 font id 1 differs from cellStyleXf id 0 without applyFont",
		"cellXf 2 border id 1 differs from cellStyleXf id 0 without applyBorder",
		"cellXf 3 number format id 164 differs from cellStyleXf id 0 without applyNumberFormat",
		"style-inheritance-dependent mutation is refused",
	} {
		if !strings.Contains(diagnostic.Message, want) {
			t.Fatalf("diagnostic %q lacks %q", diagnostic.Message, want)
		}
	}
	styles := workbook.Styles
	if len(styles) != 4 {
		t.Fatalf("styles = %d", len(styles))
	}
	if styles[1].Effective.FillColor == nil || *styles[1].Effective.FillColor != "#1F4E78" || styles[1].Effective.Bold == nil || !*styles[1].Effective.Bold || styles[1].Effective.FontColor == nil || *styles[1].Effective.FontColor != "#FFFFFF" {
		t.Fatalf("cellXf 1 did not display its own fill/font record: %#v", styles[1].Effective)
	}
	// borderId 1 has colorless sides, which the exact border projection refuses;
	// borderId 0 (the parent's) would have projected. The refusal proves the
	// cellXf's own record is the one displayed.
	if styles[0].Effective.Border == nil || styles[2].Effective.Border != nil || !slices.Contains(styles[2].Effective.Unsupported, "border") {
		t.Fatalf("cellXf 2 did not display its own border record: %#v / %#v", styles[0].Effective.Border, styles[2].Effective)
	}
	if styles[3].Effective.NumberFormat == nil || *styles[3].Effective.NumberFormat != "$#,##0" {
		t.Fatalf("cellXf 3 did not display its own number format: %#v", styles[3].Effective)
	}
	if issues := ValidateNativeWorkbookV2(workbook); len(issues) != 0 {
		t.Fatalf("contract invalid: %v", issues)
	}
	if _, err := newStyleRegistry([]byte(readEntry(t, raw, "xl/styles.xml"))); err == nil || !strings.Contains(err.Error(), "apply flag is false or absent") {
		t.Fatalf("strict mutation-boundary registry loosened: %v", err)
	}
	edit := mutation("edit", "1", CellSetValue, 1, 0)
	edit.Value = "Q9"
	if _, err := ApplyNativeWorkbookMutationTransactionV1(raw, NativeWorkbookMutationTransactionV1{ExpectedRevision: workbook.Revision, Cells: []CellMutation{edit}}); err == nil || !strings.Contains(err.Error(), "STYLE_PARENT_APPLY_MISMATCH") {
		t.Fatalf("mutation authority expanded over apply-flag mismatches: %v", err)
	}
}

func TestNewStyleRegistryForExtractionRecordsMismatchesButKeepsReferenceIDsFatal(t *testing.T) {
	t.Run("explicit false apply flag is recorded", func(t *testing.T) {
		xml := applyFlagStyleTable(`fontId="0" fillId="0" numFmtId="0"`, `fontId="0" fillId="2" numFmtId="0" applyFill="0"`, "", "")
		registry, mismatches, err := newStyleRegistryForExtraction([]byte(xml))
		if err != nil {
			t.Fatal(err)
		}
		if len(mismatches) != 1 || mismatches[0] != (styleParentApplyMismatch{cellXF: 0, component: "fill", flag: "applyFill", direct: 2, inherited: 0}) {
			t.Fatalf("mismatches = %#v", mismatches)
		}
		if projection := projectNativeWorkbookStyle(registry, 0); projection.FillColor == nil || *projection.FillColor != "#112233" {
			t.Fatalf("cellXf fill record was not displayed: %#v", projection)
		}
		if _, err := newStyleRegistry([]byte(xml)); err == nil {
			t.Fatal("strict registry accepted the mismatch")
		}
	})
	t.Run("alignment mismatch is recorded", func(t *testing.T) {
		xml := applyFlagStyleTable(`fontId="0" fillId="0" numFmtId="0"`, `fontId="0" fillId="0" numFmtId="0"`, `<x:alignment vertical="bottom"/>`, `<x:alignment horizontal="left"/>`)
		registry, mismatches, err := newStyleRegistryForExtraction([]byte(xml))
		if err != nil {
			t.Fatal(err)
		}
		if len(mismatches) != 1 || mismatches[0].component != "alignment" || mismatches[0].flag != "applyAlignment" {
			t.Fatalf("mismatches = %#v", mismatches)
		}
		if projection := projectNativeWorkbookStyle(registry, 0); projection.HorizontalAlignment == nil || *projection.HorizontalAlignment != "left" {
			t.Fatalf("cellXf alignment was not displayed: %#v", projection)
		}
		if !strings.Contains(styleParentApplyMismatchMessage(mismatches), "cellXf 0 alignment differs from its cellStyleXf parent without applyAlignment") {
			t.Fatalf("message = %q", styleParentApplyMismatchMessage(mismatches))
		}
	})
	t.Run("consistent tables record nothing", func(t *testing.T) {
		xml := applyFlagStyleTable(`fontId="0" fillId="0" numFmtId="0"`, `fontId="0" fillId="2" numFmtId="0" applyFill="1"`, "", "")
		if _, mismatches, err := newStyleRegistryForExtraction([]byte(xml)); err != nil || len(mismatches) != 0 {
			t.Fatalf("consistent table: mismatches=%#v err=%v", mismatches, err)
		}
		if styleParentApplyMismatchMessage(nil) != "" {
			t.Fatal("empty mismatch list produced a message")
		}
	})
	for _, test := range []struct{ name, cellAttrs, want string }{
		{"fill outside table", `fontId="0" fillId="9" numFmtId="0"`, "outside the fills table"},
		{"font outside table", `fontId="9" fillId="0" numFmtId="0"`, "outside the fonts table"},
		{"missing number format", `fontId="0" fillId="0" numFmtId="170"`, "numFmtId 170 is missing"},
	} {
		t.Run(test.name+" stays fatal", func(t *testing.T) {
			xml := applyFlagStyleTable(`fontId="0" fillId="0" numFmtId="0"`, test.cellAttrs, "", "")
			if _, _, err := newStyleRegistryForExtraction([]byte(xml)); err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("reference id error = %v, want %q", err, test.want)
			}
		})
	}
	t.Run("message is bounded", func(t *testing.T) {
		mismatches := make([]styleParentApplyMismatch, 0, 12)
		for index := 0; index < 12; index++ {
			mismatches = append(mismatches, styleParentApplyMismatch{cellXF: 11 - index, component: "fill", flag: "applyFill", direct: 1})
		}
		message := styleParentApplyMismatchMessage(mismatches)
		if strings.Count(message, "cellXf ") != styleParentApplyMismatchMessageLimit || !strings.Contains(message, "and 4 more") || !strings.HasPrefix(message, "cellXf 0 fill") {
			t.Fatalf("message = %q", message)
		}
	})
}
