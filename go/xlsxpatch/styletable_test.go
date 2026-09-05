package xlsxpatch

import (
	"strings"
	"testing"
)

func TestParseStyleTable_IndexesRequiredRecordsWithoutRewriting(t *testing.T) {
	styles := []byte(`<?xml version="1.0"?><x:styleSheet xmlns:x="` + spreadsheetMLTransitional + `">` +
		`<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></x:numFmts>` +
		`<x:fonts count="2"><x:font><x:name val="Calibri"/></x:font><x:font><x:b/></x:font></x:fonts>` +
		`<x:fills count="2"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="gray125"/></x:fill></x:fills>` +
		`<x:borders count="1"><x:border/></x:borders>` +
		`<x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></x:cellStyleXfs>` +
		`<x:cellXfs count="2"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><x:xf numFmtId="164" fontId="1" fillId="0" borderId="0" xfId="0"/></x:cellXfs>` +
		`<x:cellStyles count="1"><x:cellStyle name="Normal" xfId="0" builtinId="0"/></x:cellStyles>` +
		`</x:styleSheet>`)

	index, err := parseStyleTable(styles)
	if err != nil {
		t.Fatal(err)
	}
	if index.namespace != spreadsheetMLTransitional || len(index.numFmts.entries) != 1 || len(index.fonts.entries) != 2 || len(index.fills.entries) != 2 || len(index.borders.entries) != 1 || len(index.cellStyleXfs.entries) != 1 || len(index.cellXfs.entries) != 2 {
		t.Fatalf("unexpected style table index: %+v", index)
	}
	if got := string(styles[index.fonts.entries[1].span.start:index.fonts.entries[1].span.end]); got != `<x:font><x:b/></x:font>` {
		t.Fatalf("font span did not preserve source bytes: %s", got)
	}
}

func TestCellXFApplyFlagsUseComponentSpecificDefaults(t *testing.T) {
	type componentCase struct {
		name       string
		baseAttrs  string
		cellAttrs  string
		baseChild  string
		cellChild  string
		applyName  string
		errorMatch string
	}
	cases := []componentCase{
		{name: "font", baseAttrs: `fontId="0" fillId="0" numFmtId="0"`, cellAttrs: `fontId="1" fillId="0" numFmtId="0"`, applyName: "applyFont", errorMatch: "font id differs"},
		{name: "fill", baseAttrs: `fontId="0" fillId="0" numFmtId="0"`, cellAttrs: `fontId="0" fillId="2" numFmtId="0"`, applyName: "applyFill", errorMatch: "fill id differs"},
		{name: "number format", baseAttrs: `fontId="0" fillId="0" numFmtId="0"`, cellAttrs: `fontId="0" fillId="0" numFmtId="164"`, applyName: "applyNumberFormat", errorMatch: "number format id differs"},
		{name: "alignment", baseAttrs: `fontId="0" fillId="0" numFmtId="0"`, cellAttrs: `fontId="0" fillId="0" numFmtId="0"`, baseChild: `<x:alignment vertical="bottom"/>`, cellChild: `<x:alignment horizontal="left"/>`, applyName: "applyAlignment", errorMatch: "alignment differs"},
	}
	for _, test := range cases {
		t.Run(test.name, func(t *testing.T) {
			for _, flag := range []struct {
				name      string
				attribute string
				wantError bool
			}{
				{name: "absent", wantError: true},
				{name: "false", attribute: ` ` + test.applyName + `="0"`, wantError: true},
				{name: "true", attribute: ` ` + test.applyName + `="1"`},
			} {
				t.Run(flag.name, func(t *testing.T) {
					xml := applyFlagStyleTable(test.baseAttrs, test.cellAttrs+flag.attribute, test.baseChild, test.cellChild)
					_, err := newStyleRegistry([]byte(xml))
					if flag.wantError {
						if err == nil || !strings.Contains(err.Error(), test.errorMatch) {
							t.Fatalf("expected %q refusal, got %v", test.errorMatch, err)
						}
						return
					}
					if err != nil {
						t.Fatalf("apply=true should accept the direct component: %v", err)
					}
				})
			}
		})
	}
}

func TestCellStyleXFComponentsDefaultToApplied(t *testing.T) {
	xml := applyFlagStyleTable(`fontId="1" fillId="2" numFmtId="164"`, `fontId="1" fillId="2" numFmtId="164"`, `<x:alignment horizontal="right"/>`, "")
	registry, err := newStyleRegistry([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	base := registry.styleXfs[0]
	cell := registry.cellXfs[0]
	if effectiveStyleComponent(cell.fontID, base.fontID, cell.applyFont) != 1 ||
		effectiveStyleComponent(cell.fillID, base.fillID, cell.applyFill) != 2 ||
		effectiveStyleComponent(cell.numFmtID, base.numFmtID, cell.applyNumberFormat) != 164 {
		t.Fatal("absent cell apply flags did not inherit the component applied by cellStyleXf")
	}
	if got := effectiveStyleAlignment(cell, base); effectiveHorizontal(got.horizontal, nil) != "right" {
		t.Fatalf("absent cell applyAlignment did not inherit cellStyleXf alignment: %#v", got)
	}
}

func TestCellStyleXFExplicitFalseDisablesParentComponents(t *testing.T) {
	baseAttrs := `fontId="1" fillId="2" numFmtId="164" applyFont="0" applyFill="0" applyNumberFormat="0" applyAlignment="0"`
	cellAttrs := `fontId="0" fillId="0" numFmtId="0"`
	xml := applyFlagStyleTable(baseAttrs, cellAttrs, `<x:alignment horizontal="right"/>`, "")
	registry, err := newStyleRegistry([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	effective := effectiveCellStyleXF(registry.styleXfs[0])
	if effective.fontID != 0 || effective.fillID != 0 || effective.numFmtID != 0 || effective.alignment.present {
		t.Fatalf("explicit false cellStyleXf apply flags did not disable their components: %#v", effective)
	}
	resolved, err := registry.resolveStyle(0, StyleDelta{
		Bold:                ClearStyleProperty[bool](),
		FillColor:           ClearStyleProperty[string](),
		NumberFormat:        ClearStyleProperty[string](),
		HorizontalAlignment: ClearStyleProperty[HorizontalAlignment](),
	})
	if err != nil {
		t.Fatal(err)
	}
	if resolved != 0 || registry.changed() {
		t.Fatal("clearing already-disabled parent components should be a semantic no-op")
	}
}

func TestExplicitBuiltInNumberFormatIDOverridesFallbackWithoutAliasing(t *testing.T) {
	xml := applyFlagStyleTable(`fontId="0" fillId="0" numFmtId="0"`, `fontId="0" fillId="0" numFmtId="14" applyNumberFormat="1"`, "", "")
	xml = strings.Replace(xml, `<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="custom"/></x:numFmts>`, `<x:numFmts count="1"><x:numFmt numFmtId="14" formatCode="yyyy/mm/dd"/></x:numFmts>`, 1)
	registry, err := newStyleRegistry([]byte(xml))
	if err != nil {
		t.Fatal(err)
	}
	if got, found := registry.numberFormatCode(14); !found || got != "yyyy/mm/dd" {
		t.Fatalf("explicit built-in-range override = %q/%v", got, found)
	}
	if id, found := registry.numFmtByCode["mm-dd-yy"]; found && id == 14 {
		t.Fatal("overridden built-in fallback remained aliased to authored ID 14")
	}
	id, changed, err := registry.resolveNumberFormat(registry.cellXfs[0], effectiveCellStyleXF(registry.styleXfs[0]), SetStyleProperty("mm-dd-yy"))
	if err != nil || !changed || id < 164 {
		t.Fatalf("fallback code did not allocate a non-colliding custom ID: id=%d changed=%v err=%v", id, changed, err)
	}
}

func applyFlagStyleTable(baseAttrs, cellAttrs, baseChild, cellChild string) string {
	base := `<x:xf ` + baseAttrs + ` borderId="0">` + baseChild + `</x:xf>`
	cell := `<x:xf ` + cellAttrs + ` borderId="0" xfId="0">` + cellChild + `</x:xf>`
	return `<x:styleSheet xmlns:x="` + spreadsheetMLTransitional + `">` +
		`<x:numFmts count="1"><x:numFmt numFmtId="164" formatCode="custom"/></x:numFmts>` +
		`<x:fonts count="2"><x:font><x:name val="Base"/></x:font><x:font><x:name val="Direct"/></x:font></x:fonts>` +
		`<x:fills count="3"><x:fill><x:patternFill patternType="none"/></x:fill><x:fill><x:patternFill patternType="gray125"/></x:fill><x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FF112233"/></x:patternFill></x:fill></x:fills>` +
		`<x:borders count="1"><x:border/></x:borders>` +
		`<x:cellStyleXfs count="1">` + base + `</x:cellStyleXfs>` +
		`<x:cellXfs count="1">` + cell + `</x:cellXfs>` +
		`</x:styleSheet>`
}

func TestStyleTableRecordBudgetRefusesBeforeAppendingPastCumulativeLimit(t *testing.T) {
	total := maxStyleTableRecords
	if err := claimStyleTableRecord(&total); err == nil || !strings.Contains(err.Error(), "cumulative safety limit") {
		t.Fatalf("expected pre-append cumulative-record refusal, got %v", err)
	}
	if total != maxStyleTableRecords {
		t.Fatal("refused style record changed the running total")
	}
}

func TestStyleRegistryRefusesAppendBeyondCumulativeLimit(t *testing.T) {
	if err := ensureStyleRecordCountCapacity(maxStyleTableRecords); err == nil {
		t.Fatal("expected append refusal at the cumulative style-record limit")
	}
}
