package xlsxpatch

import (
	"strings"
	"testing"
)

func pivotSpec() PivotWriteSpec {
	return PivotWriteSpec{
		SourceSheetName: "Data",
		SourceRef:       "A1:D5",
		Fields:          []string{"Quarter", "Revenue", "Costs", "Profit"},
		TargetSheetName: "Data",
		TargetCellRef:   "F1",
		RowFields:       []string{"Quarter"},
		DataFields:      []PivotDataField{{Field: "Profit", Agg: "sum"}, {Field: "Revenue", Agg: "avg"}},
		Name:            "ProfitPivot",
	}
}

func TestAddPivot_WritesAllParts(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := AddPivot(orig, pivotSpec())
	if err != nil {
		t.Fatal(err)
	}

	ins, err := Inspect(out)
	if err != nil {
		t.Fatal(err)
	}
	if !ins.HasPivots() {
		t.Fatalf("no pivot parts detected: %+v", ins)
	}

	def := readEntry(t, out, "xl/pivotCache/pivotCacheDefinition1.xml")
	for _, want := range []string{`refreshOnLoad="1"`, `ref="A1:D5"`, `sheet="Data"`, `cacheFields count="4"`, `name="Quarter"`} {
		if !strings.Contains(def, want) {
			t.Errorf("cache definition missing %s", want)
		}
	}
	if got := readEntry(t, out, "xl/pivotCache/pivotCacheRecords1.xml"); !strings.Contains(got, `count="0"`) {
		t.Errorf("records part should be empty: %s", got)
	}

	table := readEntry(t, out, "xl/pivotTables/pivotTable1.xml")
	for _, want := range []string{
		`name="ProfitPivot"`, `cacheId="1"`, `ref="F1"`,
		`<rowFields count="1"><field x="0"/></rowFields>`,
		`<dataField name="Sum of Profit" fld="3"`,
		`subtotal="average"`,
		`axis="axisRow"`,
	} {
		if !strings.Contains(table, want) {
			t.Errorf("pivot table missing %s\n%s", want, table)
		}
	}

	wb := readEntry(t, out, "xl/workbook.xml")
	if !strings.Contains(wb, "<pivotCaches><pivotCache ") || !strings.Contains(wb, `cacheId="1"`) {
		t.Errorf("workbook missing pivotCaches: %s", wb)
	}
	wbRels := readEntry(t, out, "xl/_rels/workbook.xml.rels")
	if !strings.Contains(wbRels, "pivotCache/pivotCacheDefinition1.xml") {
		t.Errorf("workbook rels missing cache def: %s", wbRels)
	}
	sheetRels := readEntry(t, out, "xl/worksheets/_rels/sheet1.xml.rels")
	if !strings.Contains(sheetRels, "../pivotTables/pivotTable1.xml") {
		t.Errorf("sheet rels missing pivot table: %s", sheetRels)
	}
	sheet := readEntry(t, out, "xl/worksheets/sheet1.xml")
	if !strings.Contains(sheet, `<pivotTableParts count="1">`) || !strings.Contains(sheet, `r:id="rId1"`) {
		t.Errorf("worksheet missing pivotTableParts reference: %s", sheet)
	}
	ctypes := readEntry(t, out, "[Content_Types].xml")
	for _, want := range []string{"pivotCacheDefinition1.xml", "pivotCacheRecords1.xml", "pivotTable1.xml"} {
		if !strings.Contains(ctypes, want) {
			t.Errorf("content types missing %s", want)
		}
	}
}

func TestAddPivot_SecondPivotGetsNextIDs(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	one, err := AddPivot(orig, pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	spec2 := pivotSpec()
	spec2.Name = "SecondPivot"
	spec2.TargetCellRef = "F20"
	two, err := AddPivot(one, spec2)
	if err != nil {
		t.Fatal(err)
	}
	table2 := readEntry(t, two, "xl/pivotTables/pivotTable2.xml")
	if !strings.Contains(table2, `cacheId="2"`) {
		t.Errorf("second pivot should get cacheId 2: %s", table2)
	}
	wb := readEntry(t, two, "xl/workbook.xml")
	if strings.Count(wb, "<pivotCache ") != 2 {
		t.Errorf("workbook should list 2 caches: %s", wb)
	}
	if strings.Count(wb, "<pivotCaches>") != 1 {
		t.Errorf("pivotCaches element must not duplicate: %s", wb)
	}
	sheet := readEntry(t, two, "xl/worksheets/sheet1.xml")
	if strings.Count(sheet, "<pivotTablePart ") != 2 || !strings.Contains(sheet, `pivotTableParts count="2"`) {
		t.Errorf("worksheet should reference 2 pivots: %s", sheet)
	}
}

func TestReadPivots_HydratesWrittenPivot(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	out, err := AddPivot(orig, pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	pivots, err := ReadPivots(out)
	if err != nil {
		t.Fatal(err)
	}
	if len(pivots) != 1 {
		t.Fatalf("got %d pivots: %+v", len(pivots), pivots)
	}
	got := pivots[0]
	if got.Name != "ProfitPivot" || got.SourceSheetName != "Data" || got.SourceRef != "A1:D5" || got.TargetSheetName != "Data" || got.TargetRef != "F1" {
		t.Errorf("identity/ranges not hydrated: %+v", got)
	}
	if strings.Join(got.Fields, ",") != "Quarter,Revenue,Costs,Profit" || strings.Join(got.RowFields, ",") != "Quarter" {
		t.Errorf("fields not hydrated: %+v", got)
	}
	if len(got.DataFields) != 2 || got.DataFields[0] != (PivotDataField{Field: "Profit", Agg: "sum"}) || got.DataFields[1] != (PivotDataField{Field: "Revenue", Agg: "avg"}) {
		t.Errorf("data fields not hydrated: %+v", got.DataFields)
	}
	if !got.GrandTotals || len(got.Warnings) != 0 {
		t.Errorf("unexpected flags: %+v", got)
	}
}

func TestPivotPageMemberFilterSortRoundTrip(t *testing.T) {
	spec := pivotSpec()
	selected := "100"
	noTotals := false
	spec.PageFields = []PivotPageField{{Field: "Revenue", SelectedItem: &selected}, {Field: "Costs"}}
	spec.MemberFilters = []PivotMemberFilter{{Field: "Costs", ExcludedItems: []string{"30"}}}
	spec.Sorts = []PivotFieldSort{{Field: "Quarter", Direction: "descending"}}
	spec.FieldMembers = []PivotFieldMembers{
		{Field: "Revenue", Items: []PivotFieldMember{{Value: "100", Kind: "number"}, {Value: "200", Kind: "number"}}},
		{Field: "Costs", Items: []PivotFieldMember{{Value: "20", Kind: "number"}, {Value: "30", Kind: "number"}}},
	}
	spec.GrandTotals = &noTotals

	out, err := AddPivot(buildZip(t, fixtureWorkbook(false)), spec)
	if err != nil {
		t.Fatal(err)
	}
	table := readEntry(t, out, "xl/pivotTables/pivotTable1.xml")
	for _, want := range []string{`axis="axisPage"`, `sortType="descending"`, `multipleItemSelectionAllowed="1"`, `<item x="1" h="1"/>`, `<pageField fld="1" hier="-1" item="0"/>`, `<pageField fld="2" hier="-1"/>`, `rowGrandTotals="0" colGrandTotals="0"`} {
		if !strings.Contains(table, want) {
			t.Errorf("pivot table missing %s\n%s", want, table)
		}
	}
	pivots, err := ReadPivots(out)
	if err != nil || len(pivots) != 1 {
		t.Fatalf("hydrate configured pivot: %+v, %v", pivots, err)
	}
	got := pivots[0]
	if got.GrandTotals || len(got.Warnings) != 0 {
		t.Fatalf("unexpected hydration diagnostics: %+v", got)
	}
	if len(got.PageFields) != 2 || got.PageFields[0].Field != "Revenue" || got.PageFields[0].SelectedItem == nil || *got.PageFields[0].SelectedItem != "100" || got.PageFields[1].Field != "Costs" || got.PageFields[1].SelectedItem != nil {
		t.Errorf("page field not preserved: %+v", got.PageFields)
	}
	if len(got.MemberFilters) != 1 || got.MemberFilters[0].Field != "Costs" || strings.Join(got.MemberFilters[0].ExcludedItems, ",") != "30" {
		t.Errorf("member filter not preserved: %+v", got.MemberFilters)
	}
	if len(got.Sorts) != 1 || got.Sorts[0] != (PivotFieldSort{Field: "Quarter", Direction: "descending"}) {
		t.Errorf("sort not preserved: %+v", got.Sorts)
	}
}

func TestAddPivotRejectsUnrepresentableFilterState(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	for name, mutate := range map[string]func(*PivotWriteSpec){
		"missing member inventory": func(spec *PivotWriteSpec) {
			spec.MemberFilters = []PivotMemberFilter{{Field: "Costs", ExcludedItems: []string{"30"}}}
		},
		"unknown page member": func(spec *PivotWriteSpec) {
			selected := "missing"
			spec.PageFields = []PivotPageField{{Field: "Costs", SelectedItem: &selected}}
			spec.FieldMembers = []PivotFieldMembers{{Field: "Costs", Items: []PivotFieldMember{{Value: "30", Kind: "number"}}}}
		},
		"single page selection plus filter": func(spec *PivotWriteSpec) {
			selected := "30"
			spec.PageFields = []PivotPageField{{Field: "Costs", SelectedItem: &selected}}
			spec.MemberFilters = []PivotMemberFilter{{Field: "Costs", ExcludedItems: []string{"20"}}}
			spec.FieldMembers = []PivotFieldMembers{{Field: "Costs", Items: []PivotFieldMember{{Value: "20", Kind: "number"}, {Value: "30", Kind: "number"}}}}
		},
		"value sort": func(spec *PivotWriteSpec) {
			spec.Sorts = []PivotFieldSort{{Field: "Profit", Direction: "descending"}}
		},
		"unsupported member kind": func(spec *PivotWriteSpec) {
			spec.FieldMembers = []PivotFieldMembers{{Field: "Costs", Items: []PivotFieldMember{{Value: "2026-01-01", Kind: "date"}}}}
		},
	} {
		t.Run(name, func(t *testing.T) {
			spec := pivotSpec()
			mutate(&spec)
			if _, err := AddPivot(orig, spec); err == nil {
				t.Fatal("unrepresentable pivot state was accepted")
			}
		})
	}
}

func TestReadPivotsWarnsOnUnsupportedNativePivotFilters(t *testing.T) {
	out, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	table := readEntry(t, out, "xl/pivotTables/pivotTable1.xml")
	table = strings.Replace(table, `<dataFields`, `<pivotFilters count="1"><filter fld="0" type="captionEqual"/></pivotFilters><dataFields`, 1)
	warned, err := Apply(out, Patch{Replace: map[string][]byte{"xl/pivotTables/pivotTable1.xml": []byte(table)}})
	if err != nil {
		t.Fatal(err)
	}
	pivots, err := ReadPivots(warned)
	if err != nil || len(pivots) != 1 || len(pivots[0].Warnings) == 0 || !strings.Contains(strings.Join(pivots[0].Warnings, ";"), "value and label") {
		t.Fatalf("unsupported pivot filter warning missing: %+v, %v", pivots, err)
	}
}

func TestAddPivot_Rejections(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(false))
	bad := pivotSpec()
	bad.RowFields = []string{"NoSuchField"}
	if _, err := AddPivot(orig, bad); err == nil {
		t.Error("unknown row field must be rejected")
	}
	bad2 := pivotSpec()
	bad2.DataFields = []PivotDataField{{Field: "Profit", Agg: "median"}}
	if _, err := AddPivot(orig, bad2); err == nil {
		t.Error("unsupported aggregation must be rejected")
	}
	bad3 := pivotSpec()
	bad3.TargetSheetName = "Ghost"
	if _, err := AddPivot(orig, bad3); err == nil {
		t.Error("unknown target sheet must be rejected")
	}
	bad4 := pivotSpec()
	bad4.DataFields = nil
	if _, err := AddPivot(orig, bad4); err == nil {
		t.Error("no data fields must be rejected")
	}
}
