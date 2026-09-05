package xlsxpatch

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"io"
	"strings"
	"testing"
)

func TestRemovePivotDeletesSingleNativeGraph(t *testing.T) {
	orig := buildZip(t, fixtureWorkbook(true))
	withPivot, err := AddPivot(orig, pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	pivot := onlyHydratedPivot(t, withPivot)
	before := archiveContents(t, withPivot)

	out, err := RemovePivot(withPivot, pivot.Identity)
	if err != nil {
		t.Fatal(err)
	}
	if got, err := ReadPivots(out); err != nil || len(got) != 0 {
		t.Fatalf("pivots after delete = %+v, %v", got, err)
	}
	after := archiveContents(t, out)
	for _, removed := range []string{
		"xl/pivotTables/pivotTable1.xml",
		"xl/pivotTables/_rels/pivotTable1.xml.rels",
		"xl/pivotCache/pivotCacheDefinition1.xml",
		"xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
		"xl/pivotCache/pivotCacheRecords1.xml",
	} {
		if _, exists := after[removed]; exists {
			t.Errorf("removed graph part remains: %s", removed)
		}
	}
	for _, untouched := range []string{"xl/charts/chart1.xml", "xl/drawings/drawing1.xml", "xl/styles.xml", "docProps/core.xml"} {
		if !bytes.Equal(after[untouched], before[untouched]) {
			t.Errorf("unrelated part changed: %s", untouched)
		}
	}
	if strings.Contains(string(after["[Content_Types].xml"]), "pivot") {
		t.Errorf("pivot content-type overrides remain: %s", after["[Content_Types].xml"])
	}
}

func TestReadPivotsJSONCarriesWireCompatibleStableIdentity(t *testing.T) {
	withPivot, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	pivot := onlyHydratedPivot(t, withPivot)
	encoded, err := json.Marshal(pivot)
	if err != nil {
		t.Fatal(err)
	}
	for _, want := range []string{
		`"identity":{"part":"xl/pivotTables/pivotTable1.xml"}`,
		`"dataFields":[{"field":"Profit","agg":"sum"}`,
	} {
		if !bytes.Contains(encoded, []byte(want)) {
			t.Errorf("hydrated JSON missing %s: %s", want, encoded)
		}
	}
}

func TestRemovePivotDeletesOneOfTwoAndPreservesUnrelatedPivot(t *testing.T) {
	one, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	secondSpec := pivotSpec()
	secondSpec.Name = "SecondPivot"
	secondSpec.TargetCellRef = "F20"
	two, err := AddPivot(one, secondSpec)
	if err != nil {
		t.Fatal(err)
	}
	pivots, err := ReadPivots(two)
	if err != nil || len(pivots) != 2 {
		t.Fatalf("hydrate two pivots: %+v, %v", pivots, err)
	}
	before := archiveContents(t, two)
	out, err := RemovePivot(two, pivots[0].Identity)
	if err != nil {
		t.Fatal(err)
	}
	remaining := onlyHydratedPivot(t, out)
	if remaining.Part != pivots[1].Part || remaining.Name != "SecondPivot" {
		t.Fatalf("wrong pivot remained: %+v", remaining)
	}
	after := archiveContents(t, out)
	for _, part := range []string{pivots[1].Part, relsPartFor(pivots[1].Part), "xl/pivotCache/pivotCacheDefinition2.xml", "xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels", "xl/pivotCache/pivotCacheRecords2.xml"} {
		if !bytes.Equal(after[part], before[part]) {
			t.Errorf("unrelated pivot graph changed: %s", part)
		}
	}
	if err := RequirePreservation(two, out, []string{
		pivots[0].Part,
		relsPartFor(pivots[0].Part),
		"xl/pivotCache/pivotCacheDefinition1.xml",
		"xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels",
		"xl/pivotCache/pivotCacheRecords1.xml",
	}); err != nil {
		t.Fatalf("unrelated preservation inventory changed: %v", err)
	}
}

func TestRemovePivotKeepsSharedCacheUntilLastUser(t *testing.T) {
	shared := sharedPivotCacheWorkbook(t)
	pivots, err := ReadPivots(shared)
	if err != nil || len(pivots) != 2 {
		t.Fatalf("hydrate shared pivots: %+v, %v", pivots, err)
	}
	cacheBefore := archiveContents(t, shared)["xl/pivotCache/pivotCacheDefinition1.xml"]
	oneRemoved, err := RemovePivot(shared, pivots[0].Identity)
	if err != nil {
		t.Fatal(err)
	}
	afterOne := archiveContents(t, oneRemoved)
	if !bytes.Equal(afterOne["xl/pivotCache/pivotCacheDefinition1.xml"], cacheBefore) {
		t.Fatal("shared cache changed when one table was removed")
	}
	remaining := onlyHydratedPivot(t, oneRemoved)
	if remaining.Name != "SecondPivot" {
		t.Fatalf("wrong shared-cache pivot remained: %+v", remaining)
	}
	allRemoved, err := RemovePivot(oneRemoved, remaining.Identity)
	if err != nil {
		t.Fatal(err)
	}
	if _, exists := archiveContents(t, allRemoved)["xl/pivotCache/pivotCacheDefinition1.xml"]; exists {
		t.Fatal("last cache user removed but cache definition remains")
	}
}

func TestUpdatePivotRetainsStablePartAndReplacesOnlySelectedPivot(t *testing.T) {
	one, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	secondSpec := pivotSpec()
	secondSpec.Name = "SecondPivot"
	secondSpec.TargetCellRef = "F20"
	two, err := AddPivot(one, secondSpec)
	if err != nil {
		t.Fatal(err)
	}
	pivots, _ := ReadPivots(two)
	secondBefore := archiveContents(t, two)[pivots[1].Part]

	replacement := pivotSpec()
	replacement.Name = "UpdatedProfit"
	replacement.TargetCellRef = "J4"
	replacement.DataFields = []PivotDataField{{Field: "Revenue", Agg: "max"}}
	selected := "100"
	replacement.PageFields = []PivotPageField{{Field: "Revenue", SelectedItem: &selected}}
	replacement.MemberFilters = []PivotMemberFilter{{Field: "Costs", ExcludedItems: []string{"30"}}}
	replacement.Sorts = []PivotFieldSort{{Field: "Quarter", Direction: "ascending"}}
	replacement.FieldMembers = []PivotFieldMembers{
		{Field: "Revenue", Items: []PivotFieldMember{{Value: "100", Kind: "number"}, {Value: "200", Kind: "number"}}},
		{Field: "Costs", Items: []PivotFieldMember{{Value: "20", Kind: "number"}, {Value: "30", Kind: "number"}}},
	}
	out, err := UpdatePivot(two, pivots[0].Identity, replacement)
	if err != nil {
		t.Fatal(err)
	}
	updated, err := ReadPivots(out)
	if err != nil || len(updated) != 2 {
		t.Fatalf("hydrate updated pivots: %+v, %v", updated, err)
	}
	if updated[0].Part != pivots[0].Part || updated[0].Identity != pivots[0].Identity || updated[0].Name != "UpdatedProfit" || updated[0].TargetRef != "J4" {
		t.Fatalf("replacement identity/content mismatch: %+v", updated[0])
	}
	if len(updated[0].PageFields) != 1 || len(updated[0].MemberFilters) != 1 || len(updated[0].Sorts) != 1 {
		t.Fatalf("replacement page/filter/sort state was not preserved: %+v", updated[0])
	}
	if !bytes.Equal(archiveContents(t, out)[pivots[1].Part], secondBefore) {
		t.Fatal("unselected pivot table changed during update")
	}
}

func TestUpdateSharedCacheForksReplacementCache(t *testing.T) {
	shared := sharedPivotCacheWorkbook(t)
	pivots, _ := ReadPivots(shared)
	cacheBefore := archiveContents(t, shared)["xl/pivotCache/pivotCacheDefinition1.xml"]
	replacement := pivotSpec()
	replacement.Name = "Forked"
	replacement.SourceRef = "A1:C5"
	replacement.Fields = []string{"Quarter", "Revenue", "Costs"}
	replacement.DataFields = []PivotDataField{{Field: "Revenue", Agg: "sum"}}
	out, err := UpdatePivot(shared, pivots[0].Identity, replacement)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(archiveContents(t, out)["xl/pivotCache/pivotCacheDefinition1.xml"], cacheBefore) {
		t.Fatal("existing shared cache changed during update")
	}
	updated, _ := ReadPivots(out)
	if len(updated) != 2 || updated[0].Part != pivots[0].Part || updated[0].CacheID == updated[1].CacheID {
		t.Fatalf("replacement was not forked onto a private cache: %+v", updated)
	}
}

func TestPivotLifecycleRejectsStaleAndMalformedIdentity(t *testing.T) {
	withPivot, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	pivot := onlyHydratedPivot(t, withPivot)
	for _, identity := range []PivotIdentity{{}, {Part: "xl/pivotTables/missing.xml"}, {Part: "../pivotTable1.xml"}} {
		if _, err := RemovePivot(withPivot, identity); err == nil {
			t.Errorf("identity %+v should fail", identity)
		}
	}

	parts := archiveContents(t, withPivot)
	malformed, err := Apply(withPivot, Patch{Replace: map[string][]byte{
		relsPartFor(pivot.Part): []byte(`<Relationships><Relationship Id="rId1" Type="` + relTypePivotCacheDef + `" Target="../pivotCache/pivotCacheDefinition1.xml"></Relationships>`),
	}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := RemovePivot(malformed, pivot.Identity); err == nil {
		t.Fatal("malformed pivot relationships should fail closed")
	}
	if !bytes.Equal(parts[pivot.Part], archiveContents(t, malformed)[pivot.Part]) {
		t.Fatal("fixture setup unexpectedly changed pivot table")
	}
	first, err := RemovePivot(withPivot, pivot.Identity)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := RemovePivot(first, pivot.Identity); err == nil {
		t.Fatal("repeated remove should report a stale identity")
	}
}

func TestRemovePivotRejectsAmbiguousWorksheetOwnership(t *testing.T) {
	withPivot, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	pivot := onlyHydratedPivot(t, withPivot)
	parts := archiveContents(t, withPivot)
	rels := string(parts["xl/worksheets/_rels/sheet1.xml.rels"])
	ambiguous, err := appendRelationship(rels, "rId99", relTypePivotTable, "../pivotTables/pivotTable1.xml")
	if err != nil {
		t.Fatal(err)
	}
	workbook := mustApply(t, withPivot, Patch{Replace: map[string][]byte{
		"xl/worksheets/_rels/sheet1.xml.rels": []byte(ambiguous),
	}})
	if _, err := RemovePivot(workbook, pivot.Identity); err == nil || !strings.Contains(err.Error(), "2 worksheet owners") {
		t.Fatalf("ambiguous owner should fail closed, got %v", err)
	}
}

func onlyHydratedPivot(t *testing.T, data []byte) PivotInfo {
	t.Helper()
	pivots, err := ReadPivots(data)
	if err != nil || len(pivots) != 1 {
		t.Fatalf("expected one pivot, got %+v, %v", pivots, err)
	}
	return pivots[0]
}

func archiveContents(t *testing.T, data []byte) map[string][]byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	out := make(map[string][]byte, len(zr.File))
	for _, file := range zr.File {
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		value, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		out[file.Name] = value
	}
	return out
}

func sharedPivotCacheWorkbook(t *testing.T) []byte {
	t.Helper()
	one, err := AddPivot(buildZip(t, fixtureWorkbook(false)), pivotSpec())
	if err != nil {
		t.Fatal(err)
	}
	secondSpec := pivotSpec()
	secondSpec.Name = "SecondPivot"
	secondSpec.TargetCellRef = "F20"
	two, err := AddPivot(one, secondSpec)
	if err != nil {
		t.Fatal(err)
	}
	parts := archiveContents(t, two)
	workbook, err := collectionWithoutReference(parts["xl/workbook.xml"], "workbook", "pivotCaches", "pivotCache", "id", "rId3")
	if err != nil {
		t.Fatal(err)
	}
	workbookRels, err := relationshipDocumentWithout(parts["xl/_rels/workbook.xml.rels"], "rId3", relTypePivotCacheDef, "xl/workbook.xml", "xl/pivotCache/pivotCacheDefinition2.xml")
	if err != nil {
		t.Fatal(err)
	}
	contentTypes, err := contentTypesWithoutParts(parts["[Content_Types].xml"], []string{"xl/pivotCache/pivotCacheDefinition2.xml", "xl/pivotCache/pivotCacheRecords2.xml"})
	if err != nil {
		t.Fatal(err)
	}
	return mustApply(t, two, Patch{
		Replace: map[string][]byte{
			"xl/pivotTables/pivotTable2.xml":            bytes.Replace(parts["xl/pivotTables/pivotTable2.xml"], []byte(`cacheId="2"`), []byte(`cacheId="1"`), 1),
			"xl/pivotTables/_rels/pivotTable2.xml.rels": bytes.Replace(parts["xl/pivotTables/_rels/pivotTable2.xml.rels"], []byte("pivotCacheDefinition2.xml"), []byte("pivotCacheDefinition1.xml"), 1),
			"xl/workbook.xml":                           workbook,
			"xl/_rels/workbook.xml.rels":                workbookRels,
			"[Content_Types].xml":                       contentTypes,
		},
		Delete: map[string]bool{
			"xl/pivotCache/pivotCacheDefinition2.xml":            true,
			"xl/pivotCache/_rels/pivotCacheDefinition2.xml.rels": true,
			"xl/pivotCache/pivotCacheRecords2.xml":               true,
		},
	})
}

func mustApply(t *testing.T, orig []byte, patch Patch) []byte {
	t.Helper()
	out, err := Apply(orig, patch)
	if err != nil {
		t.Fatal(err)
	}
	return out
}
