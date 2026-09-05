package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestReadChartsHydratesStableNativeIdentity(t *testing.T) {
	workbook, err := AddChart(buildZip(t, fixtureWorkbook(false)), writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}
	charts, err := ReadCharts(workbook)
	if err != nil || len(charts) != 1 || charts[0].Identity == nil {
		t.Fatalf("charts: %+v, %v", charts, err)
	}
	want := ChartIdentity{Part: "xl/charts/chart1.xml", DrawingPart: "xl/drawings/drawing1.xml", ObjectID: 1}
	if *charts[0].Identity != want {
		t.Fatalf("identity = %+v, want %+v", *charts[0].Identity, want)
	}
	identities, err := ReadChartIdentities(workbook)
	if err != nil || len(identities) != 1 || identities[0] != want {
		t.Fatalf("identities: %+v, %v", identities, err)
	}
}

func TestUpdateChartRetainsIdentityAndPreservesSiblingGraph(t *testing.T) {
	one, _ := AddChart(buildZip(t, fixtureWorkbook(false)), writeSpec("column"))
	secondSpec := writeSpec("line")
	secondSpec.Title = "Keep me"
	secondSpec.Anchor = ChartAnchor{FromCol: 14, FromRow: 2, ToCol: 22, ToRow: 18}
	two, err := AddChart(one, secondSpec)
	if err != nil {
		t.Fatal(err)
	}
	charts, _ := ReadCharts(two)
	identity := *charts[0].Identity
	siblingIdentity := *charts[1].Identity
	before := archiveContents(t, two)
	siblingAnchor := chartAnchorBytes(t, before[siblingIdentity.DrawingPart], siblingIdentity.ObjectID)
	siblingPart := bytes.Clone(before[siblingIdentity.Part])

	replacement := writeSpec("doughnut")
	replacement.Title = "Updated"
	replacement.Anchor = ChartAnchor{FromCol: 3, FromRow: 8, ToCol: 11, ToRow: 24}
	out, err := UpdateChart(two, identity, replacement)
	if err != nil {
		t.Fatal(err)
	}
	updated, err := ReadCharts(out)
	if err != nil || len(updated) != 2 {
		t.Fatalf("updated charts: %+v, %v", updated, err)
	}
	if updated[0].Identity == nil || *updated[0].Identity != identity || updated[0].Type != "doughnut" || updated[0].Title != "Updated" {
		t.Fatalf("target mismatch: %+v", updated[0])
	}
	after := archiveContents(t, out)
	if !bytes.Equal(siblingPart, after[siblingIdentity.Part]) {
		t.Fatal("sibling chart part changed")
	}
	if !bytes.Equal(siblingAnchor, chartAnchorBytes(t, after[siblingIdentity.DrawingPart], siblingIdentity.ObjectID)) {
		t.Fatal("sibling chart anchor changed")
	}
	if !bytes.Equal(before["xl/drawings/_rels/drawing1.xml.rels"], after["xl/drawings/_rels/drawing1.xml.rels"]) {
		t.Fatal("drawing relationships changed during update")
	}
}

func TestRemoveChartDeletesOnlyTargetNativeGraph(t *testing.T) {
	one, _ := AddChart(buildZip(t, fixtureWorkbook(false)), writeSpec("column"))
	shapeSpec := rectSpec()
	shapeSpec.Text = "Keep shape"
	withShape, _ := AddShape(one, shapeSpec)
	secondSpec := writeSpec("line")
	secondSpec.Title = "Keep chart"
	secondSpec.Anchor = ChartAnchor{FromCol: 14, FromRow: 2, ToCol: 22, ToRow: 18}
	workbook, err := AddChart(withShape, secondSpec)
	if err != nil {
		t.Fatal(err)
	}
	charts, _ := ReadCharts(workbook)
	target := *charts[0].Identity
	sibling := *charts[1].Identity
	shape := onlyHydratedShape(t, workbook)
	before := archiveContents(t, workbook)
	siblingAnchor := chartAnchorBytes(t, before[sibling.DrawingPart], sibling.ObjectID)
	shapeAnchor := shapeAnchorBytes(t, before[shape.Identity.DrawingPart], shape.Identity.ObjectID)

	out, err := RemoveChart(workbook, target)
	if err != nil {
		t.Fatal(err)
	}
	remaining, err := ReadCharts(out)
	if err != nil || len(remaining) != 1 || remaining[0].Identity == nil || *remaining[0].Identity != sibling {
		t.Fatalf("remaining charts: %+v, %v", remaining, err)
	}
	after := archiveContents(t, out)
	if _, exists := after[target.Part]; exists {
		t.Fatal("target chart part still exists")
	}
	if bytes.Contains(after["[Content_Types].xml"], []byte("/"+target.Part)) || bytes.Contains(after["xl/drawings/_rels/drawing1.xml.rels"], []byte("../charts/chart1.xml")) {
		t.Fatal("target relationship graph still exists")
	}
	if !bytes.Equal(before[sibling.Part], after[sibling.Part]) || !bytes.Equal(siblingAnchor, chartAnchorBytes(t, after[sibling.DrawingPart], sibling.ObjectID)) {
		t.Fatal("sibling chart changed")
	}
	if !bytes.Equal(shapeAnchor, shapeAnchorBytes(t, after[shape.Identity.DrawingPart], shape.Identity.ObjectID)) {
		t.Fatal("sibling shape changed")
	}
}

func TestChartLifecycleRejectsStaleMismatchedAndDependentGraphs(t *testing.T) {
	workbook, _ := AddChart(buildZip(t, fixtureWorkbook(false)), writeSpec("column"))
	chart := onlyHydratedChart(t, workbook)
	for _, identity := range []ChartIdentity{
		{},
		{Part: "../charts/chart1.xml", DrawingPart: "xl/drawings/drawing1.xml", ObjectID: 1},
		{Part: chart.Part, DrawingPart: "xl/drawings/_rels/drawing1.xml", ObjectID: 1},
		{Part: chart.Part, DrawingPart: "xl/drawings/drawing1.xml", ObjectID: 99},
	} {
		if _, err := RemoveChart(workbook, identity); err == nil {
			t.Errorf("identity %+v should fail", identity)
		}
	}
	wrongSheet := writeSpec("line")
	wrongSheet.SheetName = "Other"
	if _, err := UpdateChart(workbook, *chart.Identity, wrongSheet); err == nil || !strings.Contains(err.Error(), "belongs to worksheet") {
		t.Fatalf("worksheet mismatch should fail, got %v", err)
	}
	removed, err := RemoveChart(workbook, *chart.Identity)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := RemoveChart(removed, *chart.Identity); err == nil {
		t.Fatal("repeated remove should report stale identity")
	}

	parts := archiveContents(t, workbook)
	parts[relsPartFor(chart.Part)] = []byte(relsDoc(relationshipXML("rId1", "http://example.test/dependency", "../embeddings/data.bin")))
	parts["xl/embeddings/data.bin"] = []byte("opaque")
	textParts := make(map[string]string, len(parts))
	for name, value := range parts {
		textParts[name] = string(value)
	}
	withDependency := buildZip(t, textParts)
	if _, err := UpdateChart(withDependency, *chart.Identity, writeSpec("line")); err == nil || !strings.Contains(err.Error(), "dependent relationship graph") {
		t.Fatalf("dependent chart graph should fail closed, got %v", err)
	}
}

func TestChartIdentityRejectsDuplicateDrawingObjectIDs(t *testing.T) {
	one, _ := AddChart(buildZip(t, fixtureWorkbook(false)), writeSpec("column"))
	second := writeSpec("line")
	second.Anchor = ChartAnchor{FromCol: 14, FromRow: 2, ToCol: 22, ToRow: 18}
	two, _ := AddChart(one, second)
	drawing := readEntry(t, two, "xl/drawings/drawing1.xml")
	broken := strings.Replace(drawing, `id="2"`, `id="1"`, 1)
	workbook := mustApply(t, two, Patch{Replace: map[string][]byte{"xl/drawings/drawing1.xml": []byte(broken)}})
	if _, err := ReadChartIdentities(workbook); err == nil || !strings.Contains(err.Error(), "duplicate cNvPr id 1") {
		t.Fatalf("duplicate identity should fail closed, got %v", err)
	}
}

func onlyHydratedChart(t *testing.T, data []byte) ChartInfo {
	t.Helper()
	charts, err := ReadCharts(data)
	if err != nil || len(charts) != 1 || charts[0].Identity == nil {
		t.Fatalf("expected one hydrated chart, got %+v, %v", charts, err)
	}
	return charts[0]
}

func chartAnchorBytes(t *testing.T, drawing []byte, objectID uint32) []byte {
	t.Helper()
	index, err := indexShapeDrawing(drawing)
	if err != nil {
		t.Fatal(err)
	}
	for _, anchor := range index.anchors {
		if anchor.objectID == objectID && anchor.objectType == "graphicFrame" {
			return bytes.Clone(drawing[anchor.span.start:anchor.span.end])
		}
	}
	t.Fatalf("missing chart anchor %d", objectID)
	return nil
}
