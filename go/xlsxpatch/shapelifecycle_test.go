package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

func TestRemoveShapeDeletesOneAnchorAndPreservesSiblingsAndOPCParts(t *testing.T) {
	one, err := AddShape(buildZip(t, fixtureWorkbook(false)), rectSpec())
	if err != nil {
		t.Fatal(err)
	}
	second := rectSpec()
	second.Kind = "ellipse"
	second.Text = "Keep me"
	second.Anchor = ShapeAnchor{FromCol: 10, FromRow: 3, ToCol: 14, ToRow: 7}
	two, err := AddShape(one, second)
	if err != nil {
		t.Fatal(err)
	}
	withChart, err := AddChart(two, writeSpec("column"))
	if err != nil {
		t.Fatal(err)
	}
	shapes, err := ReadShapes(withChart)
	if err != nil || len(shapes) != 2 {
		t.Fatalf("hydrate: %+v, %v", shapes, err)
	}
	before := archiveContents(t, withChart)
	secondAnchor := shapeAnchorBytes(t, before[shapes[1].Identity.DrawingPart], shapes[1].Identity.ObjectID)

	out, err := RemoveShape(withChart, shapes[0].Identity)
	if err != nil {
		t.Fatal(err)
	}
	remaining, err := ReadShapes(out)
	if err != nil || len(remaining) != 1 || remaining[0].Identity != shapes[1].Identity {
		t.Fatalf("remaining shapes: %+v, %v", remaining, err)
	}
	after := archiveContents(t, out)
	if !bytes.Equal(secondAnchor, shapeAnchorBytes(t, after[remaining[0].Identity.DrawingPart], remaining[0].Identity.ObjectID)) {
		t.Fatal("unselected shape anchor changed during remove")
	}
	for _, part := range []string{"xl/charts/chart1.xml", "xl/drawings/_rels/drawing1.xml.rels", "xl/styles.xml", "docProps/core.xml"} {
		if !bytes.Equal(before[part], after[part]) {
			t.Errorf("unrelated part changed: %s", part)
		}
	}
	if err := RequirePreservation(withChart, out, []string{shapes[0].Identity.DrawingPart}); err != nil {
		t.Fatalf("preservation inventory changed: %v", err)
	}
}

func TestUpdateShapeRetainsIdentityAndPreservesSiblingAnchor(t *testing.T) {
	one, _ := AddShape(buildZip(t, fixtureWorkbook(false)), rectSpec())
	second := rectSpec()
	second.Text = "Sibling"
	second.Anchor = ShapeAnchor{FromCol: 10, FromRow: 1, ToCol: 13, ToRow: 4}
	two, err := AddShape(one, second)
	if err != nil {
		t.Fatal(err)
	}
	shapes, _ := ReadShapes(two)
	before := archiveContents(t, two)
	siblingAnchor := shapeAnchorBytes(t, before[shapes[1].Identity.DrawingPart], shapes[1].Identity.ObjectID)
	replacement := rectSpec()
	replacement.Kind = "ellipse"
	replacement.Text = "Updated"
	replacement.Anchor = ShapeAnchor{FromCol: 2, FromRow: 8, ToCol: 8, ToRow: 14}

	out, err := UpdateShape(two, shapes[0].Identity, replacement)
	if err != nil {
		t.Fatal(err)
	}
	updated, err := ReadShapes(out)
	if err != nil || len(updated) != 2 {
		t.Fatalf("updated shapes: %+v, %v", updated, err)
	}
	if updated[0].Identity != shapes[0].Identity || updated[0].Kind != "ellipse" || updated[0].Text != "Updated" || updated[0].Anchor != replacement.Anchor {
		t.Fatalf("updated target mismatch: %+v", updated[0])
	}
	if !bytes.Equal(siblingAnchor, shapeAnchorBytes(t, archiveContents(t, out)[shapes[1].Identity.DrawingPart], shapes[1].Identity.ObjectID)) {
		t.Fatal("sibling anchor changed during update")
	}

	replacement.Text = "Updated again"
	again, err := UpdateShape(out, shapes[0].Identity, replacement)
	if err != nil {
		t.Fatal(err)
	}
	againShapes, _ := ReadShapes(again)
	if againShapes[0].Identity != shapes[0].Identity || againShapes[0].Text != "Updated again" {
		t.Fatalf("repeated update lost identity: %+v", againShapes[0])
	}
}

func TestShapeLifecycleRejectsStaleMalformedAndAmbiguousIdentities(t *testing.T) {
	one, _ := AddShape(buildZip(t, fixtureWorkbook(false)), rectSpec())
	shape := onlyHydratedShape(t, one)
	for _, identity := range []ShapeIdentity{
		{},
		{DrawingPart: "xl/drawings/drawing1.xml", ObjectID: 99},
		{DrawingPart: "../drawings/drawing1.xml", ObjectID: 1},
		{DrawingPart: "xl/drawings/_rels/drawing1.xml", ObjectID: 1},
	} {
		if _, err := RemoveShape(one, identity); err == nil {
			t.Errorf("identity %+v should fail", identity)
		}
	}
	removed, err := RemoveShape(one, shape.Identity)
	if err != nil {
		t.Fatal(err)
	}
	if _, err := RemoveShape(removed, shape.Identity); err == nil {
		t.Fatal("repeated remove should report stale identity")
	}

	second := rectSpec()
	second.Anchor = ShapeAnchor{FromCol: 10, FromRow: 0, ToCol: 12, ToRow: 2}
	two, _ := AddShape(one, second)
	drawing := readEntry(t, two, shape.Identity.DrawingPart)
	ambiguous := strings.Replace(drawing, `id="2"`, `id="1"`, 1)
	broken := mustApply(t, two, Patch{Replace: map[string][]byte{shape.Identity.DrawingPart: []byte(ambiguous)}})
	if _, err := RemoveShape(broken, shape.Identity); err == nil || !strings.Contains(err.Error(), "duplicate cNvPr id 1") {
		t.Fatalf("duplicate native identity should fail closed, got %v", err)
	}
}

func TestConnectorBindingsSurviveUpdatesAndBlockDanglingDelete(t *testing.T) {
	one, _ := AddShape(buildZip(t, fixtureWorkbook(false)), rectSpec())
	second := rectSpec()
	second.Kind = "ellipse"
	second.Anchor = ShapeAnchor{FromCol: 8, FromRow: 2, ToCol: 11, ToRow: 5}
	two, _ := AddShape(one, second)
	line := ShapeWriteSpec{SheetName: "Data", Kind: "line", Stroke: "#59636A", Anchor: ShapeAnchor{FromCol: 4, FromRow: 3, ToCol: 8, ToRow: 4}}
	three, err := AddShape(two, line)
	if err != nil {
		t.Fatal(err)
	}
	shapes, _ := ReadShapes(three)
	drawingPart := shapes[0].Identity.DrawingPart
	drawing := readEntry(t, three, drawingPart)
	bound := strings.Replace(drawing, `<xdr:cNvCxnSpPr/>`, `<xdr:cNvCxnSpPr><a:stCxn id="1" idx="0"/><a:endCxn id="2" idx="0"/></xdr:cNvCxnSpPr>`, 1)
	workbook := mustApply(t, three, Patch{Replace: map[string][]byte{drawingPart: []byte(bound)}})
	shapes, err = ReadShapes(workbook)
	if err != nil || len(shapes) != 3 {
		t.Fatalf("bound connector hydrate: %+v, %v", shapes, err)
	}
	if _, err := RemoveShape(workbook, shapes[0].Identity); err == nil || !strings.Contains(err.Error(), "inbound connector") {
		t.Fatalf("bound endpoint delete should fail closed, got %v", err)
	}

	endpointUpdate := rectSpec()
	endpointUpdate.Text = "Still connected"
	updatedEndpoint, err := UpdateShape(workbook, shapes[0].Identity, endpointUpdate)
	if err != nil {
		t.Fatal(err)
	}
	if !strings.Contains(readEntry(t, updatedEndpoint, drawingPart), `<a:stCxn id="1" idx="0"/>`) {
		t.Fatal("incoming connector binding was lost when endpoint was updated")
	}

	line.Stroke = "#FF0000"
	updatedLine, err := UpdateShape(updatedEndpoint, shapes[2].Identity, line)
	if err != nil {
		t.Fatal(err)
	}
	lineAnchor := shapeAnchorBytes(t, archiveContents(t, updatedLine)[drawingPart], shapes[2].Identity.ObjectID)
	for _, want := range [][]byte{[]byte(`<a:stCxn id="1" idx="0"/>`), []byte(`<a:endCxn id="2" idx="0"/>`), []byte(`val="FF0000"`)} {
		if !bytes.Contains(lineAnchor, want) {
			t.Errorf("updated bound connector missing %s", want)
		}
	}
	changeKind := rectSpec()
	if _, err := UpdateShape(updatedLine, shapes[2].Identity, changeKind); err == nil || !strings.Contains(err.Error(), "bound connector") {
		t.Fatalf("changing a bound connector kind should fail closed, got %v", err)
	}
}

func TestGroupedShapesRemainOpaqueAndByteIdentical(t *testing.T) {
	one, _ := AddShape(buildZip(t, fixtureWorkbook(false)), rectSpec())
	shape := onlyHydratedShape(t, one)
	drawing := readEntry(t, one, shape.Identity.DrawingPart)
	group := `<xdr:twoCellAnchor><xdr:from><xdr:col>10</xdr:col><xdr:row>10</xdr:row></xdr:from><xdr:to><xdr:col>14</xdr:col><xdr:row>14</xdr:row></xdr:to><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="9" name="Group 1"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr/><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="10" name="Child"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:sp></xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor>`
	withGroup, err := appendAnchorToDrawing(drawing, group)
	if err != nil {
		t.Fatal(err)
	}
	workbook := mustApply(t, one, Patch{Replace: map[string][]byte{shape.Identity.DrawingPart: []byte(withGroup)}})
	if got, err := ReadShapes(workbook); err != nil || len(got) != 1 {
		t.Fatalf("group children must not masquerade as top-level editable shapes: %+v, %v", got, err)
	}
	groupBefore := []byte(group)
	withoutShape, err := RemoveShape(workbook, shape.Identity)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Contains(archiveContents(t, withoutShape)[shape.Identity.DrawingPart], groupBefore) {
		t.Fatal("opaque grouped shape bytes changed during sibling deletion")
	}
	fresh := rectSpec()
	fresh.Anchor = ShapeAnchor{FromCol: 15, FromRow: 1, ToCol: 18, ToRow: 4}
	withFresh, err := AddShape(withoutShape, fresh)
	if err != nil {
		t.Fatal(err)
	}
	created := onlyHydratedShape(t, withFresh)
	if created.Identity.ObjectID != 11 {
		t.Fatalf("insert did not allocate after opaque group ids: %+v", created.Identity)
	}
}

func onlyHydratedShape(t *testing.T, data []byte) ShapeInfo {
	t.Helper()
	shapes, err := ReadShapes(data)
	if err != nil || len(shapes) != 1 {
		t.Fatalf("expected one shape, got %+v, %v", shapes, err)
	}
	return shapes[0]
}

func shapeAnchorBytes(t *testing.T, drawing []byte, objectID uint32) []byte {
	t.Helper()
	index, err := indexShapeDrawing(drawing)
	if err != nil {
		t.Fatal(err)
	}
	for _, anchor := range index.anchors {
		if anchor.objectID == objectID {
			return bytes.Clone(drawing[anchor.span.start:anchor.span.end])
		}
	}
	t.Fatalf("missing shape anchor %d", objectID)
	return nil
}
