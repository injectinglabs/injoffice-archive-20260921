package xlsxpatch

import (
	"errors"
	"strings"
	"testing"
)

func preservationFixture(t *testing.T) []byte {
	t.Helper()
	entries := fixtureWorkbook(false)
	entries["xl/charts/chart1.xml"] = barChartXML
	entries["xl/pivotTables/pivotTable1.xml"] = "<pivotTableDefinition/>"
	entries["xl/pivotCache/pivotCacheDefinition1.xml"] = "<pivotCacheDefinition/>"
	entries["xl/drawings/drawing1.xml"] = "<xdr:wsDr><xdr:twoCellAnchor><xdr:graphicFrame uri=\"unknown-object\"/></xdr:twoCellAnchor></xdr:wsDr>"
	entries["xl/drawings/_rels/drawing1.xml.rels"] = "<Relationships><Relationship Id=\"rId1\" Type=\"image\" Target=\"../media/image1.png\"/><Relationship Id=\"rId2\" Type=\"chart\" Target=\"../charts/chart1.xml\"/></Relationships>"
	entries["xl/media/image1.png"] = "png-bytes"
	entries["xl/embeddings/oleObject1.bin"] = "embedded-bytes"
	entries["xl/worksheets/sheet1.xml"] = "<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData/><drawing r:id=\"rId1\"/><pivotTableParts count=\"1\"><pivotTablePart r:id=\"rId2\"/></pivotTableParts></worksheet>"
	entries["xl/worksheets/_rels/sheet1.xml.rels"] = "<Relationships><Relationship Id=\"rId1\" Type=\"drawing\" Target=\"../drawings/drawing1.xml\"/><Relationship Id=\"rId2\" Type=\"pivotTable\" Target=\"../pivotTables/pivotTable1.xml\"/></Relationships>"
	return buildZip(t, entries)
}

func TestPreservationInventoryCoversOfficeObjectsAndRelationships(t *testing.T) {
	inventory, err := PreservationInventory(preservationFixture(t))
	if err != nil {
		t.Fatal(err)
	}
	joined := ""
	for _, item := range inventory {
		joined += item.Name + "\n"
	}
	for _, want := range []string{"xl/charts/chart1.xml", "xl/pivotTables/pivotTable1.xml", "xl/drawings/drawing1.xml", "xl/media/image1.png", "xl/embeddings/oleObject1.bin", "@rel/xl/worksheets/sheet1.xml", "@ref/xl/worksheets/sheet1.xml"} {
		if !strings.Contains(joined, want) {
			t.Errorf("inventory missing %q:\n%s", want, joined)
		}
	}
}

func TestPreservationInventoryRefusesRelationshipTraversalAbovePackageRoot(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheetData/><drawing r:id="rId1"/></worksheet>`
	entries["xl/worksheets/_rels/sheet1.xml.rels"] = `<Relationships><Relationship Id="rId1" Type="drawing" Target="../../../xl/drawings/drawing1.xml"/></Relationships>`
	entries["xl/drawings/drawing1.xml"] = `<xdr:wsDr/>`

	inventory, err := PreservationInventory(buildZip(t, entries))
	if err == nil || !strings.Contains(err.Error(), "traverses above the package root") {
		t.Fatalf("expected relationship traversal refusal, inventory=%v err=%v", inventory, err)
	}
}

func TestPreservationInventoryToleratesExternalHyperlinkRelationships(t *testing.T) {
	for _, target := range []string{"mailto:native@example.test", "file:///tmp/native.xlsx", "relative-external-target"} {
		t.Run(target, func(t *testing.T) {
			entries := fixtureWorkbook(false)
			entries["xl/worksheets/_rels/sheet1.xml.rels"] = `<Relationships><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="` + target + `" TargetMode="External"/></Relationships>`
			original := buildZip(t, entries)
			if _, err := PreservationInventory(original); err != nil {
				t.Fatalf("valid external relationship was rejected: %v", err)
			}
			if err := RequirePreservation(original, original, nil); err != nil {
				t.Fatalf("valid external relationship was not preserved: %v", err)
			}
		})
	}
}

func TestPreservationInventoryRefusesUnknownRelationshipTargetMode(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/_rels/sheet1.xml.rels"] = `<Relationships><Relationship Id="rIdUnknown" Type="hyperlink" Target="relative-target" TargetMode="Sideways"/></Relationships>`
	inventory, err := PreservationInventory(buildZip(t, entries))
	if err == nil || !strings.Contains(err.Error(), `unsupported target mode "Sideways"`) {
		t.Fatalf("expected unknown target-mode refusal, inventory=%v err=%v", inventory, err)
	}
}

func TestRequirePreservationAllowsCellChangesButRejectsObjectLoss(t *testing.T) {
	original := preservationFixture(t)
	cellOnly, err := Apply(original, Patch{Replace: map[string][]byte{
		"xl/worksheets/sheet1.xml": []byte("<worksheet xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheetData><row r=\"1\"><c r=\"A1\"><v>changed</v></c></row></sheetData><drawing r:id=\"rId1\"/><pivotTableParts count=\"1\"><pivotTablePart r:id=\"rId2\"/></pivotTableParts></worksheet>"),
	}})
	if err != nil {
		t.Fatal(err)
	}
	if err := RequirePreservation(original, cellOnly, nil); err != nil {
		t.Fatalf("cell-only edit should pass: %v", err)
	}

	lost, err := Apply(original, Patch{Delete: map[string]bool{"xl/media/image1.png": true}})
	if err != nil {
		t.Fatal(err)
	}
	err = RequirePreservation(original, lost, nil)
	var preservationErr *PreservationError
	if !errors.As(err, &preservationErr) || len(preservationErr.Report.Missing) == 0 {
		t.Fatalf("expected preservation loss report, got %v", err)
	}
}

func TestRequirePreservationDetectsUnknownDrawingChangesAndSupportsExplicitAllowlist(t *testing.T) {
	original := preservationFixture(t)
	changed, err := Apply(original, Patch{Replace: map[string][]byte{
		"xl/drawings/drawing1.xml": []byte("<xdr:wsDr/>"),
	}})
	if err != nil {
		t.Fatal(err)
	}
	if err := RequirePreservation(original, changed, nil); err == nil {
		t.Fatal("unknown drawing loss must be rejected")
	}
	if err := RequirePreservation(original, changed, []string{"xl/drawings/drawing1.xml"}); err != nil {
		t.Fatalf("explicitly allowed drawing edit should pass: %v", err)
	}
}
