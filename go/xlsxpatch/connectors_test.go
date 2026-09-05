package xlsxpatch

import (
	"bytes"
	"encoding/base64"
	"encoding/json"
	"encoding/xml"
	"strings"
	"testing"
)

func connectorFixture(t *testing.T) []byte {
	t.Helper()
	return buildZip(t, map[string]string{
		"[Content_Types].xml":        `<Types xmlns="` + nativeContentTypesNamespace + `"><Default Extension="rels" ContentType="` + nativeRelationshipsType + `"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/xl/workbook.xml" ContentType="` + nativeWorkbookContentType + `"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="` + nativeWorksheetContentType + `"/></Types>`,
		"_rels/.rels":                `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rIdOffice" Type="` + relTypeOfficeDocumentTransitional + `" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml":            `<workbook xmlns="` + spreadsheetMLTransitional + `" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="vendor" Type="https://vendor.example/opaque" Target="../vendor/data.bin"/></Relationships>`,
		"xl/worksheets/sheet1.xml":   `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>preserve</t></is></c></row></sheetData></worksheet>`,
		"vendor/data.bin":            "opaque-vendor-bytes",
	})
}

func connectorDefinitions() []ConnectorDefinition {
	return []ConnectorDefinition{
		{
			ID: "weather", Name: "Weather", Source: ConnectorSource{Kind: "http", URL: "/gateway/weather", Format: "json", Path: "data.items"},
			Target: ConnectorTarget{SheetID: "sheet-1", StartRow: 2, StartColumn: 1}, Refresh: "interval",
			Schedule: &ConnectorSchedule{IntervalMS: 60_000}, Cache: &ConnectorCachePolicy{Mode: "memory", TTLMS: 30_000},
			Schema: &ConnectorSchema{Columns: []ConnectorColumnSchema{{Index: 0, Type: "string"}, {Index: 1, Type: "number", Nullable: true}}},
		},
		{
			ID: "inventory", Name: "Inventory", Source: ConnectorSource{Kind: "http", URL: "https://gateway.example/connectors/inventory", Format: "csv"},
			Target: ConnectorTarget{SheetID: "sheet-1", StartRow: 20, StartColumn: 0}, Refresh: "manual",
			Cache: &ConnectorCachePolicy{Mode: "none", TTLMS: 0},
		},
	}
}

func TestConnectorDefinitions_AddReadAndPreserveUnknownPackageContent(t *testing.T) {
	original := connectorFixture(t)
	before := archiveContents(t, original)
	if got, err := ReadConnectorDefinitions(original); err != nil || len(got) != 0 {
		t.Fatalf("empty read = %+v, %v", got, err)
	}

	out, err := SetConnectorDefinitions(original, connectorDefinitions())
	if err != nil {
		t.Fatal(err)
	}
	got, err := ReadConnectorDefinitions(out)
	if err != nil {
		t.Fatal(err)
	}
	if encoded, wanted := mustJSON(t, got), mustJSON(t, connectorDefinitions()); encoded != wanted {
		t.Fatalf("round trip = %s, want %s", encoded, wanted)
	}
	after := archiveContents(t, out)
	for _, untouched := range []string{"_rels/.rels", "xl/workbook.xml", "xl/worksheets/sheet1.xml", "vendor/data.bin"} {
		if !bytes.Equal(after[untouched], before[untouched]) {
			t.Errorf("untouched part %q changed", untouched)
		}
	}
	if !bytes.Contains(after["xl/_rels/workbook.xml.rels"], before["xl/_rels/workbook.xml.rels"][:len(before["xl/_rels/workbook.xml.rels"])-len("</Relationships>")]) {
		t.Error("workbook relationships were not extended surgically")
	}
	if !bytes.Contains(after["[Content_Types].xml"], before["[Content_Types].xml"][:len(before["[Content_Types].xml"])-len("</Types>")]) {
		t.Error("content types were not extended surgically")
	}
	if !bytes.Contains(after["customXml/injofficeConnectors.xml"], []byte(ConnectorExtensionNamespaceV1)) {
		t.Error("connector custom XML part is missing its namespace")
	}
}

func TestConnectorDefinitions_UpdateTouchesOnlyOwnedPart(t *testing.T) {
	withConnectors, err := SetConnectorDefinitions(connectorFixture(t), connectorDefinitions())
	if err != nil {
		t.Fatal(err)
	}
	before := archiveContents(t, withConnectors)
	updated := connectorDefinitions()
	updated[0].Name = "Forecast"
	updated[0].Refresh = "onOpen"
	updated[0].Schedule = nil
	out, err := SetConnectorDefinitions(withConnectors, updated)
	if err != nil {
		t.Fatal(err)
	}
	after := archiveContents(t, out)
	for name, value := range before {
		if name == ConnectorExtensionPartNameV1 {
			continue
		}
		if !bytes.Equal(after[name], value) {
			t.Errorf("update changed unrelated part %q", name)
		}
	}
	got, err := ReadConnectorDefinitions(out)
	if err != nil || len(got) != 2 || got[0].Name != "Forecast" || got[0].Schedule != nil {
		t.Fatalf("updated read = %+v, %v", got, err)
	}
}

func TestConnectorDefinitions_RemoveKeepsUnknownRelationshipsAndParts(t *testing.T) {
	withConnectors, err := SetConnectorDefinitions(connectorFixture(t), connectorDefinitions())
	if err != nil {
		t.Fatal(err)
	}
	out, err := SetConnectorDefinitions(withConnectors, nil)
	if err != nil {
		t.Fatal(err)
	}
	got, err := ReadConnectorDefinitions(out)
	if err != nil || len(got) != 0 {
		t.Fatalf("removed read = %+v, %v", got, err)
	}
	parts := archiveContents(t, out)
	if _, exists := parts[ConnectorExtensionPartNameV1]; exists {
		t.Fatal("connector part survived removal")
	}
	if !bytes.Contains(parts["xl/_rels/workbook.xml.rels"], []byte(`Id="vendor"`)) || !bytes.Equal(parts["vendor/data.bin"], []byte("opaque-vendor-bytes")) {
		t.Fatal("unknown relationship or target was not preserved")
	}
	if bytes.Contains(parts["xl/_rels/workbook.xml.rels"], []byte(ConnectorExtensionRelationshipTypeV1)) || bytes.Contains(parts["[Content_Types].xml"], []byte(ConnectorExtensionContentTypeV1)) {
		t.Fatal("connector relationship or content type survived removal")
	}
}

func TestConnectorDefinitions_FailClosedOnSecretsAndInvalidModels(t *testing.T) {
	for name, mutate := range map[string]func(*[]ConnectorDefinition){
		"query secret":   func(items *[]ConnectorDefinition) { (*items)[0].Source.URL = "/gateway/weather?token=secret" },
		"userinfo":       func(items *[]ConnectorDefinition) { (*items)[0].Source.URL = "https://user:pass@example.test/data" },
		"duplicate id":   func(items *[]ConnectorDefinition) { (*items)[1].ID = (*items)[0].ID },
		"bad coordinate": func(items *[]ConnectorDefinition) { (*items)[0].Target.StartColumn = 16_384 },
		"bad schedule":   func(items *[]ConnectorDefinition) { (*items)[0].Schedule = nil },
	} {
		t.Run(name, func(t *testing.T) {
			items := connectorDefinitions()
			mutate(&items)
			if _, err := SetConnectorDefinitions(connectorFixture(t), items); err == nil {
				t.Fatal("expected refusal")
			}
		})
	}
}

func TestConnectorDefinitions_RefuseUnknownOrAmbiguousOwnedState(t *testing.T) {
	valid, err := SetConnectorDefinitions(connectorFixture(t), connectorDefinitions())
	if err != nil {
		t.Fatal(err)
	}
	part := readEntry(t, valid, ConnectorExtensionPartNameV1)
	for name, replacement := range map[string]string{
		"unknown version":  strings.Replace(part, `version="1"`, `version="2"`, 1),
		"unknown encoding": strings.Replace(part, `encoding="base64-json"`, `encoding="xml"`, 1),
		"nested content":   strings.Replace(part, `</connectors>`, `<future/></connectors>`, 1),
		"null connectors":  connectorPartWithJSON(`{"version":1,"connectors":null}`),
		"unknown JSON":     connectorPartWithJSON(`{"version":1,"connectors":[],"future":true}`),
	} {
		t.Run(name, func(t *testing.T) {
			corrupt, err := Apply(valid, Patch{Replace: map[string][]byte{ConnectorExtensionPartNameV1: []byte(replacement)}})
			if err != nil {
				t.Fatal(err)
			}
			if _, err := ReadConnectorDefinitions(corrupt); err == nil {
				t.Fatal("expected read refusal")
			}
			if _, err := SetConnectorDefinitions(corrupt, connectorDefinitions()); err == nil {
				t.Fatal("expected overwrite refusal")
			}
		})
	}

	rels := readEntry(t, valid, "xl/_rels/workbook.xml.rels")
	duplicate, err := appendRelationship(rels, "rIdDuplicate", ConnectorExtensionRelationshipTypeV1, "../customXml/injofficeConnectors.xml")
	if err != nil {
		t.Fatal(err)
	}
	ambiguous, err := Apply(valid, Patch{Replace: map[string][]byte{"xl/_rels/workbook.xml.rels": []byte(duplicate)}})
	if err != nil {
		t.Fatal(err)
	}
	if _, err := ReadConnectorDefinitions(ambiguous); err == nil || !strings.Contains(err.Error(), "multiple") {
		t.Fatalf("ambiguous relationship error = %v", err)
	}
}

func connectorPartWithJSON(payload string) string {
	return xml.Header + `<connectors xmlns="` + ConnectorExtensionNamespaceV1 + `" version="1" encoding="base64-json">` +
		base64.StdEncoding.EncodeToString([]byte(payload)) + `</connectors>`
}

func TestConnectorDefinitions_RoutesRelocatedWorkbook(t *testing.T) {
	parts := archiveContents(t, connectorFixture(t))
	parts["Custom/Office/Book.xml"] = bytes.ReplaceAll(parts["xl/workbook.xml"], []byte("xl/"), []byte("Custom/Office/"))
	parts["Custom/Office/_rels/Book.xml.rels"] = bytes.ReplaceAll(parts["xl/_rels/workbook.xml.rels"], []byte(`Target="worksheets/sheet1.xml"`), []byte(`Target="../../xl/worksheets/sheet1.xml"`))
	delete(parts, "xl/workbook.xml")
	delete(parts, "xl/_rels/workbook.xml.rels")
	parts["_rels/.rels"] = bytes.ReplaceAll(parts["_rels/.rels"], []byte(`Target="xl/workbook.xml"`), []byte(`Target="Custom/Office/Book.xml"`))
	parts["[Content_Types].xml"] = bytes.ReplaceAll(parts["[Content_Types].xml"], []byte(`/xl/workbook.xml`), []byte(`/Custom/Office/Book.xml`))
	stringsMap := map[string]string{}
	for name, value := range parts {
		stringsMap[name] = string(value)
	}
	original := buildZip(t, stringsMap)
	out, err := SetConnectorDefinitions(original, connectorDefinitions())
	if err != nil {
		t.Fatal(err)
	}
	rels := readEntry(t, out, "Custom/Office/_rels/Book.xml.rels")
	if !strings.Contains(rels, `Target="../../customXml/injofficeConnectors.xml"`) {
		t.Fatalf("relocated target = %s", rels)
	}
	if got, err := ReadConnectorDefinitions(out); err != nil || len(got) != 2 {
		t.Fatalf("relocated read = %+v, %v", got, err)
	}
}

func mustJSON(t *testing.T, value any) string {
	t.Helper()
	encoded, err := json.Marshal(value)
	if err != nil {
		t.Fatal(err)
	}
	return string(encoded)
}
