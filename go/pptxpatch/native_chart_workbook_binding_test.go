package pptxpatch

import (
	"strings"
	"testing"
)

func nativeWorkbookBindingFixture(t *testing.T, strict bool, relType, target, mode, contentType, bytes string) *nativeExtractor {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
	if relType == "" {
		relType = d.rels + "/package"
	}
	if target == "" {
		target = "../embeddings/Data.xlsx"
	}
	if contentType == "" {
		contentType = nativeChartWorkbookContentType
	}
	modeAttr := ""
	if mode != "" {
		modeAttr = ` TargetMode="` + mode + `"`
	}
	payload := nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, extraParts: []nativeExtractZipPart{
		{name: "relocated/charts/source.xml", data: `<c:chartSpace xmlns:c="` + d.chart + `"/>`},
		{name: "relocated/embeddings/Data.xlsx", data: bytes},
		{name: "relocated/charts/_rels/source.xml.rels", data: `<Relationships xmlns="` + nsPackageRels + `"><Relationship Id="workbook" Type="` + relType + `" Target="` + target + `"` + modeAttr + `/></Relationships>`},
	}, mutate: func(parts map[string]string) {
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/relocated/charts/source.xml" ContentType="`+contentTypeChart+`"/><Override PartName="/relocated/embeddings/Data.xlsx" ContentType="`+contentType+`"/></Types>`, 1)
	}})
	pkg, e := openNativeExtractPackage(payload)
	if e != nil {
		t.Fatal(e)
	}
	return &nativeExtractor{pkg: pkg, relationshipCache: map[string][]nativeExtractRelationship{}}
}
func TestNativeChartEmbeddedWorkbookBinding(t *testing.T) {
	for _, strict := range []bool{false, true} {
		d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
		source := `<c:externalData xmlns:c="` + d.chart + `" xmlns:r="` + d.rels + `" r:id="workbook"><c:autoUpdate val="0"/></c:externalData>`
		node, e := parseNativeXML([]byte(source), "chart.xml")
		if e != nil {
			t.Fatal(e)
		}
		extractor := nativeWorkbookBindingFixture(t, strict, "", "", "", "", "Workbook bytes are validated by the injected XLSX engine")
		before := string(extractor.pkg.parts["relocated/embeddings/Data.xlsx"])
		binding, ok, e := extractor.extractNativeChartWorkbookBinding(node, "relocated/charts/source.xml", d)
		if e != nil || !ok || binding.Part != "relocated/embeddings/Data.xlsx" || binding.SHA256 != nativeSHA256([]byte(before)) || binding.ByteLength != int64(len(before)) || binding.AutoUpdate == nil || *binding.AutoUpdate {
			t.Fatalf("source binding missing %+v %v", binding, e)
		}
		if string(extractor.pkg.parts[binding.Part]) != before {
			t.Fatal("embedded source mutated")
		}
		for _, pair := range [][2]string{{`r:id="workbook"`, `r:id="missing"`}, {`r:id="workbook"`, `id="workbook"`}, {`val="0"`, `val="yes"`}, {`<c:autoUpdate val="0"/>`, `<c:extLst/>`}, {`<c:autoUpdate val="0"/>`, `<c:autoUpdate/>`}} {
			bad, e := parseNativeXML([]byte(strings.Replace(source, pair[0], pair[1], 1)), "bad.xml")
			if e != nil {
				continue
			}
			if _, ok, e := extractor.extractNativeChartWorkbookBinding(bad, "relocated/charts/source.xml", d); ok || e != nil {
				t.Fatalf("unqualified binding accepted %v %v", pair, e)
			}
		}
		for _, item := range []struct{ kind, target, mode, content, bytes string }{
			{target: "https://example.invalid/book.xlsx", mode: "External", bytes: "x"},
			{kind: d.rels + "/oleObject", bytes: "x"},
			{content: "application/vnd.ms-excel.sheet.macroEnabled.12", bytes: "x"},
			{bytes: ""},
			{bytes: strings.Repeat("x", nativeChartWorkbookMaxBytes+1)},
		} {
			other := nativeWorkbookBindingFixture(t, strict, item.kind, item.target, item.mode, item.content, item.bytes)
			if _, ok, e := other.extractNativeChartWorkbookBinding(node, "relocated/charts/source.xml", d); ok || e != nil {
				t.Fatalf("unsupported workbook binding accepted %v", e)
			}
		}
	}
}
