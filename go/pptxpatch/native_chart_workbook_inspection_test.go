package pptxpatch

import (
	"encoding/base64"
	"fmt"
	"sort"
	"strings"
	"testing"
)

func nativeWorkbookInspectionFixture(t *testing.T, strict bool, family string, mutation func(map[string]string)) []byte {
	d, _ := nativeDialectForPresentation(xmlNamePresentation(strict))
	return nativeExtractFixture(t, nativeExtractFixtureOptions{strict: strict, extraParts: []nativeExtractZipPart{
		{name: "relocated/charts/source.xml", data: nativeWorkbookChartXML(strict, family)},
		{name: "relocated/embeddings/Data.xlsx", data: "Exact embedded bytes; XLSX validation belongs to the injected engine"},
		{name: "relocated/charts/_rels/source.xml.rels", data: `<Relationships xmlns="` + d.packageRels + `"><Relationship Id="workbook" Type="` + d.rels + `/package" Target="../embeddings/Data.xlsx"/></Relationships>`},
	}, mutate: func(parts map[string]string) {
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/relocated/charts/source.xml" ContentType="`+contentTypeChart+`"/><Override PartName="/relocated/embeddings/Data.xlsx" ContentType="`+nativeChartWorkbookContentType+`"/></Types>`, 1)
		parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`, `<Relationship Id="rIdChart" Type="`+d.rels+`/chart" Target="../charts/source.xml"/></Relationships>`, 1)
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, nativeChartGraphicFrameXML(strict, 3, "Chart", "")+`</p:spTree>`, 1)
		if mutation != nil {
			mutation(parts)
		}
	}})
}
func TestNativeChartWorkbookInspectionSourceClosure(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, family := range []string{"column", "bar", "line", "scatter"} {
			input := nativeWorkbookInspectionFixture(t, strict, family, nil)
			before := append([]byte(nil), input...)
			result, e := InspectNativePPTXChartWorkbooks(input)
			if e != nil {
				t.Fatal(e)
			}
			if len(result.Charts) != 1 || len(result.Workbooks) != 1 || len(result.Omissions) != 0 || result.PackageSHA256 != nativeSHA256(input) || result.SourceRevision != "rev-"+result.PackageSHA256 {
				t.Fatalf("invalid source closure: %#v", result)
			}
			chart := result.Charts[0]
			resource := result.Workbooks[0]
			decoded, e := base64.StdEncoding.DecodeString(resource.BytesBase64)
			if e != nil {
				t.Fatal(e)
			}
			if nativeSHA256(decoded) != resource.SHA256 || int64(len(decoded)) != resource.ByteLength || resource.Part != chart.Workbook.Part || resource.SHA256 != chart.Workbook.SHA256 || chart.Workbook.RelationshipID != "workbook" || chart.ChartRelationshipID != "rIdChart" || chart.ObjectID != "cNvPr-3" || chart.Source.PlotVisibleOnly || len(chart.Source.Series) != 1 {
				t.Fatal("binding lost source identity")
			}
			pkg, e := openNativeExtractPackage(input)
			if e != nil {
				t.Fatal(e)
			}
			if chart.ChartSHA256 != nativeSHA256(pkg.parts[chart.ChartPart]) || chart.SlideSHA256 != nativeSHA256(pkg.parts[chart.SlidePart]) {
				t.Fatal("source hash mismatch")
			}
			deck, e := ExtractNativePPTX(input, nativeTestExtractOptions())
			if e != nil {
				t.Fatal(e)
			}
			var found bool
			for _, element := range deck.Slides[0].Elements {
				if element.ID == chart.ElementID {
					found = true
					if element.Source.FingerprintSHA256 != chart.FrameSHA256 || element.Compatibility.Status != NativeCompatibilityStatusPreserveOnly {
						t.Fatal("inspection changed frame authority")
					}
				}
			}
			if !found {
				t.Fatal("frame not source-bound")
			}
			if string(input) != string(before) {
				t.Fatal("inspection mutated package")
			}
		}
	}
}
func TestNativeChartWorkbookInspectionOmissionAndResourceDedup(t *testing.T) {
	input := nativeWorkbookInspectionFixture(t, false, "column", func(parts map[string]string) {
		parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, nativeChartGraphicFrameXML(false, 4, "Second", "")+`</p:spTree>`, 1)
	})
	result, e := InspectNativePPTXChartWorkbooks(input)
	if e != nil || len(result.Charts) != 2 || len(result.Workbooks) != 1 {
		t.Fatalf("dedup failed: %#v %v", result, e)
	}
	input = nativeWorkbookInspectionFixture(t, false, "column", func(parts map[string]string) {
		parts["relocated/charts/source.xml"] = strings.Replace(parts["relocated/charts/source.xml"], `plotVisOnly val="0"`, `plotVisOnly val="1"`, 1)
	})
	result, e = InspectNativePPTXChartWorkbooks(input)
	if e != nil || len(result.Charts) != 0 || len(result.Workbooks) != 0 || len(result.Omissions) != 1 {
		t.Fatalf("refused source not disclosed: %#v %v", result, e)
	}
}

func TestNativeChartWorkbookInspectionRecordBudget(t *testing.T) {
	for _, count := range []int{64, 65} {
		input := nativeWorkbookInspectionFixture(t, false, "column", func(parts map[string]string) {
			var frames strings.Builder
			for i := 1; i < count; i++ {
				frames.WriteString(nativeChartGraphicFrameXML(false, i+3, "Chart", ""))
			}
			parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, frames.String()+`</p:spTree>`, 1)
		})
		result, err := InspectNativePPTXChartWorkbooks(input)
		if count == 64 {
			if err != nil || len(result.Charts) != 64 {
				t.Fatalf("exact record limit: %v %d", err, len(result.Charts))
			}
		} else if err == nil || !strings.Contains(err.Error(), "record budget") {
			t.Fatalf("record overflow: %v", err)
		}
	}
}

func TestNativeChartWorkbookInspectionResourceBudgets(t *testing.T) {
	for _, test := range []struct {
		count     int
		large     bool
		wantError string
	}{{8, false, ""}, {9, false, "resource count"}, {2, true, ""}, {3, true, "aggregate workbook"}} {
		var completeParts map[string]string
		_ = nativeWorkbookInspectionFixture(t, false, "column", func(parts map[string]string) {
			completeParts = parts
			d, _ := nativeDialectForPresentation(xmlNamePresentation(false))
			if test.large {
				parts["relocated/embeddings/Data.xlsx"] = strings.Repeat("x", 8*1024*1024)
			}
			for i := 1; i < test.count; i++ {
				chart := fmt.Sprintf("chart%d.xml", i)
				book := fmt.Sprintf("Data%d.xlsx", i)
				rel := fmt.Sprintf("rChart%d", i)
				parts["relocated/charts/"+chart] = parts["relocated/charts/source.xml"]
				parts["relocated/charts/_rels/"+chart+".rels"] = strings.Replace(parts["relocated/charts/_rels/source.xml.rels"], "Data.xlsx", book, 1)
				size := 1
				if test.large && i == 1 {
					size = 8 * 1024 * 1024
				}
				parts["relocated/embeddings/"+book] = strings.Repeat("y", size)
				parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/relocated/charts/`+chart+`" ContentType="`+contentTypeChart+`"/><Override PartName="/relocated/embeddings/`+book+`" ContentType="`+nativeChartWorkbookContentType+`"/></Types>`, 1)
				parts["relocated/slides/_rels/slide-a.xml.rels"] = strings.Replace(parts["relocated/slides/_rels/slide-a.xml.rels"], `</Relationships>`, `<Relationship Id="`+rel+`" Type="`+d.rels+`/chart" Target="../charts/`+chart+`"/></Relationships>`, 1)
				frame := strings.Replace(nativeChartGraphicFrameXML(false, i+3, "Chart", ""), `r:id="rIdChart"`, `r:id="`+rel+`"`, 1)
				parts["relocated/slides/slide-a.xml"] = strings.Replace(parts["relocated/slides/slide-a.xml"], `</p:spTree>`, frame+`</p:spTree>`, 1)
			}
		})
		names := make([]string, 0, len(completeParts))
		for name := range completeParts {
			names = append(names, name)
		}
		sort.Strings(names)
		zipParts := make([]nativeExtractZipPart, 0, len(names))
		for _, name := range names {
			zipParts = append(zipParts, nativeExtractZipPart{name: name, data: completeParts[name]})
		}
		input := writeNativeExtractZip(t, zipParts)
		result, err := InspectNativePPTXChartWorkbooks(input)
		if test.wantError == "" {
			if err != nil || len(result.Workbooks) != test.count {
				t.Fatalf("exact resource limit: %v %d", err, len(result.Workbooks))
			}
		} else if err == nil || !strings.Contains(err.Error(), test.wantError) {
			t.Fatalf("resource overflow: %v", err)
		}
	}
}
