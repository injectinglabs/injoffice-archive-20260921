package xlsxpatch

import (
	"strings"
	"testing"
)

func TestChartReferenceCachesLiteralRanges(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData><row r="1"><c r="B1" t="inlineStr"><is><t>Actual &amp; planned</t></is></c></row><row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2"><v>1200.5</v></c></row><row r="4"><c r="A4" t="inlineStr"><is><r><t>Food</t></r><r><t> &amp; drink</t></r></is></c><c r="B4"><v>-25</v></c></row></sheetData></worksheet>`
	entries["xl/sharedStrings.xml"] = `<sst><si><t>Rent</t></si></sst>`
	entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], "</Relationships>", `<Relationship Id="rShared" Type="`+relTypeSharedStringsTransitional+`" Target="sharedStrings.xml"/></Relationships>`, 1)
	spec := writeSpec("column")
	spec.Series = spec.Series[:1]
	spec.Series[0].Name = "Wrong caller label"
	out, err := AddChart(buildZip(t, entries), spec)
	if err != nil {
		t.Fatal(err)
	}
	chart := readEntry(t, out, "xl/charts/chart1.xml")
	for _, want := range []string{`<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>Actual &amp; planned</c:v>`, `<c:ptCount val="4"/>`, `<c:pt idx="0"><c:v>Rent</c:v>`, `<c:pt idx="2"><c:v>Food &amp; drink</c:v>`, `<c:numCache><c:formatCode>General</c:formatCode>`, `<c:pt idx="0"><c:v>1200.5</c:v>`, `<c:pt idx="2"><c:v>-25</c:v>`} {
		if !strings.Contains(chart, want) {
			t.Errorf("missing %s in %s", want, chart)
		}
	}
	if strings.Contains(chart, "Wrong caller label") {
		t.Fatal("invented cache from caller label")
	}
	if spec.Series[0].nameCache != "" {
		t.Fatal("mutated caller series")
	}
	charts, err := ReadCharts(out)
	if err != nil || len(charts) != 1 || charts[0].Series[0].Name != "Actual & planned" {
		t.Fatalf("readback: %+v %v", charts, err)
	}
	ids, err := ReadChartIdentities(out)
	if err != nil || len(ids) != 1 {
		t.Fatalf("chart identity: %+v %v", ids, err)
	}
	updated, err := UpdateChart(out, ids[0], spec)
	if err != nil {
		t.Fatal(err)
	}
	if got := readEntry(t, updated, "xl/charts/chart1.xml"); got != chart {
		t.Fatal("update lost source-derived caches")
	}
}

func TestChartCachesOmitUnsupportedOrPotentiallyStaleValues(t *testing.T) {
	for _, test := range []struct {
		name, cell, ref string
		numeric         bool
	}{
		{"formula with stale cache", `<c r="A1"><f>1+1</f><v>99</v></c>`, "Data!A1", true},
		{"formula without cache", `<c r="A1"><f>1+1</f></c>`, "Data!A1", true},
		{"array dependent cache", `<c r="A1"><f t="array" ref="A1:B1">SEQUENCE(1,2)</f><v>1</v></c><c r="B1"><v>99</v></c>`, "Data!B1", true},
		{"error", `<c r="A1" t="e"><v>#REF!</v></c>`, "Data!A1", true},
		{"numeric text", `<c r="A1" t="inlineStr"><is><t>123</t></is></c>`, "Data!A1", true},
		{"nonfinite", `<c r="A1"><v>NaN</v></c>`, "Data!A1", true},
		{"formatted category", `<c r="A1" s="5"><v>45000</v></c>`, "Data!A1", false},
		{"two dimensional", ``, "Data!A1:B2", true},
		{"external", ``, "[other.xlsx]Data!A1", true},
		{"budget", ``, "Data!A1:A10001", true},
	} {
		t.Run(test.name, func(t *testing.T) {
			entries := fixtureWorkbook(false)
			entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData><row r="1">` + test.cell + `</row></sheetData></worksheet>`
			read := func(name string) (string, bool) { value, ok := entries[name]; return value, ok }
			if got := chartReferenceCache(read, test.ref, test.numeric, false); got != "" {
				t.Fatalf("unsafe cache: %s", got)
			}
		})
	}
}

func TestChartCacheReferenceFormsAndNameRefWithoutInventedLabel(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/workbook.xml"] = strings.ReplaceAll(entries["xl/workbook.xml"], `name="Data"`, `name="Team's data"`)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData></worksheet>`
	read := func(name string) (string, bool) { value, ok := entries[name]; return value, ok }
	cache := chartReferenceCache(read, "'Team''s data'!$A$1:$B$1", true, false)
	if !strings.Contains(cache, `<c:pt idx="1"><c:v>2</c:v>`) {
		t.Fatal(cache)
	}
	if cache := chartReferenceCache(read, "'Team''s data'!C1", false, true); cache != "" {
		t.Fatal(cache)
	}
	if strings.Contains(serTxXML(WriteSeries{Name: "guess", NameRef: "Data!A1"}, 0), "strCache") {
		t.Fatal("guessed series cache")
	}
}

func TestChartCachesResolveActualSharedStringRelationship(t *testing.T) {
	entries := fixtureWorkbook(false)
	entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row></sheetData></worksheet>`
	entries["xl/sharedStrings.xml"] = `<sst><si><t>Decoy</t></si></sst>`
	entries["xl/strings/actual.xml"] = `<sst><si><t>Actual label</t></si></sst>`
	base := entries["xl/_rels/workbook.xml.rels"]
	read := func(name string) (string, bool) { value, ok := entries[name]; return value, ok }
	if got := chartReferenceCache(read, "Data!A1", false, false); got != "" {
		t.Fatal("used unrelated decoy", got)
	}
	for _, relationship := range []string{
		`<Relationship Id="rShared" Type="` + relTypeSharedStringsTransitional + `" Target="strings/actual.xml"/>`,
		`<Relationship Id="rShared" Type="` + relTypeSharedStringsStrict + `" Target="/xl/strings/actual.xml"/>`,
	} {
		entries["xl/_rels/workbook.xml.rels"] = strings.Replace(base, "</Relationships>", relationship+"</Relationships>", 1)
		if got := chartReferenceCache(read, "Data!A1", false, false); !strings.Contains(got, "Actual label") || strings.Contains(got, "Decoy") {
			t.Fatal(got)
		}
	}
	for _, relationships := range []string{
		`<Relationship Id="rShared" Type="` + relTypeSharedStringsTransitional + `" Target="sharedStrings.xml" TargetMode="External"/>`,
		`<Relationship Id="rShared" Type="` + relTypeSharedStringsTransitional + `" Target="sharedStrings.xml"/><Relationship Id="rOther" Type="` + relTypeSharedStringsTransitional + `" Target="strings/actual.xml"/>`,
	} {
		entries["xl/_rels/workbook.xml.rels"] = strings.Replace(base, "</Relationships>", relationships+"</Relationships>", 1)
		if got := chartReferenceCache(read, "Data!A1", false, false); got != "" {
			t.Fatal("ambiguous/external relation", got)
		}
	}
}

func TestChartNumericCategoryOmitsCustomStyleZero(t *testing.T) {
	for _, style := range []string{"", ` s="0"`} {
		entries := fixtureWorkbook(false)
		entries["xl/worksheets/sheet1.xml"] = `<worksheet><sheetData><row r="1"><c r="A1"` + style + `><v>45000</v></c></row></sheetData></worksheet>`
		entries["xl/styles.xml"] = `<styleSheet><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><cellXfs count="1"><xf numFmtId="164" applyNumberFormat="1"/></cellXfs></styleSheet>`
		entries["xl/_rels/workbook.xml.rels"] = strings.Replace(entries["xl/_rels/workbook.xml.rels"], "</Relationships>", `<Relationship Id="rStyle" Type="`+relTypeStylesTransitional+`" Target="styles.xml"/></Relationships>`, 1)
		read := func(name string) (string, bool) { value, ok := entries[name]; return value, ok }
		if got := chartReferenceCache(read, "Data!A1", false, false); got != "" {
			t.Fatal("cached serial number as date label", got)
		}
		if got := chartReferenceCache(read, "Data!A1", true, false); !strings.Contains(got, "45000") {
			t.Fatal("numeric series value unexpectedly lost", got)
		}
	}
}
