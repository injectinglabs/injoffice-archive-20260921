package officecompat_test

import (
	"archive/zip"
	"bytes"
	"errors"
	"sort"
	"testing"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const (
	contentTypes  = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>`
	emptyRootRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`
)

func TestInspectIsStableAcrossZIPOrderAndCompression(t *testing.T) {
	entries := map[string][]byte{
		"[Content_Types].xml": []byte(contentTypes),
		"_rels/.rels":         []byte(emptyRootRels),
		"custom/data.bin":     []byte("opaque payload"),
	}
	first := buildPackage(t, entries, false, false)
	second := buildPackage(t, entries, true, true)

	left, err := officecompat.Inspect(first)
	if err != nil {
		t.Fatal(err)
	}
	right, err := officecompat.Inspect(second)
	if err != nil {
		t.Fatal(err)
	}
	if left.Fingerprint != right.Fingerprint {
		t.Fatalf("content-equivalent packages differ: %s != %s", left.Fingerprint, right.Fingerprint)
	}
	if len(left.Parts) != 3 || left.Parts[0].Name != "[Content_Types].xml" || left.Parts[2].Name != "custom/data.bin" {
		t.Fatalf("unexpected deterministic inventory: %+v", left.Parts)
	}
}

func TestRequireUntouchedPartsDetectsEscapedMutations(t *testing.T) {
	beforeEntries := map[string][]byte{
		"[Content_Types].xml":   []byte(contentTypes),
		"_rels/.rels":           []byte(emptyRootRels),
		"word/document.xml":     []byte("before"),
		"word/styles.xml":       []byte("styles"),
		"word/theme/theme1.xml": []byte("theme"),
	}
	afterEntries := cloneEntries(beforeEntries)
	afterEntries["word/document.xml"] = []byte("after")
	afterEntries["word/styles.xml"] = []byte("changed styles")
	afterEntries["custom/new.xml"] = []byte("new")
	delete(afterEntries, "word/theme/theme1.xml")
	before := buildPackage(t, beforeEntries, false, false)
	after := buildPackage(t, afterEntries, false, false)

	err := officecompat.RequireUntouchedParts(before, after, []string{"word/document.xml"})
	var preservationErr *officecompat.PreservationError
	if !errors.As(err, &preservationErr) {
		t.Fatalf("expected PreservationError, got %v", err)
	}
	if len(preservationErr.Report.Changed) != 1 || preservationErr.Report.Changed[0].Before.Name != "word/styles.xml" {
		t.Fatalf("changed report: %+v", preservationErr.Report.Changed)
	}
	if len(preservationErr.Report.Added) != 1 || preservationErr.Report.Added[0].Name != "custom/new.xml" {
		t.Fatalf("added report: %+v", preservationErr.Report.Added)
	}
	if len(preservationErr.Report.Missing) != 1 || preservationErr.Report.Missing[0].Name != "word/theme/theme1.xml" {
		t.Fatalf("missing report: %+v", preservationErr.Report.Missing)
	}

	if err := officecompat.RequireUntouchedParts(before, after, []string{"word/document.xml", "word/styles.xml", "word/theme/theme1.xml", "custom/new.xml"}); err != nil {
		t.Fatalf("declared mutations should pass: %v", err)
	}
}

func TestInspectRejectsNonOPCAndAmbiguousPackages(t *testing.T) {
	withoutRoots := buildPackage(t, map[string][]byte{"word/document.xml": []byte("doc")}, false, false)
	if _, err := officecompat.Inspect(withoutRoots); err == nil {
		t.Fatal("ZIP without OPC roots should fail")
	}

	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for _, name := range []string{"[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/document.xml"} {
		writer, err := zw.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := writer.Write([]byte(name)); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	if _, err := officecompat.Inspect(buf.Bytes()); err == nil {
		t.Fatal("duplicate OPC part should fail")
	}
}

func TestRoundTripXLSXMutationPreservesEveryOtherPart(t *testing.T) {
	original := minimalXLSX(t)
	updatedSheet := []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>after</t></is></c></row></sheetData></worksheet>`)
	candidate, err := xlsxpatch.Apply(original, xlsxpatch.Patch{Replace: map[string][]byte{
		"xl/worksheets/sheet1.xml": updatedSheet,
	}})
	if err != nil {
		t.Fatal(err)
	}
	assertSingleScopedMutation(t, original, candidate, "xl/worksheets/sheet1.xml")
}

func TestRoundTripPPTXMutationPreservesEveryOtherPart(t *testing.T) {
	beforeDeck := pptxpatch.ExampleDeck()
	afterDeck := pptxpatch.ExampleDeck()
	afterDeck.Slides[0].Shapes[0].Paragraphs[0].Runs[0].Text = "Native Office baseline"
	original, err := pptxpatch.BuildPPTX(beforeDeck)
	if err != nil {
		t.Fatal(err)
	}
	candidate, err := pptxpatch.BuildPPTX(afterDeck)
	if err != nil {
		t.Fatal(err)
	}
	assertSingleScopedMutation(t, original, candidate, "ppt/slides/slide1.xml")
}

func TestRoundTripDOCXMutationPreservesEveryOtherPart(t *testing.T) {
	original := minimalDOCX(t)
	candidate, err := docxpatch.Apply(original, []docxpatch.Edit{{Op: "set", Index: 0, Text: "Native Office baseline"}})
	if err != nil {
		t.Fatal(err)
	}
	assertSingleScopedMutation(t, original, candidate, "word/document.xml")
}

func assertSingleScopedMutation(t *testing.T, before, after []byte, mutable string) {
	t.Helper()
	beforeInventory, err := officecompat.Inspect(before)
	if err != nil {
		t.Fatal(err)
	}
	afterInventory, err := officecompat.Inspect(after)
	if err != nil {
		t.Fatal(err)
	}
	if beforeInventory.Fingerprint == afterInventory.Fingerprint {
		t.Fatal("edited package fingerprint did not change")
	}
	if err := officecompat.RequireUntouchedParts(before, after, []string{mutable}); err != nil {
		t.Fatalf("mutation escaped %s: %v", mutable, err)
	}
	if err := officecompat.RequireUntouchedParts(before, after, nil); err == nil {
		t.Fatalf("baseline did not detect changed part %s", mutable)
	}
	report, err := officecompat.ComparePackages(before, after)
	if err != nil {
		t.Fatal(err)
	}
	if len(report.Changed) != 1 || report.Changed[0].Before.Name != mutable {
		t.Fatalf("unexpected lexical package diff: %+v", report)
	}
	if report.Changed[0].XML == nil || report.Changed[0].XML.StructuralStatus != officecompat.XMLStructuralDifferent {
		t.Fatalf("native XML mutation lacks structural diff: %+v", report.Changed[0])
	}
}

func minimalXLSX(t *testing.T) []byte {
	t.Helper()
	entries := map[string][]byte{
		"[Content_Types].xml":        []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`),
		"_rels/.rels":                []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`),
		"xl/workbook.xml":            []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`),
		"xl/_rels/workbook.xml.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`),
		"xl/worksheets/sheet1.xml":   []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>before</t></is></c></row></sheetData></worksheet>`),
		"customXml/item1.xml":        []byte(`<compatibility-fixture preserve="true"/>`),
	}
	return buildPackage(t, entries, false, false)
}

func minimalDOCX(t *testing.T) []byte {
	t.Helper()
	entries := map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`),
		"_rels/.rels":         []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`),
		"word/document.xml":   []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>before</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`),
		"word/styles.xml":     []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>`),
		"customXml/item1.xml": []byte(`<compatibility-fixture preserve="true"/>`),
	}
	return buildPackage(t, entries, false, false)
}

func buildPackage(t *testing.T, entries map[string][]byte, reverse, store bool) []byte {
	t.Helper()
	names := make([]string, 0, len(entries))
	for name := range entries {
		names = append(names, name)
	}
	sort.Strings(names)
	if reverse {
		for left, right := 0, len(names)-1; left < right; left, right = left+1, right-1 {
			names[left], names[right] = names[right], names[left]
		}
	}

	var buf bytes.Buffer
	zw := zip.NewWriter(&buf)
	for _, name := range names {
		header := &zip.FileHeader{Name: name, Method: zip.Deflate}
		if store {
			header.Method = zip.Store
		}
		writer, err := zw.CreateHeader(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := writer.Write(entries[name]); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return buf.Bytes()
}

func cloneEntries(source map[string][]byte) map[string][]byte {
	cloned := make(map[string][]byte, len(source))
	for name, payload := range source {
		cloned[name] = bytes.Clone(payload)
	}
	return cloned
}
