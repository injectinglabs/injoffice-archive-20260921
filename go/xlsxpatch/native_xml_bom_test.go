package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"strings"
	"testing"
)

const nativeTestUTF8BOM = "\xef\xbb\xbf"

func TestNativeXMLLeadingUTF8BOM(t *testing.T) {
	for _, body := range []string{`<root/>`, `<?xml version="1.0" encoding="UTF-8"?><root/>`, "<root>\ufefftext\ufeff</root>"} {
		plainTokens, plainElements, err := preflightNativeCoreXML([]byte(body))
		if err != nil {
			t.Fatal(err)
		}
		data := []byte(nativeTestUTF8BOM + body)
		before := bytes.Clone(data)
		tokens, elements, err := preflightNativeCoreXML(data)
		if err != nil || tokens != plainTokens || elements != plainElements || !bytes.Equal(data, before) {
			t.Fatalf("BOM decoding changed structure/source: %d/%d %v", tokens, elements, err)
		}
	}
	for name, data := range map[string]string{
		"double":                  nativeTestUTF8BOM + nativeTestUTF8BOM + `<root/>`,
		"after-space":             " " + nativeTestUTF8BOM + `<root/>`,
		"after-comment":           `<!--x-->` + nativeTestUTF8BOM + `<root/>`,
		"after-declaration":       `<?xml version="1.0"?>` + nativeTestUTF8BOM + `<root/>`,
		"after-root":              `<root/>` + nativeTestUTF8BOM,
		"declaration-after-space": nativeTestUTF8BOM + ` <?xml version="1.0"?><root/>`,
		"doctype":                 nativeTestUTF8BOM + `<!DOCTYPE root><root/>`,
		"pi":                      nativeTestUTF8BOM + `<?evil x?><root/>`,
		"broken":                  nativeTestUTF8BOM + `<root>`,
		"partial-signature":       "\xef\xbb<root/>",
		"utf16":                   "\xff\xfe<root/>",
		"wrong-encoding":          nativeTestUTF8BOM + `<?xml version="1.0" encoding="UTF-16"?><root/>`,
	} {
		t.Run(name, func(t *testing.T) {
			if _, _, err := preflightNativeCoreXML([]byte(data)); err == nil {
				t.Fatal("invalid XML accepted")
			}
		})
	}
}

func TestNativeXMLLeadingUTF8BOMRouting(t *testing.T) {
	body := `<Relationships xmlns="` + packageRelationshipsNamespace + `"><Relationship Id="r1" Type="urn:test" Target="part.xml"/></Relationships>`
	for _, prefix := range []string{nativeTestUTF8BOM, nativeTestUTF8BOM + `<?xml version="1.0" encoding="UTF-8"?>`} {
		relationships, err := parseRoutingRelationships([]byte(prefix + body))
		if err != nil || len(relationships) != 1 {
			t.Fatalf("legal BOM routing failed: %v", err)
		}
	}
	for _, data := range []string{nativeTestUTF8BOM + nativeTestUTF8BOM + body, " " + nativeTestUTF8BOM + body, body + nativeTestUTF8BOM, nativeTestUTF8BOM + `<!DOCTYPE Relationships>` + body, nativeTestUTF8BOM + strings.Replace(body, packageRelationshipsNamespace, "urn:foreign", 1), nativeTestUTF8BOM + ` <?xml version="1.0"?>` + body} {
		if _, err := parseRoutingRelationships([]byte(data)); err == nil {
			t.Fatal("invalid BOM routing accepted")
		}
	}
}

func TestNativeXMLLeadingUTF8BOMPublicAndMutationPreservation(t *testing.T) {
	for _, strict := range []bool{false, true} {
		t.Run(map[bool]string{false: "transitional", true: "strict"}[strict], func(t *testing.T) {
			entries := nativeMutationFixture(strict)
			entries["Charts/chart1.xml"] = previewChartFixture()
			entries["Sheets/s1.xml"] = strings.Replace(entries["Sheets/s1.xml"], "<t>old</t>", "<t>\ufeffold\ufeff</t>", 1)
			for name, content := range entries {
				if strings.HasPrefix(content, "<") {
					entries[name] = nativeTestUTF8BOM + content
				}
			}
			source := buildZip(t, entries)
			before := bytes.Clone(source)
			workbook, err := ExtractNativeWorkbookV2(source)
			if err != nil {
				t.Fatal(err)
			}
			objects, err := InspectNativeWorkbookObjectsV1(source)
			if err != nil {
				t.Fatal(err)
			}
			if workbook.Source.PackageSHA256 != nativeWorkbookDigest(source) || objects.PackageSHA256 != nativeWorkbookDigest(source) || !bytes.Equal(source, before) {
				t.Fatal("public reads changed source/hash")
			}
			v1, err := ExtractNativeWorkbookV1(source)
			if err != nil {
				t.Fatal(err)
			}
			cell := findNativeCell(t, v1, "7", "B1")
			if cell.Value == nil || cell.Value.Text == nil || *cell.Value.Text != "\ufeffold\ufeff" {
				t.Fatal("in-element U+FEFF was stripped")
			}
			change := mutation("bom-change", "7", CellSetValue, 0, 0)
			change.Value = "changed"
			result, err := ApplyNativeWorkbookMutationTransactionV1(source, NativeWorkbookMutationTransactionV1{ExpectedRevision: v1.Revision, Cells: []CellMutation{change}})
			if err != nil {
				t.Fatal(err)
			}
			if !bytes.Equal(source, before) {
				t.Fatal("mutation changed caller source")
			}
			if err := compareUntouchedRawZipEntries(source, result.Package, map[string]bool{"Book/Workbook.xml": true, "Sheets/s1.xml": true}); err != nil {
				t.Fatal(err)
			}
			for _, part := range []string{"Sheets/s1.xml", "Book/Workbook.xml"} {
				if !strings.HasPrefix(readEntry(t, result.Package, part), nativeTestUTF8BOM) {
					t.Fatalf("BOM removed from %s", part)
				}
			}
			saved := findNativeCell(t, result.Workbook, "7", "B1")
			if saved.Value == nil || saved.Value.Text == nil || *saved.Value.Text != "\ufeffold\ufeff" {
				t.Fatal("mutation lost untouched in-element U+FEFF")
			}
		})
	}
}

func TestNativeXMLLeadingUTF8BOMRootOffsets(t *testing.T) {
	for _, prefix := range []string{nativeTestUTF8BOM, nativeTestUTF8BOM + " \t\r\n", nativeTestUTF8BOM + `<?xml version="1.0"?>`} {
		data := prefix + "<root>\ufeffsaved</root>"
		decoder := xml.NewDecoder(strings.NewReader(data))
		root, err := nativeReadXMLRoot(decoder)
		if err != nil || root.Name.Local != "root" || decoder.InputOffset() != int64(len(prefix+"<root>")) {
			t.Fatalf("root/source offset mismatch: %v offset=%d", err, decoder.InputOffset())
		}
		token, err := decoder.Token()
		if err != nil || string(token.(xml.CharData)) != "\ufeffsaved" {
			t.Fatal("in-element FEFF changed")
		}
	}
	for _, prefix := range []string{nativeTestUTF8BOM + nativeTestUTF8BOM, " " + nativeTestUTF8BOM, `<!--x-->` + nativeTestUTF8BOM} {
		if _, err := nativeReadXMLRoot(xml.NewDecoder(strings.NewReader(prefix + `<root/>`))); err == nil {
			t.Fatal("noninitial signature accepted")
		}
	}
	for _, prefix := range []string{nativeTestUTF8BOM + nativeTestUTF8BOM, " " + nativeTestUTF8BOM, nativeTestUTF8BOM + ` <?xml version="1.0"?>`} {
		if _, err := workbookRootDialect([]byte(prefix + `<workbook xmlns="` + spreadsheetMLTransitional + `"/>`)); err == nil {
			t.Fatal("invalid workbook signature accepted")
		}
	}
}
