package docxpatch

import (
	"fmt"
	"strings"
	"testing"
)

func sectionsColumnsDOCX(t *testing.T, body string) []byte {
	t.Helper()
	parts := transitionalNativeParts()
	parts["Custom/Main.XML"] = `<?xml version="1.0"?><w:document xmlns:w="` + testW + `"><w:body>` + body + `</w:body></w:document>`
	return buildNativeDOCX(t, nativeEntries(parts))
}

func sectionParagraph(text, breakType, columns string) string {
	return `<w:p><w:pPr><w:sectPr><w:type w:val="` + breakType + `"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` + columns + `</w:sectPr></w:pPr><w:r><w:t>` + text + `</w:t></w:r></w:p>`
}

func TestExtractNativeSectionsAndColumnsFromRealOOXML(t *testing.T) {
	equal := `<w:cols w:num="2" w:space="720"/>`
	explicit := `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="4320" w:space="720"/><w:col w:w="4320"/></w:cols>`
	body := sectionParagraph("continuous", "continuous", equal) +
		sectionParagraph("next page", "nextPage", explicit) +
		sectionParagraph("even page", "evenPage", equal) +
		sectionParagraph("odd page", "oddPage", explicit) +
		`<w:p><w:r><w:t>next column</w:t></w:r></w:p><w:sectPr><w:type w:val="nextColumn"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` + equal + `</w:sectPr>`
	doc, err := ExtractNativeDocumentV1(sectionsColumnsDOCX(t, body))
	if err != nil {
		t.Fatal(err)
	}
	wantBreaks := []string{"continuous", "next-page", "even-page", "odd-page", "next-column"}
	if len(doc.Sections) != len(wantBreaks) {
		t.Fatalf("sections = %d, want %d", len(doc.Sections), len(wantBreaks))
	}
	for index, section := range doc.Sections {
		if section.BreakType != wantBreaks[index] {
			t.Errorf("section %d break = %q, want %q", index, section.BreakType, wantBreaks[index])
		}
		if section.StartsAtBlockID != doc.Body.Blocks[index].ID {
			t.Errorf("section %d starts at %q, want block %q", index, section.StartsAtBlockID, doc.Body.Blocks[index].ID)
		}
		if len(section.Page.ColumnDefinitions) != 2 {
			t.Fatalf("section %d definitions = %#v", index, section.Page.ColumnDefinitions)
		}
		for ordinal, column := range section.Page.ColumnDefinitions {
			if column.ID != nativeColumnID(section.ID, ordinal) || nativeIntValue(column.Ordinal) != ordinal {
				t.Errorf("section %d column %d identity = %#v", index, ordinal, column)
			}
		}
		if index%2 == 0 && section.Page.ColumnLayout != "equal-width" {
			t.Errorf("section %d layout = %q", index, section.Page.ColumnLayout)
		}
		if index%2 == 1 {
			if section.Page.ColumnLayout != "explicit" || nativeInt64Value(section.Page.ColumnDefinitions[0].WidthTwips) != 4320 || nativeInt64Value(section.Page.ColumnDefinitions[0].SpaceAfterTwips) != 720 {
				t.Errorf("section %d explicit geometry = %#v", index, section.Page)
			}
		}
	}
	first, err := EncodeNativeDocumentV1(doc)
	if err != nil {
		t.Fatal(err)
	}
	second, err := EncodeNativeDocumentV1(doc)
	if err != nil || string(first) != string(second) {
		t.Fatalf("canonical section/column encoding is not deterministic: %v", err)
	}
}

func TestExtractNativeColumnsFromStrictNamespace(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/odd/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_RELS/.RELS":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="` + relBaseStrict + `officeDocument" Target="ODD/Main.XML"/></Relationships>`,
		"Odd/Main.XML":        `<?xml version="1.0"?><w:document xmlns:w="` + testWS + `"><w:body><w:p><w:r><w:t>strict columns</w:t></w:r></w:p><w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="2" w:space="720"/></w:sectPr></w:body></w:document>`,
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Sections) != 1 || doc.Sections[0].BreakType != "odd-page" || doc.Sections[0].Page.ColumnLayout != "equal-width" || nativeIntValue(doc.Sections[0].Page.Columns) != 2 || len(doc.Sections[0].Page.ColumnDefinitions) != 2 {
		t.Fatalf("strict section/column projection was not preserved: %#v", doc.Sections)
	}
	for _, unsupported := range doc.Unsupported {
		if unsupported.Capability == "sections" {
			t.Fatalf("exact strict section geometry was refused: %#v", doc.Unsupported)
		}
	}
}

func TestExtractNativeColumnsRecordsBlockingAdversaries(t *testing.T) {
	cases := []struct {
		name string
		xml  string
		code string
	}{
		{"separator", `<w:cols w:num="2" w:space="720" w:sep="1"/>`, "COLUMN_SEPARATOR_UNSUPPORTED"},
		{"unequal", `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="4000" w:space="720"/><w:col w:w="4640"/></w:cols>`, "UNEQUAL_SECTION_COLUMNS"},
		{"count mismatch", `<w:cols w:equalWidth="0" w:num="3"><w:col w:w="4320" w:space="720"/><w:col w:w="4320"/></w:cols>`, "COLUMN_COUNT_MISMATCH"},
		{"missing width", `<w:cols w:equalWidth="0" w:num="2"><w:col w:space="720"/><w:col w:w="4320"/></w:cols>`, "INVALID_COLUMN_WIDTH"},
		{"zero width", `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="0" w:space="720"/><w:col w:w="4320"/></w:cols>`, "INVALID_COLUMN_WIDTH"},
		{"negative width", `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="-1" w:space="720"/><w:col w:w="4320"/></w:cols>`, "INVALID_COLUMN_WIDTH"},
		{"overflow width", `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="9007199254740992" w:space="720"/><w:col w:w="4320"/></w:cols>`, "INVALID_COLUMN_WIDTH"},
		{"too many", `<w:cols w:num="46" w:space="720"/>`, "INVALID_SECTION_COLUMNS"},
		{"zero count", `<w:cols w:num="0" w:space="720"/>`, "INVALID_SECTION_COLUMNS"},
		{"negative count", `<w:cols w:num="-1" w:space="720"/>`, "INVALID_SECTION_COLUMNS"},
		{"lexical count", `<w:cols w:num="two" w:space="720"/>`, "INVALID_SECTION_COLUMNS"},
		{"overflow count", `<w:cols w:num="999999999999999999999" w:space="720"/>`, "INVALID_SECTION_COLUMNS"},
		{"negative spacing", `<w:cols w:num="2" w:space="-1"/>`, "INVALID_COLUMN_SPACING"},
		{"lexical spacing", `<w:cols w:num="2" w:space="wide"/>`, "INVALID_COLUMN_SPACING"},
		{"overflow spacing", `<w:cols w:num="2" w:space="9007199254740992"/>`, "INVALID_COLUMN_SPACING"},
		{"ambiguous spacing", `<w:cols w:num="2"/>`, "AMBIGUOUS_COLUMN_SPACING"},
		{"case malformed boolean", `<w:cols w:equalWidth="FALSE" w:num="2" w:space="720"/>`, "INVALID_SECTION_COLUMNS"},
		{"case malformed separator", `<w:cols w:num="2" w:space="720" w:sep="TRUE"/>`, "COLUMN_SEPARATOR_UNSUPPORTED"},
		{"conflicting definitions", `<w:cols w:num="2" w:space="720"><w:col w:w="4320" w:space="720"/><w:col w:w="4320"/></w:cols>`, "CONFLICTING_COLUMN_DEFINITIONS"},
		{"nondivisible equal", `<w:cols w:num="2" w:space="719"/>`, "AMBIGUOUS_SECTION_COLUMNS"},
		{"explicit total mismatch", `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="4000" w:space="720"/><w:col w:w="4000"/></w:cols>`, "AMBIGUOUS_SECTION_COLUMNS"},
		{"trailing space", `<w:cols w:equalWidth="0" w:num="2"><w:col w:w="4320" w:space="720"/><w:col w:w="4320" w:space="1"/></w:cols>`, "TRAILING_COLUMN_SPACING"},
	}
	for _, test := range cases {
		t.Run(test.name, func(t *testing.T) {
			body := `<w:p><w:r><w:t>adversarial</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>` + test.xml + `</w:sectPr>`
			doc, err := ExtractNativeDocumentV1(sectionsColumnsDOCX(t, body))
			if err != nil {
				t.Fatal(err)
			}
			found := false
			for _, unsupported := range doc.Unsupported {
				if unsupported.Code == test.code && unsupported.Capability == "sections" && unsupported.ScopeID == doc.Sections[0].ID {
					found = true
				}
			}
			if !found {
				codes := make([]string, len(doc.Unsupported))
				for index, unsupported := range doc.Unsupported {
					codes[index] = unsupported.Code
				}
				t.Fatalf("missing %s in %s", test.code, strings.Join(codes, ","))
			}
			if _, err := EncodeNativeDocumentV1(doc); err != nil {
				t.Fatalf("blocking extraction must remain a valid no-partial native contract: %v\n%s", err, fmt.Sprint(doc.Sections[0].Page))
			}
		})
	}
}

func TestExtractNativeColumnsAcceptsExactFortyFiveColumnBoundary(t *testing.T) {
	body := `<w:p><w:r><w:t>boundary</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:num="45" w:space="0"/></w:sectPr>`
	doc, err := ExtractNativeDocumentV1(sectionsColumnsDOCX(t, body))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Sections) != 1 || nativeIntValue(doc.Sections[0].Page.Columns) != 45 || len(doc.Sections[0].Page.ColumnDefinitions) != 45 {
		t.Fatalf("45-column boundary was not preserved: %#v", doc.Sections)
	}
	for _, unsupported := range doc.Unsupported {
		if unsupported.Capability == "sections" {
			t.Fatalf("exact 45-column boundary was refused: %#v", doc.Unsupported)
		}
	}
}

func TestExtractNativeSectionsRefuseImplicitOrPartialPageGeometry(t *testing.T) {
	cases := []struct {
		name       string
		properties string
		code       string
	}{
		{"missing page size", `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols/>`, "MISSING_PAGE_SIZE"},
		{"missing page margins", `<w:pgSz w:w="12240" w:h="15840"/><w:cols/>`, "MISSING_PAGE_MARGINS"},
		{"partial page margins", `<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/><w:cols/>`, "MISSING_PAGE_MARGIN"},
	}
	for _, test := range cases {
		t.Run(test.name, func(t *testing.T) {
			body := `<w:p><w:r><w:t>geometry</w:t></w:r></w:p><w:sectPr>` + test.properties + `</w:sectPr>`
			doc, err := ExtractNativeDocumentV1(sectionsColumnsDOCX(t, body))
			if err != nil {
				t.Fatal(err)
			}
			for _, unsupported := range doc.Unsupported {
				if unsupported.Code == test.code && unsupported.Capability == "sections" && unsupported.ScopeID == doc.Sections[0].ID {
					if _, err := EncodeNativeDocumentV1(doc); err != nil {
						t.Fatalf("blocking geometry record must retain a valid all-or-nothing native contract: %v", err)
					}
					return
				}
			}
			t.Fatalf("missing blocking %s record: %#v", test.code, doc.Unsupported)
		})
	}
}
