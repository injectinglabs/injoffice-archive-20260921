package docxpatch

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"fmt"
	"io"
	"sort"
	"strings"
	"testing"
)

const (
	testW  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	testR  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	testWS = "http://purl.oclc.org/ooxml/wordprocessingml/main"
	testRS = "http://purl.oclc.org/ooxml/officeDocument/relationships"
)

type nativeZipEntry struct {
	name   string
	data   []byte
	method uint16
}

func buildNativeDOCX(t *testing.T, entries []nativeZipEntry) []byte {
	t.Helper()
	var out bytes.Buffer
	zw := zip.NewWriter(&out)
	for _, entry := range entries {
		header := &zip.FileHeader{Name: entry.name, Method: entry.method}
		if header.Method == 0 {
			header.Method = zip.Deflate
		}
		writer, err := zw.CreateHeader(header)
		if err != nil {
			t.Fatal(err)
		}
		if _, err := writer.Write(entry.data); err != nil {
			t.Fatal(err)
		}
	}
	if err := zw.Close(); err != nil {
		t.Fatal(err)
	}
	return out.Bytes()
}

func nativeEntries(parts map[string]string) []nativeZipEntry {
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	entries := make([]nativeZipEntry, 0, len(names))
	for _, name := range names {
		entries = append(entries, nativeZipEntry{name: name, data: []byte(parts[name]), method: zip.Deflate})
	}
	return entries
}

func transitionalNativeParts() map[string]string {
	parts := map[string]string{
		"[Content_Types].xml":        `<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/custom/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/custom/stories/headera.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/custom/stories/footera.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/custom/notes/foot.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/custom/notes/end.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/custom/notes/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>`,
		"_rels/.rels":                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdOffice" Type="` + relBaseTransitional + `officeDocument" Target="custom/MAIN.xml"/></Relationships>`,
		"Custom/Main.XML":            `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="` + testW + `" xmlns:r="` + testR + `" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:aext="urn:alpha" xmlns:bext="urn:beta"><w:body><aext:widget/><bext:widget/><w:p w14:paraId="00112233"><w:pPr><w:pStyle w:val="Heading1"/><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1000" w:right="1100" w:bottom="1200" w:left="1300" w:header="500" w:footer="600" w:gutter="0"/><w:headerReference w:type="default" r:id="rHeader"/></w:sectPr></w:pPr><w:commentRangeStart w:id="3"/><w:r><w:rPr><w:b/><w:color w:val="00aaFF"/></w:rPr><w:t xml:space="preserve">Hello 世界 👋 </w:t><w:tab/></w:r><w:hyperlink r:id="rLink"><w:r><w:t>relative link</w:t></w:r></w:hyperlink><w:r><w:fldChar w:fldCharType="begin"/><w:instrText> DATE </w:instrText><w:t>field result</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:endnoteReference w:id="2"/></w:r><w:r><w:commentReference w:id="3"/></w:r><w:commentRangeEnd w:id="3"/></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:tblHeader/><w:trHeight w:val="480"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2400"/><w:gridSpan w:val="1"/></w:tcPr><w:p w14:paraId="AABBCCDD"><w:r><w:t>Cell Ω</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p w14:paraId="44556677"><w:r><w:t>Second section</w:t><w:br w:type="page"/></w:r></w:p><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/><w:cols w:num="2" w:space="360"/><w:footerReference w:type="default" r:id="rFooter"/></w:sectPr></w:body></w:document>`,
		"Custom/_RELS/Main.XML.RELS": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rHeader" Type="` + relBaseTransitional + `header" Target="stories/headera.xml"/><Relationship Id="rFooter" Type="` + relBaseTransitional + `footer" Target="stories/FOOTERA.xml"/><Relationship Id="rFoot" Type="` + relBaseTransitional + `footnotes" Target="notes/foot.xml"/><Relationship Id="rEnd" Type="` + relBaseTransitional + `endnotes" Target="notes/end.xml"/><Relationship Id="rComments" Type="` + relBaseTransitional + `comments" Target="notes/comments.xml"/><Relationship Id="rLink" Type="` + relBaseTransitional + `hyperlink" Target="../relative/link.html" TargetMode="External"/><Relationship Id="rImage" Type="` + relBaseTransitional + `image" Target="media/image.png"/></Relationships>`,
		"Custom/Stories/HeaderA.XML": `<w:hdr xmlns:w="` + testW + `"><w:p><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>`,
		"Custom/Stories/FooterA.XML": `<w:ftr xmlns:w="` + testW + `"><w:p><w:r><w:t>Native footer</w:t></w:r></w:p></w:ftr>`,
		"Custom/Notes/Foot.XML":      `<w:footnotes xmlns:w="` + testW + `"><w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:r><w:t>Footnote α</w:t></w:r></w:p></w:footnote></w:footnotes>`,
		"Custom/Notes/End.XML":       `<w:endnotes xmlns:w="` + testW + `"><w:endnote w:id="2"><w:p><w:r><w:t>Endnote β</w:t></w:r></w:p></w:endnote></w:endnotes>`,
		"Custom/Notes/Comments.XML":  `<w:comments xmlns:w="` + testW + `"><w:comment w:id="3" w:author="Ada" w:initials="AL" w:date="2026-08-27T12:00:00Z"><w:p><w:r><w:t>Comment γ</w:t></w:r></w:p></w:comment></w:comments>`,
		"Custom/Media/image.PNG":     "native-image-bytes",
		"Custom/Styles.XML":          `<w:styles xmlns:w="` + testW + `"/>`,
	}
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `xmlns:aext="urn:alpha"`, `xmlns:wp="`+wordDrawingTransitional+`" xmlns:a="`+drawingMLTransitional+`" xmlns:pic="`+pictureMLTransitional+`" xmlns:aext="urn:alpha"`, 1)
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `w14:paraId="AABBCCDD"`, `w14:paraId="2ABBCCDD"`, 1)
	drawing := `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="Native picture" descr="Accessible image"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="Native picture"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm/><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:hyperlink r:id="rLink">`, drawing+`<w:hyperlink r:id="rLink">`, 1)
	return parts
}

func TestExtractNativeDocumentV1EndToEnd(t *testing.T) {
	data := buildNativeDOCX(t, nativeEntries(transitionalNativeParts()))
	doc, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	if doc.Source.MainPart != "Custom/Main.XML" {
		t.Fatalf("main part = %q", doc.Source.MainPart)
	}
	if len(doc.Sections) != 2 || len(doc.Headers) != 1 || len(doc.Footers) != 1 || len(doc.Notes) != 3 || len(doc.Comments) != 1 || len(doc.CommentStories) != 1 {
		t.Fatalf("unexpected related extraction: sections=%d headers=%d footers=%d notes=%d comments=%d comment stories=%d", len(doc.Sections), len(doc.Headers), len(doc.Footers), len(doc.Notes), len(doc.Comments), len(doc.CommentStories))
	}
	if len(doc.Body.Blocks) != 3 || doc.Body.Blocks[1].Kind != "table" {
		t.Fatalf("unexpected body blocks: %#v", doc.Body.Blocks)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph == nil || paragraph.EditPolicy.Mode != "read-only" || paragraph.EditPolicy.Refusal == nil {
		t.Fatalf("complex paragraph must be honestly read-only: %#v", paragraph)
	}
	var text strings.Builder
	var kinds []string
	var drawing *NativeDrawingV1
	for _, run := range paragraph.Runs {
		kinds = append(kinds, run.Kind)
		if run.Text != nil {
			text.WriteString(*run.Text)
		}
		if run.Drawing != nil {
			drawing = run.Drawing
		}
	}
	if !strings.Contains(text.String(), "Hello 世界 👋") || !strings.Contains(text.String(), "relative link") || !strings.Contains(text.String(), "field result") {
		t.Fatalf("visible Unicode/text extraction lost content: %q", text.String())
	}
	if !strings.Contains(strings.Join(kinds, ","), "reference") {
		t.Fatalf("expected native note/comment references: %v", kinds)
	}
	if drawing == nil || drawing.Placement != "inline" || drawing.MediaPart == nil || *drawing.MediaPart != "Custom/Media/image.PNG" || *drawing.WidthEMU != 914400 || drawing.EditPolicy.Mode != "read-only" {
		t.Fatalf("basic native picture was not modeled conservatively: %#v", drawing)
	}
	if doc.Body.Blocks[1].Table.EditPolicy.Mode != "read-only" {
		t.Fatal("extract-only tables must not advertise mutation support")
	}
	if got := *doc.Body.Blocks[1].Table.Rows[0].Cells[0].Paragraphs[0].Runs[0].Text; got != "Cell Ω" {
		t.Fatalf("table cell text = %q", got)
	}
	if doc.Sections[0].HeaderRefs[0].StoryID != doc.Headers[0].ID || doc.Sections[1].FooterRefs[0].StoryID != doc.Footers[0].ID {
		t.Fatal("section header/footer references were not typed to stories")
	}
	if doc.Comments[0].BodyStoryID != doc.CommentStories[0].ID || *doc.CommentStories[0].NativeStoryID != doc.Comments[0].NativeCommentID {
		t.Fatal("comment metadata/body story identity mismatch")
	}
	assertNativeAnchors(t, doc, transitionalNativeParts())
	if len(doc.PassthroughParts) < 4 || !hasNativePassthrough(doc, "Custom/Media/image.PNG") || !hasNativePassthrough(doc, "Custom/_RELS/Main.XML.RELS") || !hasNativePassthrough(doc, "Custom/Styles.XML") {
		t.Fatalf("incomplete preserve-verbatim inventory: %#v", doc.PassthroughParts)
	}
	if !hasUnsupportedCode(doc, "HYPERLINK_SEMANTICS") || !hasUnsupportedCode(doc, "FIELD_SEMANTICS") {
		t.Fatalf("missing explicit unsupported records: %#v", doc.Unsupported)
	}
	var extensionPaths []string
	for _, entry := range doc.Unsupported {
		if entry.Code == "UNMODELED_BODY_BLOCK" && entry.Anchor != nil {
			extensionPaths = append(extensionPaths, entry.Anchor.Path)
		}
	}
	if len(extensionPaths) != 2 || extensionPaths[0] == extensionPaths[1] {
		t.Fatalf("extension namespace paths must remain distinct: %v", extensionPaths)
	}
	encodedA, err := EncodeNativeDocumentV1(doc)
	if err != nil {
		t.Fatal(err)
	}
	docAgain, err := ExtractNativeDocumentV1(data)
	if err != nil {
		t.Fatal(err)
	}
	encodedB, _ := EncodeNativeDocumentV1(docAgain)
	if !bytes.Equal(encodedA, encodedB) {
		t.Fatal("identical package extraction was not deterministic")
	}
}

func TestExtractNativeDocumentV1AcceptsASCIIcaseInsensitiveImageMIME(t *testing.T) {
	parts := transitionalNativeParts()
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `ContentType="image/png"`, `ContentType="IMAGE/PNG"`, 1)
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	var drawing *NativeDrawingV1
	for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
		if run.Drawing != nil {
			drawing = run.Drawing
		}
	}
	if drawing == nil || drawing.ContentType == nil || *drawing.ContentType != "IMAGE/PNG" {
		t.Fatalf("case-insensitive image MIME was not preserved and classified: %#v", drawing)
	}
}

func TestExtractNativeDocumentProjectsExactBoundedTablePaintProperties(t *testing.T) {
	parts := transitionalNativeParts()
	main := parts["Custom/Main.XML"]
	main = strings.Replace(main, `<w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr>`, `<w:tblPr><w:tblW w:w="2400" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="8" w:color="112233"/><w:right w:val="single" w:sz="8" w:color="112233"/><w:bottom w:val="single" w:sz="8" w:color="112233"/><w:left w:val="single" w:sz="8" w:color="112233"/></w:tblBorders></w:tblPr>`, 1)
	main = strings.Replace(main, `<w:trPr><w:tblHeader/><w:trHeight w:val="480"/></w:trPr>`, `<w:trPr><w:tblHeader w:val="0"/><w:cantSplit/></w:trPr>`, 1)
	main = strings.Replace(main, `<w:tcPr><w:tcW w:w="2400"/><w:gridSpan w:val="1"/></w:tcPr>`, `<w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:gridSpan w:val="1"/><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr>`, 1)
	parts["Custom/Main.XML"] = main
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	table := doc.Body.Blocks[1].Table
	if table == nil || len(table.GridWidthsTwips) != 1 || table.GridWidthsTwips[0] != 2400 || table.WidthTwips == nil || *table.WidthTwips != 2400 || table.Layout == nil || *table.Layout != "fixed" || table.Alignment == nil || *table.Alignment != "left" || table.IndentTwips == nil || *table.IndentTwips != 0 || table.CellMargins == nil || table.Borders == nil {
		t.Fatalf("exact table projection missing: %#v", table)
	}
	row := table.Rows[0]
	cell := row.Cells[0]
	if row.CantSplit == nil || !*row.CantSplit || row.RepeatHeader == nil || *row.RepeatHeader || cell.ShadingRGB == nil || *cell.ShadingRGB != "DDEEFF" {
		t.Fatalf("exact row/cell projection missing: row=%#v cell=%#v", row, cell)
	}
	for _, unsupported := range doc.Unsupported {
		if unsupported.ScopeID == table.ID && (unsupported.Code == "TABLE_GRID_PRESERVED" || unsupported.Code == "UNMODELED_TABLE_PROPERTY" || unsupported.Code == "UNMODELED_CELL_PROPERTY") {
			t.Fatalf("exact modeled table property was incorrectly refused: %#v", unsupported)
		}
	}
}

func TestExtractNativeDocumentProjectsExactThemeSrgbRunAndTablePaint(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="theme" Type="` + relBaseTransitional + `theme" Target="theme/theme1.xml"/></Relationships>`,
		"word/theme/theme1.xml":        nativeThemeSrgbFixtureXML(drawingMLTransitional),
		"word/document.xml": `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` +
			`<w:p><w:r><w:rPr><w:color w:val="FF0000" w:themeColor="accent1"/></w:rPr><w:t>theme</w:t></w:r></w:p>` +
			`<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="100" w:type="dxa"/><w:right w:w="100" w:type="dxa"/><w:bottom w:w="100" w:type="dxa"/><w:left w:w="100" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:color="FF0000" w:themeColor="accent1"/><w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="8" w:space="0" w:themeColor="hyperlink"/><w:left w:val="single" w:sz="8" w:color="112233"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FF0000" w:themeFill="accent2"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
			`<w:sectPr/></w:body></w:document>`,
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph == nil || len(paragraph.Runs) != 1 || paragraph.Runs[0].Properties == nil || paragraph.Runs[0].Properties.Color == nil || *paragraph.Runs[0].Properties.Color != "4472C4" {
		t.Fatalf("theme srgb run color was not projected: %#v", paragraph)
	}
	for _, unsupported := range doc.Unsupported {
		if unsupported.Code == "PARTIAL_RUN_PROPERTIES" {
			t.Fatalf("exact theme srgb run color was treated as partial: %#v", unsupported)
		}
	}
	table := doc.Body.Blocks[1].Table
	if table == nil || table.Borders == nil || table.Borders.Top == nil || table.Borders.Top.ColorRGB == nil || *table.Borders.Top.ColorRGB != "4472C4" {
		t.Fatalf("theme srgb table border was not projected: %#v", table)
	}
	if table.Borders.Right == nil || table.Borders.Right.Style != "none" {
		t.Fatalf("identity none border with space=0 was not projected: %#v", table.Borders)
	}
	if table.Borders.Bottom == nil || table.Borders.Bottom.ColorRGB == nil || *table.Borders.Bottom.ColorRGB != "0563C1" {
		t.Fatalf("themeColor-only table border was not projected from srgb: %#v", table.Borders.Bottom)
	}
	if table.Borders.Left == nil || table.Borders.Left.ColorRGB == nil || *table.Borders.Left.ColorRGB != "112233" {
		t.Fatalf("explicit RGB table border was lost: %#v", table.Borders.Left)
	}
	cell := table.Rows[0].Cells[0]
	if cell.ShadingRGB == nil || *cell.ShadingRGB != "ED7D31" {
		t.Fatalf("themeFill clear shading was not projected from srgb: %#v", cell)
	}
	for _, unsupported := range doc.Unsupported {
		if unsupported.ScopeID == table.ID && (unsupported.Code == "UNMODELED_TABLE_PROPERTY" || unsupported.Code == "UNMODELED_CELL_PROPERTY") {
			t.Fatalf("exact theme srgb table paint was incorrectly refused: %#v", unsupported)
		}
	}

	parts["word/document.xml"] = strings.ReplaceAll(parts["word/document.xml"], `w:themeColor="accent1"`, `w:themeColor="accent1" w:themeTint="99"`)
	tintedDoc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	sawPartial := false
	for _, unsupported := range tintedDoc.Unsupported {
		if unsupported.Code == "PARTIAL_RUN_PROPERTIES" {
			sawPartial = true
		}
	}
	if !sawPartial {
		t.Fatalf("themeTint run color was not fail-closed: %#v", tintedDoc.Unsupported)
	}
	if tintedDoc.Body.Blocks[1].Table == nil || tintedDoc.Body.Blocks[1].Table.EditPolicy.Refusal == nil || tintedDoc.Body.Blocks[1].Table.EditPolicy.Refusal.Code != "UNMODELED_TABLE_MARKUP" {
		t.Fatalf("themeTint table border was not fail-closed: %#v", tintedDoc.Body.Blocks[1].Table)
	}
}

func TestExtractNativeDocumentProjectsSysClrLastClrAndRowHeightRule(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml":          `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
		"_rels/.rels":                  `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/_rels/document.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="theme" Type="` + relBaseTransitional + `theme" Target="theme/theme1.xml"/></Relationships>`,
		"word/theme/theme1.xml":        nativeThemeSrgbFixtureXML(drawingMLTransitional),
		"word/document.xml": `<w:document xmlns:w="` + wordMLTransitional + `"><w:body>` +
			`<w:p><w:r><w:rPr><w:color w:themeColor="text1"/></w:rPr><w:t>sys</w:t></w:r></w:p>` +
			`<w:tbl><w:tblPr><w:tblW w:w="2400" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:jc w:val="left"/><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/></w:tblCellMar><w:tblBorders><w:top w:val="single" w:sz="8" w:space="0" w:themeColor="light1"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="480" w:hRule="atLeast"/><w:cantSplit/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2400" w:type="dxa"/><w:gridSpan w:val="1"/></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
			`<w:sectPr/></w:body></w:document>`,
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	paragraph := doc.Body.Blocks[0].Paragraph
	if paragraph == nil || len(paragraph.Runs) != 1 || paragraph.Runs[0].Properties == nil || paragraph.Runs[0].Properties.Color == nil || *paragraph.Runs[0].Properties.Color != "000000" {
		t.Fatalf("sysClr lastClr run color was not projected: %#v", paragraph)
	}
	table := doc.Body.Blocks[1].Table
	if table == nil || table.Borders == nil || table.Borders.Top == nil || table.Borders.Top.ColorRGB == nil || *table.Borders.Top.ColorRGB != "FFFFFF" {
		t.Fatalf("sysClr lastClr table border was not projected: %#v", table)
	}
	row := table.Rows[0]
	if row.HeightTwips == nil || *row.HeightTwips != 480 || row.HeightRule == nil || *row.HeightRule != "atLeast" || row.CantSplit == nil || !*row.CantSplit {
		t.Fatalf("atLeast row height was not projected: %#v", row)
	}

	parts["word/document.xml"] = strings.ReplaceAll(parts["word/document.xml"], `w:hRule="atLeast"`, `w:hRule="auto"`)
	autoDoc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if autoDoc.Body.Blocks[1].Table == nil || autoDoc.Body.Blocks[1].Table.Rows[0].HeightTwips != nil || autoDoc.Body.Blocks[1].Table.Rows[0].HeightRule != nil {
		t.Fatalf("auto row height val was treated as a layout constraint: %#v", autoDoc.Body.Blocks[1].Table)
	}
}

func TestExtractNativeDocumentDefaultsOmittedCellWidthTypeToDxa(t *testing.T) {
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(transitionalNativeParts())))
	if err != nil {
		t.Fatal(err)
	}
	table := doc.Body.Blocks[1].Table
	if table == nil || len(table.Rows) != 1 || len(table.Rows[0].Cells) != 1 {
		t.Fatalf("baseline table projection missing: %#v", table)
	}
	cell := table.Rows[0].Cells[0]
	if cell.WidthTwips == nil || *cell.WidthTwips != 2400 {
		t.Fatalf("omitted CT_TblWidth type did not default to dxa: %#v", cell.WidthTwips)
	}
	if table.EditPolicy.Refusal != nil && table.EditPolicy.Refusal.Code == "UNMODELED_TABLE_MARKUP" {
		t.Fatalf("schema-defaulted cell width incorrectly made the table unsafe: %#v", table.EditPolicy)
	}
}

func assertNativeAnchors(t *testing.T, doc *NativeDocumentV1, parts map[string]string) {
	t.Helper()
	check := func(anchor *NativeSourceAnchorV1) {
		if anchor == nil || anchor.StartByte == nil || anchor.EndByte == nil {
			t.Fatal("missing anchor extent")
		}
		raw := []byte(parts[anchor.PartName])
		if *anchor.StartByte < 0 || *anchor.EndByte > int64(len(raw)) || *anchor.EndByte <= *anchor.StartByte || raw[*anchor.StartByte] != '<' {
			t.Fatalf("invalid exact anchor %#v for part length %d", anchor, len(raw))
		}
		digest := sha256.Sum256(raw[*anchor.StartByte:*anchor.EndByte])
		if got, want := anchor.XMLSHA256, fmt.Sprintf("sha256:%x", digest); got != want {
			t.Fatalf("anchor digest = %q, want %q", got, want)
		}
	}
	check(doc.Body.Anchor)
	for _, block := range doc.Body.Blocks {
		if block.Paragraph != nil {
			check(&block.Paragraph.Anchor)
			for _, run := range block.Paragraph.Runs {
				check(&run.Anchor)
			}
		}
		if block.Table != nil {
			check(&block.Table.Anchor)
			for _, row := range block.Table.Rows {
				check(&row.Anchor)
				for _, cell := range row.Cells {
					check(&cell.Anchor)
				}
			}
		}
	}
	for _, section := range doc.Sections {
		check(&section.Anchor)
	}
}

func hasNativePassthrough(doc *NativeDocumentV1, name string) bool {
	for _, part := range doc.PassthroughParts {
		if part.PartName == name && part.Policy == "preserve-verbatim" {
			return true
		}
	}
	return false
}

func hasUnsupportedCode(doc *NativeDocumentV1, code string) bool {
	for _, entry := range doc.Unsupported {
		if entry.Code == code {
			return true
		}
	}
	return false
}

func TestExtractNativeDocumentV1StrictAndArbitraryMainLocation(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/odd/folder/doc.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_RELS/.RELS":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="` + relBaseStrict + `officeDocument" Target="ODD/folder/DOC.xml"/></Relationships>`,
		"Odd/Folder/Doc.XML":  `<w:document xmlns:w="` + testWS + `" xmlns:r="` + testRS + `"><w:body><w:p><w:r><w:t>Strict native</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	got := *doc.Body.Blocks[0].Paragraph.Runs[0].Text
	if doc.Source.MainPart != "Odd/Folder/Doc.XML" || got != "Strict native" {
		t.Fatalf("strict extraction failed: %#v", doc)
	}
}

func TestExtractNativeDocumentHeaderFooterPagePaintInputsInBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS, relNS, relBase := "Transitional", testW, testR, relBaseTransitional
		if strict {
			name, wordNS, relNS, relBase = "Strict", testWS, testRS, relBaseStrict
		}
		t.Run(name, func(t *testing.T) {
			parts := map[string]string{
				"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/odd/main.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/odd/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/odd/footer.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`,
				"_rels/.rels":         `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBase + `officeDocument" Target="odd/main.xml"/></Relationships>`,
				"odd/main.xml": `<w:document xmlns:w="` + wordNS + `" xmlns:r="` + relNS + `"><w:body>` +
					`<w:p><w:pPr><w:sectPr><w:type w:val="nextPage"/><w:titlePg/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:headerReference w:type="first" r:id="rHeader"/></w:sectPr></w:pPr><w:r><w:t>section one</w:t></w:r></w:p>` +
					`<w:p><w:r><w:t>section two</w:t></w:r></w:p><w:sectPr><w:type w:val="oddPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:headerReference w:type="even" r:id="rHeader"/><w:footerReference w:type="default" r:id="rFooter"/></w:sectPr>` +
					`</w:body></w:document>`,
				"odd/_rels/main.xml.rels": `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="rHeader" Type="` + relBase + `header" Target="header.xml"/><Relationship Id="rFooter" Type="` + relBase + `footer" Target="footer.xml"/></Relationships>`,
				"odd/header.xml":          `<w:hdr xmlns:w="` + wordNS + `"><w:p><w:r><w:t>Exact header</w:t></w:r></w:p></w:hdr>`,
				"odd/footer.xml":          `<w:ftr xmlns:w="` + wordNS + `"><w:p><w:r><w:t>Exact footer</w:t></w:r></w:p></w:ftr>`,
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			first, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if len(first.Sections) != 2 || first.Sections[0].TitlePage == nil || !*first.Sections[0].TitlePage || first.Sections[1].TitlePage == nil || *first.Sections[1].TitlePage || first.Sections[0].HeaderRefs[0].Kind != "first" || first.Sections[1].HeaderRefs[0].Kind != "even" || first.Sections[1].FooterRefs[0].Kind != "default" {
				t.Fatalf("exact section/header/footer inputs were not modeled: %#v", first.Sections)
			}
			if first.Sections[0].HeaderRefs[0].StoryID != first.Headers[0].ID || first.Sections[1].HeaderRefs[0].StoryID != first.Headers[0].ID || first.Sections[1].FooterRefs[0].StoryID != first.Footers[0].ID {
				t.Fatalf("relationship targets did not exact-join stories: sections=%#v headers=%#v footers=%#v", first.Sections, first.Headers, first.Footers)
			}
			encoded, err := EncodeNativeDocumentV1(first)
			if err != nil {
				t.Fatal(err)
			}
			again, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			encodedAgain, _ := EncodeNativeDocumentV1(again)
			if !bytes.Equal(encoded, encodedAgain) {
				t.Fatal("section/header/footer native extraction is not deterministic")
			}
		})
	}
}

func TestExtractNativeDocumentHeaderFooterRelationshipsFailClosed(t *testing.T) {
	base := transitionalNativeParts()
	for _, test := range []struct {
		name   string
		mutate func(map[string]string)
	}{
		{"dangling", func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `r:id="rHeader"`, `r:id="missing"`, 1)
		}},
		{"external", func(parts map[string]string) {
			parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], `Target="stories/headera.xml"`, `Target="https://example.invalid/header" TargetMode="External"`, 1)
		}},
		{"wrong kind", func(parts map[string]string) {
			parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], relBaseTransitional+`header`, relBaseTransitional+`footer`, 1)
		}},
		{"duplicate variant", func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:headerReference w:type="default" r:id="rHeader"/>`, `<w:headerReference w:type="default" r:id="rHeader"/><w:headerReference w:type="default" r:id="rHeader"/>`, 1)
		}},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := make(map[string]string, len(base))
			for key, value := range base {
				parts[key] = value
			}
			test.mutate(parts)
			if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil {
				t.Fatal("malformed header/footer relationship produced a partial native document")
			}
		})
	}
}

func TestExtractNativeDocumentIdentityContinuity(t *testing.T) {
	parts := transitionalNativeParts()
	firstBytes := buildNativeDOCX(t, nativeEntries(parts))
	first, err := ExtractNativeDocumentV1(firstBytes)
	if err != nil {
		t.Fatal(err)
	}
	firstTableID := first.Body.Blocks[1].ID
	firstParagraphID := first.Body.Blocks[0].ID

	parts["Custom/Media/image.PNG"] = "unrelated changed media bytes"
	unrelated, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if unrelated.DocumentID != first.DocumentID || unrelated.Body.Blocks[0].ID != firstParagraphID || unrelated.Body.Blocks[1].ID != firstTableID {
		t.Fatal("unrelated part change churned document/object identity")
	}

	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:tbl>`, `<w:p w14:paraId="5EADBEEF"><w:r><w:t>Inserted sibling</w:t></w:r></w:p><w:tbl>`, 1)
	withSibling, err := ExtractNativeDocumentV1WithOptions(buildNativeDOCX(t, nativeEntries(parts)), NativeExtractionOptions{Previous: first})
	if err != nil {
		t.Fatal(err)
	}
	if withSibling.DocumentID != first.DocumentID || withSibling.Body.Blocks[0].ID != firstParagraphID || withSibling.Body.Blocks[2].ID != firstTableID {
		t.Fatalf("previous identity did not stabilize unchanged siblings: old=%s new=%s", firstTableID, withSibling.Body.Blocks[2].ID)
	}
	explicit, err := ExtractNativeDocumentV1WithOptions(firstBytes, NativeExtractionOptions{DocumentID: "doc:application-owned"})
	if err != nil || explicit.DocumentID != "doc:application-owned" {
		t.Fatalf("explicit document identity failed: doc=%#v err=%v", explicit, err)
	}
}

func TestExtractNativeDocumentParaIDStabilizesIdenticalSiblingInsertion(t *testing.T) {
	parts := transitionalNativeParts()
	identical := `<w:p w14:paraId="11111111"><w:r><w:t>Identical sibling</w:t></w:r></w:p><w:p w14:paraId="22222222"><w:r><w:t>Identical sibling</w:t></w:r></w:p>`
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:tbl>`, identical+`<w:tbl>`, 1)
	first, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	firstA, firstB := first.Body.Blocks[1].ID, first.Body.Blocks[2].ID
	inserted := `<w:p w14:paraId="33333333"><w:r><w:t>Identical sibling</w:t></w:r></w:p>`
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], identical, inserted+identical, 1)
	next, err := ExtractNativeDocumentV1WithOptions(buildNativeDOCX(t, nativeEntries(parts)), NativeExtractionOptions{Previous: first})
	if err != nil {
		t.Fatal(err)
	}
	if next.Body.Blocks[2].ID != firstA || next.Body.Blocks[3].ID != firstB || next.Body.Blocks[1].ID == firstA || next.Body.Blocks[1].ID == firstB {
		t.Fatalf("w14:paraId identity churned across identical insertion: before=(%s,%s), after=(%s,%s,%s)", firstA, firstB, next.Body.Blocks[1].ID, next.Body.Blocks[2].ID, next.Body.Blocks[3].ID)
	}
}

func TestExtractNativeDocumentInvalidAndDuplicateParaIDsFallBack(t *testing.T) {
	for _, invalid := range []string{"00000000", "80000000", "AABBCCDD", "not-hex!"} {
		if nativeValidOfficeHexID(invalid) {
			t.Fatalf("invalid w14:paraId %q was accepted", invalid)
		}
	}
	if !nativeValidOfficeHexID("00000001") || !nativeValidOfficeHexID("7FFFFFFF") {
		t.Fatal("valid edge w14:paraId values were rejected")
	}
	parts := transitionalNativeParts()
	problemParagraphs := `<w:p w14:paraId="00112233"><w:r><w:t>Duplicate native id</w:t></w:r></w:p><w:p w14:paraId="00000000"><w:r><w:t>Zero native id</w:t></w:r></w:p><w:p w14:paraId="AABBCCDD"><w:r><w:t>Out-of-range native id</w:t></w:r></w:p>`
	parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<w:tbl>`, problemParagraphs+`<w:tbl>`, 1)
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatalf("identity hints must not make otherwise extractable content disappear: %v", err)
	}
	if !hasUnsupportedCode(doc, "DUPLICATE_NATIVE_PARAGRAPH_ID") || !hasUnsupportedCode(doc, "INVALID_NATIVE_PARAGRAPH_ID") {
		t.Fatalf("missing paraId fallback diagnostics: %#v", doc.Unsupported)
	}
	if len(doc.Body.Blocks) != 6 {
		t.Fatalf("paragraph fallback dropped content: %d blocks", len(doc.Body.Blocks))
	}
}

func TestExtractNativeDocumentPreviousIdentityCanonicalizesPartCase(t *testing.T) {
	parts := transitionalNativeParts()
	first, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	lowerParts := make(map[string]string, len(parts))
	for name, value := range parts {
		if name == "[Content_Types].xml" || name == "_rels/.rels" {
			lowerParts[name] = value
		} else {
			lowerParts[strings.ToLower(name)] = value
		}
	}
	next, err := ExtractNativeDocumentV1WithOptions(buildNativeDOCX(t, nativeEntries(lowerParts)), NativeExtractionOptions{Previous: first})
	if err != nil {
		t.Fatal(err)
	}
	if next.DocumentID != first.DocumentID || next.Body.Blocks[0].ID != first.Body.Blocks[0].ID || next.Body.Blocks[1].ID != first.Body.Blocks[1].ID || next.Sections[0].ID != first.Sections[0].ID {
		t.Fatal("ASCII case-equivalent OPC spelling churned previous native identities")
	}
}

func TestExtractNativeDocumentContentTypeEqualityIsASCIIOnly(t *testing.T) {
	for _, test := range []struct {
		name        string
		contentType string
		valid       bool
	}{
		{"ASCII case", strings.ToUpper("application/vnd.openxmlformats-package.relationships+xml"), true},
		{"Unicode long s", "application/vnd.openxmlformats-package.relationship\u017f+xml", false},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := cloneNativeParts(transitionalNativeParts())
			parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], "application/vnd.openxmlformats-package.relationships+xml", test.contentType, 1)
			_, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if test.valid {
				if err != nil {
					t.Fatalf("ASCII case-equivalent MIME should be accepted: %v", err)
				}
				return
			}
			if err == nil || !strings.Contains(err.Error(), "root relationship part") {
				t.Fatalf("non-ASCII MIME fold must be rejected, err=%v", err)
			}
		})
	}
}

func TestExtractNativeDocumentRejectsAdversarialPackages(t *testing.T) {
	base := transitionalNativeParts()
	tests := []struct {
		name    string
		entries func() []nativeZipEntry
		want    string
	}{
		{name: "duplicate exact entry", entries: func() []nativeZipEntry {
			entries := nativeEntries(base)
			entries = append(entries, nativeZipEntry{name: "Custom/Main.XML", data: []byte(base["Custom/Main.XML"]), method: zip.Deflate})
			return entries
		}, want: "duplicate ZIP entry"},
		{name: "case equivalent entry", entries: func() []nativeZipEntry {
			entries := nativeEntries(base)
			entries = append(entries, nativeZipEntry{name: "custom/main.xml", data: []byte(base["Custom/Main.XML"]), method: zip.Deflate})
			return entries
		}, want: "ambiguous ZIP entries"},
		{name: "percent equivalent entry", entries: func() []nativeZipEntry {
			entries := nativeEntries(base)
			entries = append(entries, nativeZipEntry{name: "Custom/%4Dain.XML", data: []byte(base["Custom/Main.XML"]), method: zip.Deflate})
			return entries
		}, want: "ambiguous ZIP entries"},
		{name: "traversal target", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `custom/MAIN.xml`, `../Custom/Main.XML`, 1)
			return nativeEntries(parts)
		}, want: "escapes the package root"},
		{name: "encoded traversal target", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			parts["_rels/.rels"] = strings.Replace(parts["_rels/.rels"], `custom/MAIN.xml`, `%2E%2E/Custom/Main.XML`, 1)
			return nativeEntries(parts)
		}, want: "unsafe percent-encoded"},
		{name: "missing internal target", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			delete(parts, "Custom/Media/image.PNG")
			return nativeEntries(parts)
		}, want: "targets missing part"},
		{name: "word namespace spoof", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], testW, "urn:spoofed-word", 1)
			return nativeEntries(parts)
		}, want: "expected"},
		{name: "DTD forbidden", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<?xml version="1.0" encoding="UTF-8"?>`, `<?xml version="1.0"?><!DOCTYPE document [<!ENTITY boom "boom">]>`, 1)
			return nativeEntries(parts)
		}, want: "directives/DOCTYPE"},
		{name: "unknown entity", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `Hello 世界 👋 `, `&notDeclared;`, 1)
			return nativeEntries(parts)
		}, want: "invalid character entity"},
		{name: "zip bomb ratio", entries: func() []nativeZipEntry {
			parts := cloneNativeParts(base)
			entries := nativeEntries(parts)
			entries = append(entries, nativeZipEntry{name: "Custom/Bomb.bin", data: bytes.Repeat([]byte{0}, 4*1024*1024), method: zip.Deflate})
			// Supply a content type before building.
			for index := range entries {
				if entries[index].name == "[Content_Types].xml" {
					entries[index].data = []byte(strings.Replace(string(entries[index].data), `<Default Extension="png"`, `<Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="png"`, 1))
				}
			}
			return entries
		}, want: "compression-ratio limit"},
	}
	for _, test := range tests {
		t.Run(test.name, func(t *testing.T) {
			_, err := ExtractNativeDocumentV1(buildNativeDOCX(t, test.entries()))
			if err == nil || !strings.Contains(err.Error(), test.want) {
				t.Fatalf("error = %v, want substring %q", err, test.want)
			}
		})
	}
}

func cloneNativeParts(source map[string]string) map[string]string {
	clone := make(map[string]string, len(source))
	for key, value := range source {
		clone[key] = value
	}
	return clone
}

func TestNativePackageExternalRelativeRelationshipIsPreserved(t *testing.T) {
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(transitionalNativeParts())))
	if err != nil {
		t.Fatal(err)
	}
	if !hasNativePassthrough(doc, "Custom/_RELS/Main.XML.RELS") || !hasUnsupportedCode(doc, "HYPERLINK_SEMANTICS") {
		t.Fatal("relative external relationship was not safely preserved")
	}
}

func TestExtractNativeDocumentDrawingMLFloatingAndCropRefusal(t *testing.T) {
	t.Run("floating picture", func(t *testing.T) {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent`, `<wp:anchor distT="0" distB="0" distL="0" distR="0"><wp:positionH relativeFrom="column"><wp:posOffset>12345</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>-23456</wp:posOffset></wp:positionV><wp:wrapSquare/><wp:extent`, 1)
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `</wp:inline>`, `</wp:anchor>`, 1)
		doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		var drawing *NativeDrawingV1
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				drawing = run.Drawing
			}
		}
		if drawing == nil || drawing.Placement != "floating" || drawing.XEMU == nil || *drawing.XEMU != 12345 || drawing.YEMU == nil || *drawing.YEMU != -23456 || drawing.Wrap == nil || *drawing.Wrap != "square" {
			t.Fatalf("floating picture projection = %#v", drawing)
		}
		if !hasNativePassthrough(doc, "Custom/Media/image.PNG") {
			t.Fatal("floating image bytes must remain passthrough")
		}
	})

	t.Run("out-of-bounds crop remains preserve-only", func(t *testing.T) {
		parts := transitionalNativeParts()
		parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<pic:blipFill><a:blip`, `<pic:blipFill><a:srcRect l="100000"/><a:blip`, 1)
		doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
			if run.Drawing != nil {
				t.Fatal("cropped drawing must not be projected as fully representable")
			}
		}
		if !hasUnsupportedCode(doc, "PICTURE_CROP_PRESERVED") || !hasNativePassthrough(doc, "Custom/Media/image.PNG") {
			t.Fatalf("crop refusal/passthrough missing: unsupported=%#v passthrough=%#v", doc.Unsupported, doc.PassthroughParts)
		}
	})

	for _, test := range []struct {
		name   string
		mutate func(map[string]string)
		code   string
	}{
		{name: "rotation remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<a:xfrm/>`, `<a:xfrm rot="5400000"/>`, 1)
		}, code: "PICTURE_TRANSFORM_PRESERVED"},
		{name: "mismatched transform extent remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<a:xfrm/>`, `<a:xfrm><a:off x="0" y="0"/><a:ext cx="914401" cy="457200"/></a:xfrm>`, 1)
		}, code: "PICTURE_TRANSFORM_PRESERVED"},
		{name: "hidden picture remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<wp:docPr id="1"`, `<wp:docPr hidden="1" id="1"`, 1)
		}, code: "DRAWING_METADATA_PRESERVED"},
		{name: "hidden inner picture remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<pic:cNvPr id="0"`, `<pic:cNvPr hidden="1" id="0"`, 1)
		}, code: "PICTURE_NONVISUAL_PRESERVED"},
		{name: "malformed picture lock remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<pic:cNvPicPr/>`, `<pic:cNvPicPr><a:picLocks noCrop="on"/></pic:cNvPicPr>`, 1)
		}, code: "PICTURE_NONVISUAL_PRESERVED"},
		{name: "unknown compression state remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<a:blip r:embed="rImage"/>`, `<a:blip r:embed="rImage" cstate="lossless-ish"/>`, 1)
		}, code: "PICTURE_EFFECTS_PRESERVED"},
		{name: "unmodeled inline geometry remains preserve-only", mutate: func(parts map[string]string) {
			parts["Custom/Main.XML"] = strings.Replace(parts["Custom/Main.XML"], `<wp:extent`, `<wp:sizeRelH/><wp:extent`, 1)
		}, code: "INLINE_DRAWING_SEMANTICS_PRESERVED"},
	} {
		t.Run(test.name, func(t *testing.T) {
			parts := transitionalNativeParts()
			test.mutate(parts)
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			for _, run := range doc.Body.Blocks[0].Paragraph.Runs {
				if run.Drawing != nil {
					t.Fatal("non-identity picture transform must not produce a partially authoritative drawing")
				}
			}
			if !hasUnsupportedCode(doc, test.code) || !hasNativePassthrough(doc, "Custom/Media/image.PNG") {
				t.Fatalf("transform refusal/passthrough missing: unsupported=%#v", doc.Unsupported)
			}
		})
	}

	t.Run("duplicate image relationship id refuses the package", func(t *testing.T) {
		parts := transitionalNativeParts()
		parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], `</Relationships>`, `<Relationship Id="rImage" Type="`+relBaseTransitional+`image" Target="media/image.png"/></Relationships>`, 1)
		if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "duplicate relationship id") {
			t.Fatalf("duplicate image relationship must fail closed, err=%v", err)
		}
	})

	t.Run("image relationship traversal refuses the package", func(t *testing.T) {
		parts := transitionalNativeParts()
		parts["Custom/_RELS/Main.XML.RELS"] = strings.Replace(parts["Custom/_RELS/Main.XML.RELS"], `Target="media/image.png"`, `Target="../../../escape.png"`, 1)
		if _, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts))); err == nil || !strings.Contains(err.Error(), "escapes the package root") {
			t.Fatalf("traversing image relationship must fail closed, err=%v", err)
		}
	})
}

func TestExtractNativeDocumentInfersSchemaOptionalDefaultSection(t *testing.T) {
	parts := map[string]string{
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="` + relBaseTransitional + `officeDocument" Target="word/document.xml"/></Relationships>`,
		"word/document.xml":   `<w:document xmlns:w="` + testW + `"><w:body><w:p><w:r><w:t>Default page</w:t></w:r></w:p></w:body></w:document>`,
	}
	doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Sections) != 1 || *doc.Sections[0].Page.WidthTwips != 12240 || doc.Sections[0].Anchor.XMLSHA256 != doc.Body.Anchor.XMLSHA256 || !hasUnsupportedCode(doc, "DEFAULT_SECTION_INFERRED") {
		t.Fatalf("default section inference = %#v unsupported=%#v", doc.Sections, doc.Unsupported)
	}
	if !hasNativePassthrough(doc, "_rels/.rels") {
		t.Fatal("root relationships must remain preserve-verbatim")
	}
}

func TestNativeXMLAnchorStartsAtElementAndRejectsTrailingGarbage(t *testing.T) {
	root, err := parseNativeXML("x.xml", []byte(" \n<r xmlns=\"urn:x\"><a/></r> \n"))
	if err != nil {
		t.Fatal(err)
	}
	if root.Start != 2 || root.End <= root.Start {
		t.Fatalf("unexpected exact root extent: %d..%d", root.Start, root.End)
	}
	if _, err := parseNativeXML("x.xml", []byte("<r/>garbage")); err == nil {
		t.Fatal("expected trailing XML garbage rejection")
	}
	if _, err := parseNativeXML("x.xml", []byte("\u00a0<r/>")); err == nil {
		t.Fatal("Unicode whitespace outside the root is not XML S and must be rejected")
	}
}

func TestExtractNativeDocumentPaginationSingletonsFailClosedInBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS, relBase := "Transitional", wordMLTransitional, relBaseTransitional
		if strict {
			name, wordNS, relBase = "Strict", wordMLStrict, relBaseStrict
		}
		t.Run(name, func(t *testing.T) {
			parts := map[string]string{
				"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
				"_rels/.rels":         `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBase + `officeDocument" Target="word/document.xml"/></Relationships>`,
				"word/document.xml":   `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr><w:keepNext/><w:keepNext/><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgSz w:w="11906" w:h="16838"/></w:sectPr></w:pPr><w:r><w:t>ambiguous</w:t></w:r></w:p></w:body></w:document>`,
			}
			doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if !hasUnsupportedCode(doc, "DUPLICATE_PARAGRAPH_PROPERTY") || !hasUnsupportedCode(doc, "DUPLICATE_SECTION_PROPERTY") {
				t.Fatalf("pagination singleton ambiguity was not refused: %#v", doc.Unsupported)
			}

			parts["word/document.xml"] = `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr><w:keepNext bogus="1"/><w:sectPr><w:pgSz w:w="12240" w:h="15840" bogus="1"/></w:sectPr></w:pPr><w:r><w:t>opaque</w:t></w:r></w:p></w:body></w:document>`
			doc, err = ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if !hasUnsupportedCode(doc, "PARTIAL_PARAGRAPH_PROPERTIES") || !hasUnsupportedCode(doc, "UNMODELED_SECTION_PROPERTY") {
				t.Fatalf("unmodeled pagination attributes were not refused: %#v", doc.Unsupported)
			}

			for _, test := range []struct {
				name      string
				pPrAttrs  string
				pPrText   string
				sectAttrs string
				sectText  string
			}{
				{name: "paragraph container attribute", pPrAttrs: ` bogus="1"`},
				{name: "paragraph container non XML S", pPrText: `&#xA0;`},
				{name: "section container attribute", sectAttrs: ` bogus="1"`},
				{name: "section container non XML S", sectText: `&#xA0;`},
			} {
				t.Run(test.name, func(t *testing.T) {
					parts["word/document.xml"] = `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr` + test.pPrAttrs + `>` + test.pPrText + `<w:keepNext/><w:sectPr` + test.sectAttrs + `>` + test.sectText + `<w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>opaque container</w:t></w:r></w:p></w:body></w:document>`
					doc, err := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
					if err != nil {
						t.Fatal(err)
					}
					if (test.pPrAttrs != "" || test.pPrText != "") && !hasUnsupportedCode(doc, "PARTIAL_PARAGRAPH_PROPERTIES") {
						t.Fatalf("pPr container smuggling was not refused: %#v", doc.Unsupported)
					}
					if (test.sectAttrs != "" || test.sectText != "") && !hasUnsupportedCode(doc, "UNMODELED_SECTION_PROPERTY") {
						t.Fatalf("sectPr container smuggling was not refused: %#v", doc.Unsupported)
					}
				})
			}

			parts["word/document.xml"] = `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr w:rsidRPr="00AA11BB"><w:keepNext/><w:sectPr w:rsidRPr="00AA11BB" w:rsidDel="00000000" w:rsidR="00CC22DD" w:rsidSect="00EE33FF"><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>known revision metadata</w:t></w:r></w:p></w:body></w:document>`
			doc, err = ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
			if err != nil {
				t.Fatal(err)
			}
			if hasUnsupportedCode(doc, "PARTIAL_PARAGRAPH_PROPERTIES") || hasUnsupportedCode(doc, "UNMODELED_SECTION_PROPERTY") {
				t.Fatalf("exact layout-neutral revision attributes must remain accepted: %#v", doc.Unsupported)
			}
		})
	}
}

func TestExtractNativeDocumentAcceptsOnlyExactResolvedParagraphGeometryInBothDialects(t *testing.T) {
	for _, strict := range []bool{false, true} {
		name, wordNS, relBase := "Transitional", wordMLTransitional, relBaseTransitional
		if strict {
			name, wordNS, relBase = "Strict", wordMLStrict, relBaseStrict
		}
		t.Run(name, func(t *testing.T) {
			parts := map[string]string{
				"[Content_Types].xml": `<Types xmlns="` + opcContentTypesNS + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
				"_rels/.rels":         `<Relationships xmlns="` + opcRelationshipsNS + `"><Relationship Id="office" Type="` + relBase + `officeDocument" Target="word/document.xml"/></Relationships>`,
				"word/document.xml":   `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="atLeast"/><w:ind w:start="720" w:end="120" w:hanging="180"/><w:bidi w:val="false"/><w:keepNext/><w:widowControl/></w:pPr><w:r><w:t>exact direct geometry</w:t><w:br w:clear="none"/><w:t>line two</w:t><w:cr/></w:r></w:p><w:sectPr/></w:body></w:document>`,
			}
			data := buildNativeDOCX(t, nativeEntries(parts))
			doc, err := ExtractNativeDocumentV1(data)
			if err != nil {
				t.Fatal(err)
			}
			if hasUnsupportedCode(doc, "PARTIAL_PARAGRAPH_PROPERTIES") || hasUnsupportedCode(doc, "UNMODELED_BREAK") || hasUnsupportedCode(doc, "UNMODELED_CONTROL") {
				t.Fatalf("exact pagination prerequisites were self-refused: %#v", doc.Unsupported)
			}
			if policy := doc.Body.Blocks[0].Paragraph.EditPolicy; policy.Mode != "read-only" || policy.Refusal == nil {
				t.Fatalf("exact resolver-owned geometry must not expand extractor mutation authority: %#v", policy)
			}
			resolved, err := ResolveNativeDocumentLayoutV1(data)
			if err != nil {
				t.Fatal(err)
			}
			properties := resolved.Paragraphs[0].Properties
			if properties.SpacingBeforeTwips == nil || *properties.SpacingBeforeTwips != 120 || properties.SpacingAfterTwips == nil || *properties.SpacingAfterTwips != 240 || properties.Line == nil || *properties.Line != 360 || properties.LineRule == nil || *properties.LineRule != "atLeast" || properties.IndentStartTwips == nil || *properties.IndentStartTwips != 720 || properties.IndentEndTwips == nil || *properties.IndentEndTwips != 120 || properties.HangingTwips == nil || *properties.HangingTwips != 180 || properties.KeepNext == nil || !*properties.KeepNext || properties.WidowControl == nil || !*properties.WidowControl {
				t.Fatalf("exact direct paragraph geometry did not reach resolved layout: %#v", properties)
			}

			for _, hostile := range []string{
				`<w:spacing w:before="120" bogus="1"/>`,
				`<w:spacing w:before="120"><w:keepNext/></w:spacing>`,
				`<w:ind w:start="720" bogus="1"/>`,
				`<w:bidi bogus="1"/>`,
			} {
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:pPr>` + hostile + `</w:pPr><w:r><w:t>smuggled</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
				hostileDoc, extractErr := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
				if extractErr != nil {
					t.Fatal(extractErr)
				}
				if !hasUnsupportedCode(hostileDoc, "PARTIAL_PARAGRAPH_PROPERTIES") {
					t.Fatalf("pagination-sensitive paragraph smuggling was accepted: %s %#v", hostile, hostileDoc.Unsupported)
				}
			}

			for _, hostileBreak := range []string{`<w:br bogus="1"/>`, `<w:br w:clear="all"/>`, `<w:cr bogus="1"/>`, `<w:tab bogus="1"/>`, `<w:softHyphen bogus="1"/>`} {
				parts["word/document.xml"] = `<w:document xmlns:w="` + wordNS + `"><w:body><w:p><w:r><w:t>before</w:t>` + hostileBreak + `<w:t>after</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`
				hostileDoc, extractErr := ExtractNativeDocumentV1(buildNativeDOCX(t, nativeEntries(parts)))
				if extractErr != nil {
					t.Fatal(extractErr)
				}
				if !hasUnsupportedCode(hostileDoc, "UNMODELED_BREAK") && !hasUnsupportedCode(hostileDoc, "BREAK_CLEAR_UNSUPPORTED") && !hasUnsupportedCode(hostileDoc, "UNMODELED_CONTROL") {
					t.Fatalf("hostile break/control was accepted: %s %#v", hostileBreak, hostileDoc.Unsupported)
				}
			}
		})
	}
}

func readNativeZipPart(t *testing.T, data []byte, name string) []byte {
	t.Helper()
	zr, err := zip.NewReader(bytes.NewReader(data), int64(len(data)))
	if err != nil {
		t.Fatal(err)
	}
	for _, file := range zr.File {
		if file.Name != name {
			continue
		}
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			t.Fatal(err)
		}
		return content
	}
	t.Fatalf("missing %s", name)
	return nil
}
