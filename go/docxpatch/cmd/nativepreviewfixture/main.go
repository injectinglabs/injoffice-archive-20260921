// nativepreviewfixture builds a deterministic test document from a caller-owned
// embeddable DejaVu Sans font. No font bytes are checked into the repository.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/hex"
	"flag"
	"fmt"
	"os"
	"sort"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

const wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const pns = "http://schemas.openxmlformats.org/package/2006/relationships"
const fontKey = "{00112233-4455-6677-8899-AABBCCDDEEFF}"

func build(font []byte) ([]byte, error) {
	if len(font) < 32 || len(font) > 8*1024*1024 {
		return nil, fmt.Errorf("font must be 32 bytes–8 MiB")
	}
	stored := append([]byte(nil), font...)
	key, _ := hex.DecodeString("00112233445566778899aabbccddeeff")
	for i := 0; i < 32; i++ {
		stored[i] ^= key[15-i%16]
	}
	runProps := `<w:rFonts w:ascii="DejaVu Sans" w:hAnsi="DejaVu Sans"/><w:b w:val="false"/><w:i w:val="false"/><w:color w:val="141816"/><w:sz w:val="24"/><w:lang w:val="en-US"/>`
	paragraph := func(text string, pageBreak bool) string {
		br := ""
		if pageBreak {
			br = `<w:pageBreakBefore/>`
		}
		return `<w:p><w:pPr>` + br + `<w:jc w:val="left"/><w:spacing w:before="0" w:after="120"/></w:pPr><w:r><w:rPr>` + runProps + `</w:rPr><w:t>` + text + `</w:t></w:r></w:p>`
	}
	parts := map[string][]byte{}
	parts["[Content_Types].xml"] = []byte(`<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/fonts/regular.odttf" ContentType="application/vnd.openxmlformats-officedocument.obfuscatedFont"/></Types>`)
	parts["_rels/.rels"] = []byte(`<Relationships xmlns="` + pns + `"><Relationship Id="office" Type="` + rns + `/officeDocument" Target="word/document.xml"/></Relationships>`)
	parts["word/document.xml"] = []byte(`<w:document xmlns:w="` + wns + `"><w:body>` + paragraph("Native document preview", false) + paragraph("This page is shaped from the font embedded in this DOCX. The browser replays glyph paths without measuring text.", false) + paragraph("Second page", true) + paragraph("The explicit page break and source geometry produce a separate native page.", false) + `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`)
	parts["word/styles.xml"] = []byte(`<w:styles xmlns:w="` + wns + `"><w:docDefaults><w:rPrDefault><w:rPr>` + runProps + `</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:jc w:val="left"/><w:spacing w:before="0" w:after="120"/></w:pPr></w:pPrDefault></w:docDefaults></w:styles>`)
	parts["word/settings.xml"] = []byte(`<w:settings xmlns:w="` + wns + `"><w:defaultTabStop w:val="720"/><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat></w:settings>`)
	parts["word/_rels/document.xml.rels"] = []byte(`<Relationships xmlns="` + pns + `"><Relationship Id="fonts" Type="` + rns + `/fontTable" Target="fontTable.xml"/><Relationship Id="styles" Type="` + rns + `/styles" Target="styles.xml"/><Relationship Id="settings" Type="` + rns + `/settings" Target="settings.xml"/></Relationships>`)
	parts["word/fontTable.xml"] = []byte(`<w:fonts xmlns:w="` + wns + `" xmlns:r="` + rns + `"><w:font w:name="DejaVu Sans"><w:embedRegular r:id="regular" w:fontKey="` + fontKey + `" w:subsetted="false"/></w:font></w:fonts>`)
	parts["word/_rels/fontTable.xml.rels"] = []byte(`<Relationships xmlns="` + pns + `"><Relationship Id="regular" Type="` + rns + `/font" Target="fonts/regular.odttf"/></Relationships>`)
	parts["word/fonts/regular.odttf"] = stored
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	names := make([]string, 0, len(parts))
	for name := range parts {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		entry, err := writer.Create(name)
		if err != nil {
			return nil, err
		}
		if _, err = entry.Write(parts[name]); err != nil {
			return nil, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, err
	}
	data := output.Bytes()
	if _, err := docxpatch.ExtractNativeDocumentV1(data); err != nil {
		return nil, err
	}
	if _, err := docxpatch.ResolveNativeDocumentLayoutV1(data); err != nil {
		return nil, err
	}
	if _, err := docxpatch.ExtractNativePaginationSettingsV1(data); err != nil {
		return nil, err
	}
	inventory, err := docxpatch.ExtractNativeDOCXFontInventoryV1(data)
	if err != nil {
		return nil, err
	}
	if _, err = docxpatch.ResolveNativeDOCXPagePaintFontAssetsV1(data, inventory); err != nil {
		return nil, err
	}
	return data, nil
}

func main() {
	fontPath := flag.String("font", "", "path to licensed embeddable DejaVuSans.ttf")
	out := flag.String("out", "", "new temporary DOCX output path")
	flag.Parse()
	if *fontPath == "" || *out == "" {
		fmt.Fprintln(os.Stderr, "-font and -out are required")
		os.Exit(2)
	}
	font, err := os.ReadFile(*fontPath)
	if err == nil {
		var data []byte
		data, err = build(font)
		if err == nil {
			err = os.WriteFile(*out, data, 0600)
		}
	}
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
