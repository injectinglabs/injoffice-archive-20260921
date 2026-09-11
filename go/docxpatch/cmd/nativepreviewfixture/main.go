// nativepreviewfixture builds a deterministic test document from a caller-owned
// embeddable DejaVu Sans font. No font bytes are checked into the repository.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/hex"
	"flag"
	"fmt"
	"image"
	"image/color"
	"image/jpeg"
	"os"
	"sort"
	"strings"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

const wns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
const rns = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const pns = "http://schemas.openxmlformats.org/package/2006/relationships"
const fontKey = "{00112233-4455-6677-8899-AABBCCDDEEFF}"

func build(font []byte) ([]byte, error) {
	return buildWithJPEG(font, false)
}

func buildWithJPEG(font []byte, withJPEG bool) ([]byte, error) {
	return buildWithUnderline(font, withJPEG, "")
}

func buildWithUnderline(font []byte, withJPEG bool, underline string) ([]byte, error) {
	return buildWithDecorations(font, withJPEG, false, underline)
}

func buildWithFields(font []byte, withJPEG, withFields bool) ([]byte, error) {
	return buildWithDecorations(font, withJPEG, withFields, "")
}

func buildWithDecorations(font []byte, withJPEG, withFields bool, underline string) ([]byte, error) {
	return buildWithAllDecorations(font, withJPEG, withFields, false, underline)
}

func buildWithScripts(font []byte, withJPEG, withFields, withScripts bool) ([]byte, error) {
	return buildWithAllDecorations(font, withJPEG, withFields, withScripts, "")
}

func buildWithAllDecorations(font []byte, withJPEG, withFields, withScripts bool, underline string) ([]byte, error) {
	return buildWithAllRendering(font, withJPEG, withFields, withScripts, underline, false)
}

func buildWithFloating(font []byte, withJPEG, withFields bool, underline string, withFloating bool) ([]byte, error) {
	return buildWithAllRendering(font, withJPEG, withFields, false, underline, withFloating)
}

func buildWithAllRendering(font []byte, withJPEG, withFields, withScripts bool, underline string, withFloating bool) ([]byte, error) {
	if underline != "" && underline != "single" && underline != "double" && underline != "words" {
		return nil, fmt.Errorf("unsupported underline style")
	}
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
		highlight := ""
		if withJPEG && text == "Native document preview" {
			highlight = `<w:highlight w:val="yellow"/>`
		}
		if underline != "" && text == "Native document preview" {
			highlight += `<w:u w:val="` + underline + `"/>`
		}
		return `<w:p><w:pPr>` + br + `<w:jc w:val="left"/><w:spacing w:before="0" w:after="120"/><w:rPr>` + runProps + `</w:rPr></w:pPr><w:r><w:rPr>` + runProps + highlight + `</w:rPr><w:t>` + text + `</w:t></w:r></w:p>`
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
	parts["word/document.xml"] = []byte(strings.Replace(string(parts["word/document.xml"]), `<w:body>`, `<w:body><w:p><w:pPr><w:rPr>`+runProps+`</w:rPr></w:pPr></w:p>`, 1))
	if withJPEG {
		pixels := image.NewRGBA(image.Rect(0, 0, 16, 8))
		for y := 0; y < 8; y++ {
			for x := 0; x < 16; x++ {
				if x < 8 {
					pixels.Set(x, y, color.RGBA{220, 40, 40, 255})
				} else {
					pixels.Set(x, y, color.RGBA{30, 80, 220, 255})
				}
			}
		}
		var encoded bytes.Buffer
		if err := jpeg.Encode(&encoded, pixels, &jpeg.Options{Quality: 95}); err != nil {
			return nil, err
		}
		// Go emits baseline JPEG without APP0. Attest the JFIF YCbCr convention
		// explicitly; no external image, EXIF, ICC or orientation metadata.
		jfif := []byte{0xff, 0xe0, 0, 16, 'J', 'F', 'I', 'F', 0, 1, 2, 0, 0, 1, 0, 1, 0, 0}
		media := append([]byte{0xff, 0xd8}, jfif...)
		media = append(media, encoded.Bytes()[2:]...)
		parts["word/media/bands.jpg"] = media
		parts["[Content_Types].xml"] = []byte(strings.Replace(string(parts["[Content_Types].xml"]), "</Types>", `<Default Extension="jpg" ContentType="image/jpeg"/></Types>`, 1))
		parts["word/_rels/document.xml.rels"] = []byte(strings.Replace(string(parts["word/_rels/document.xml.rels"]), "</Relationships>", `<Relationship Id="rImage" Type="`+rns+`/image" Target="media/bands.jpg"/></Relationships>`, 1))
		drawing := `<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="1828800" cy="457200"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="1" name="JPEG bands" descr="Red and blue JPEG bands"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="JPEG bands"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm/><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`
		if withFloating {
			floating := strings.Replace(drawing, `<wp:inline distT="0" distB="0" distL="0" distR="0">`, `<wp:anchor simplePos="0" relativeHeight="7" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" distT="0" distB="0" distL="0" distR="0"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="page"><wp:posOffset>3657600</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>2743200</wp:posOffset></wp:positionV><wp:wrapNone/>`, 1)
			floating = strings.Replace(floating, `</wp:inline>`, `</wp:anchor>`, 1)
			floating = strings.Replace(floating, `docPr id="1"`, `docPr id="2"`, 1)
			// Keep an inline picture as a distinct flow-layout regression.
			secondPage := paragraph("Second page", true)
			parts["word/document.xml"] = []byte(strings.Replace(string(parts["word/document.xml"]), secondPage, secondPage+floating, 1))
		}
		document := strings.Replace(string(parts["word/document.xml"]), `<w:body>`, `<w:body>`+drawing, 1)
		document = strings.Replace(document, `<w:document `, `<w:document xmlns:r="`+rns+`" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" `, 1)
		parts["word/document.xml"] = []byte(document)
	}
	if withScripts {
		run := func(text, alignment string) string {
			extra := ""
			if alignment != "" {
				extra = `<w:vertAlign w:val="` + alignment + `"/>`
			}
			return `<w:r><w:rPr>` + runProps + extra + `</w:rPr><w:t xml:space="preserve">` + text + `</w:t></w:r>`
		}
		formula := `<w:p><w:pPr><w:spacing w:before="0" w:after="120"/></w:pPr>` + run("H", "") + run("2", "subscript") + run("O + x", "") + run("2", "superscript") + `</w:p>`
		parts["word/document.xml"] = []byte(strings.Replace(string(parts["word/document.xml"]), paragraph("Second page", true), formula+paragraph("Second page", true), 1))
	}
	if withFields {
		field := func(instruction string) string {
			return `<w:fldSimple w:instr=" ` + instruction + ` "><w:r><w:rPr>` + runProps + `</w:rPr><w:t>999</w:t></w:r></w:fldSimple>`
		}
		for _, region := range []string{"header", "footer"} {
			tag := "hdr"
			if region == "footer" {
				tag = "ftr"
			}
			parts["word/"+region+"1.xml"] = []byte(`<w:` + tag + ` xmlns:w="` + wns + `"><w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t xml:space="preserve">Page </w:t></w:r>` + field("PAGE") + `<w:r><w:t xml:space="preserve"> of </w:t></w:r>` + field("NUMPAGES") + `</w:p></w:` + tag + `>`)
			parts["[Content_Types].xml"] = []byte(strings.Replace(string(parts["[Content_Types].xml"]), "</Types>", `<Override PartName="/word/`+region+`1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.`+region+`+xml"/></Types>`, 1))
			parts["word/_rels/document.xml.rels"] = []byte(strings.Replace(string(parts["word/_rels/document.xml.rels"]), "</Relationships>", `<Relationship Id="`+region+`" Type="`+rns+`/`+region+`" Target="`+region+`1.xml"/></Relationships>`, 1))
			parts["word/document.xml"] = []byte(strings.Replace(string(parts["word/document.xml"]), `<w:sectPr>`, `<w:sectPr><w:`+region+`Reference xmlns:r="`+rns+`" w:type="default" r:id="`+region+`"/>`, 1))
		}
	}
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
	withJPEG := flag.Bool("jpeg", false, "include generated baseline JFIF JPEG bands")
	withFloating := flag.Bool("floating", false, "include page-relative no-wrap foreground JPEG anchor (requires -jpeg)")
	underline := flag.String("underline", "", "title underline: single, double, or words")
	withFields := flag.Bool("page-fields", false, "include source-bound PAGE/NUMPAGES with stale caches in header and footer")
	withScripts := flag.Bool("scripts", false, "include H2O + x2 using font-metric subscript and superscript")
	flag.Parse()
	if *fontPath == "" || *out == "" {
		fmt.Fprintln(os.Stderr, "-font and -out are required")
		os.Exit(2)
	}
	font, err := os.ReadFile(*fontPath)
	if err == nil {
		var data []byte
		data, err = buildWithAllRendering(font, *withJPEG, *withFields, *withScripts, *underline, *withFloating)
		if err == nil {
			err = os.WriteFile(*out, data, 0600)
		}
	}
	if err != nil {
		fmt.Fprintln(os.Stderr, err)
		os.Exit(1)
	}
}
