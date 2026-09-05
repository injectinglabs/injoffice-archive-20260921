// Command playground_sample generates the populated PPTX used by the
// browser-native playground proof. Its narrow OOXML mirrors the exact native
// v1 projection rather than relying on an application template.
package main

import (
	"archive/zip"
	"bytes"
	"fmt"
	"os"
	"strings"
	"time"
)

const (
	presentationNS = "http://schemas.openxmlformats.org/presentationml/2006/main"
	drawingNS      = "http://schemas.openxmlformats.org/drawingml/2006/main"
	officeRelNS    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
	packageRelNS   = "http://schemas.openxmlformats.org/package/2006/relationships"
)

type textElement struct {
	id, name, text, color string
	x, y, cx, cy          int
	size                  int
	bold                  bool
}

type shapeElement struct {
	id, name, preset, fill string
	x, y, cx, cy           int
}

type slide struct {
	texts  []textElement
	shapes []shapeElement
}

func boolValue(value bool) string {
	if value {
		return "1"
	}
	return "0"
}

func xmlEscape(value string) string {
	return strings.NewReplacer("&", "&amp;", "<", "&lt;", ">", "&gt;", `"`, "&quot;", "'", "&apos;").Replace(value)
}

func textXML(value textElement) string {
	return fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="l" lvl="0"><a:buNone/></a:pPr><a:r><a:rPr b="%s" i="0" sz="%d"><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:latin typeface="Aptos"/></a:rPr><a:t>%s</a:t></a:r></a:p></p:txBody></p:sp>`,
		value.id, xmlEscape(value.name), value.x, value.y, value.cx, value.cy,
		boolValue(value.bold), value.size, value.color, xmlEscape(value.text))
}

func shapeXML(value shapeElement) string {
	return fmt.Sprintf(`<p:sp><p:nvSpPr><p:cNvPr id="%s" name="%s"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="%d" y="%d"/><a:ext cx="%d" cy="%d"/></a:xfrm><a:prstGeom prst="%s"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="%s"/></a:solidFill><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:srgbClr val="29265F"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln></p:spPr></p:sp>`,
		value.id, xmlEscape(value.name), value.x, value.y, value.cx, value.cy,
		value.preset, value.fill)
}

func slideXML(value slide) string {
	var content strings.Builder
	for _, shape := range value.shapes {
		content.WriteString(shapeXML(shape))
	}
	for _, text := range value.texts {
		content.WriteString(textXML(text))
	}
	return `<p:sld xmlns:p="` + presentationNS + `" xmlns:a="` + drawingNS + `" xmlns:r="` + officeRelNS + `"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` + content.String() + `</p:spTree></p:cSld></p:sld>`
}

func fixtureSlides() []slide {
	return []slide{
		{
			shapes: []shapeElement{
				{id: "10", name: "Activation card", preset: "rect", fill: "EEF0FF", x: 700000, y: 2700000, cx: 3200000, cy: 1900000},
				{id: "11", name: "Time to value card", preset: "rect", fill: "FFF1ED", x: 4490000, y: 2700000, cx: 3200000, cy: 1900000},
				{id: "12", name: "Pilot conversion card", preset: "rect", fill: "ECF8F3", x: 8280000, y: 2700000, cx: 3200000, cy: 1900000},
			},
			texts: []textElement{
				{id: "2", name: "Review title", text: "Northstar launch review", color: "171A2B", x: 700000, y: 500000, cx: 10800000, cy: 700000, size: 3400, bold: true},
				{id: "3", name: "Review subtitle", text: "Q3 operating brief · September 2026", color: "626779", x: 700000, y: 1400000, cx: 10800000, cy: 400000, size: 1500},
				{id: "4", name: "Activation value", text: "128%", color: "29265F", x: 950000, y: 3000000, cx: 2500000, cy: 500000, size: 2700, bold: true},
				{id: "5", name: "Activation label", text: "of quarterly target", color: "626779", x: 950000, y: 3800000, cx: 2500000, cy: 350000, size: 1200},
				{id: "6", name: "Time to value", text: "7 days · down from 11", color: "29265F", x: 4740000, y: 3200000, cx: 2700000, cy: 500000, size: 1900, bold: true},
				{id: "7", name: "Pilot conversion", text: "64% · enterprise cohort", color: "29265F", x: 8530000, y: 3200000, cx: 2700000, cy: 500000, size: 1900, bold: true},
				{id: "8", name: "Recommendation", text: "Approve the October rollout with weekly checkpoints.", color: "E05046", x: 700000, y: 5400000, cx: 10800000, cy: 600000, size: 1900, bold: true},
			},
		},
		{
			shapes: []shapeElement{{id: "8", name: "Reliability checkpoint", preset: "ellipse", fill: "F47B70", x: 10100000, y: 5100000, cx: 700000, cy: 700000}},
			texts: []textElement{
				{id: "2", name: "Plan title", text: "The next 30 days", color: "171A2B", x: 700000, y: 500000, cx: 10800000, cy: 700000, size: 3200, bold: true},
				{id: "3", name: "Trust gate", text: "01 · Close trust gaps", color: "E05046", x: 900000, y: 1900000, cx: 9500000, cy: 500000, size: 2000, bold: true},
				{id: "4", name: "Trust detail", text: "Audit-log export and accessibility findings remain release gates.", color: "626779", x: 900000, y: 2500000, cx: 9500000, cy: 500000, size: 1500},
				{id: "5", name: "Adoption gate", text: "02 · Scale adoption", color: "625BF6", x: 900000, y: 3600000, cx: 9500000, cy: 500000, size: 2000, bold: true},
				{id: "6", name: "Adoption detail", text: "Publish migration playbooks and name an owner for every pilot.", color: "626779", x: 900000, y: 4200000, cx: 9500000, cy: 500000, size: 1500},
			},
		},
		{
			shapes: []shapeElement{{id: "6", name: "Decision marker", preset: "diamond", fill: "625BF6", x: 5200000, y: 4000000, cx: 1600000, cy: 1600000}},
			texts: []textElement{
				{id: "2", name: "Decision eyebrow", text: "DECISION", color: "E05046", x: 700000, y: 700000, cx: 10800000, cy: 350000, size: 1200, bold: true},
				{id: "3", name: "Decision title", text: "Approve the staged rollout", color: "171A2B", x: 700000, y: 1450000, cx: 10800000, cy: 700000, size: 3400, bold: true},
				{id: "4", name: "Decision detail", text: "Proceed with weekly adoption and reliability checkpoints.", color: "626779", x: 700000, y: 2500000, cx: 10800000, cy: 500000, size: 1600},
			},
		},
	}
}

func packageParts() map[string]string {
	parts := map[string]string{
		"[Content_Types].xml":                          `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/><Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/><Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/></Types>`,
		"_rels/.rels":                                  `<Relationships xmlns="` + packageRelNS + `"><Relationship Id="rIdRoot" Type="` + officeRelNS + `/officeDocument" Target="ppt/presentation.xml"/></Relationships>`,
		"ppt/presentation.xml":                         `<p:presentation xmlns:p="` + presentationNS + `" xmlns:r="` + officeRelNS + `"><p:sldIdLst><p:sldId id="256" r:id="rId1"/><p:sldId id="257" r:id="rId2"/><p:sldId id="258" r:id="rId3"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>`,
		"ppt/_rels/presentation.xml.rels":              `<Relationships xmlns="` + packageRelNS + `"><Relationship Id="rId1" Type="` + officeRelNS + `/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="` + officeRelNS + `/slide" Target="slides/slide2.xml"/><Relationship Id="rId3" Type="` + officeRelNS + `/slide" Target="slides/slide3.xml"/></Relationships>`,
		"ppt/slideLayouts/slideLayout1.xml":            `<p:sldLayout xmlns:p="` + presentationNS + `"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>`,
		"ppt/slideLayouts/_rels/slideLayout1.xml.rels": `<Relationships xmlns="` + packageRelNS + `"><Relationship Id="rIdMaster" Type="` + officeRelNS + `/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`,
		"ppt/slideMasters/slideMaster1.xml":            `<p:sldMaster xmlns:p="` + presentationNS + `"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldMaster>`,
		"ppt/slideMasters/_rels/slideMaster1.xml.rels": `<Relationships xmlns="` + packageRelNS + `"><Relationship Id="rIdTheme" Type="` + officeRelNS + `/theme" Target="../theme/theme1.xml"/></Relationships>`,
		"ppt/theme/theme1.xml":                         `<a:theme xmlns:a="` + drawingNS + `" name="InjOffice fixture"><a:themeElements/></a:theme>`,
	}
	for index, value := range fixtureSlides() {
		parts[fmt.Sprintf("ppt/slides/slide%d.xml", index+1)] = slideXML(value)
		parts[fmt.Sprintf("ppt/slides/_rels/slide%d.xml.rels", index+1)] = `<Relationships xmlns="` + packageRelNS + `"><Relationship Id="rIdLayout" Type="` + officeRelNS + `/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>`
	}
	return parts
}

func buildPackage() ([]byte, error) {
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	fixed := time.Date(1980, time.January, 1, 0, 0, 0, 0, time.UTC)
	parts := packageParts()
	names := []string{
		"[Content_Types].xml", "_rels/.rels", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels",
		"ppt/slides/slide1.xml", "ppt/slides/_rels/slide1.xml.rels", "ppt/slides/slide2.xml", "ppt/slides/_rels/slide2.xml.rels", "ppt/slides/slide3.xml", "ppt/slides/_rels/slide3.xml.rels",
		"ppt/slideLayouts/slideLayout1.xml", "ppt/slideLayouts/_rels/slideLayout1.xml.rels", "ppt/slideMasters/slideMaster1.xml", "ppt/slideMasters/_rels/slideMaster1.xml.rels", "ppt/theme/theme1.xml",
	}
	for _, name := range names {
		part, err := writer.CreateHeader(&zip.FileHeader{Name: name, Method: zip.Store, Modified: fixed})
		if err != nil {
			return nil, err
		}
		if _, err := part.Write([]byte(parts[name])); err != nil {
			return nil, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}

func main() {
	out := "testdata/playground_northstar_review.pptx"
	if len(os.Args) > 1 {
		out = os.Args[1]
	}
	data, err := buildPackage()
	if err != nil {
		fmt.Fprintln(os.Stderr, "pptxpatch: build playground sample:", err)
		os.Exit(1)
	}
	if err := os.WriteFile(out, data, 0o644); err != nil {
		fmt.Fprintln(os.Stderr, "pptxpatch: write playground sample:", err)
		os.Exit(1)
	}
	fmt.Println("wrote", out)
}
