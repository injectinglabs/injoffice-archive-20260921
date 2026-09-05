// Command nativeplaygroundfixture writes the deterministic, repository-owned
// DOCX artifact used by the InjOffice Documents playground.
package main

import (
	"archive/zip"
	"bytes"
	"encoding/base64"
	"fmt"
	"os"
	"path/filepath"
	"sort"
)

const (
	wordNS  = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
	word14  = "http://schemas.microsoft.com/office/word/2010/wordml"
	pkgRels = "http://schemas.openxmlformats.org/package/2006/relationships"
	office  = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
	content = "http://schemas.openxmlformats.org/package/2006/content-types"
)

func main() {
	if err := run(); err != nil {
		fmt.Fprintln(os.Stderr, "nativeplaygroundfixture:", err)
		os.Exit(1)
	}
}

func run() error {
	root, err := repoRoot()
	if err != nil {
		return err
	}
	outDir := filepath.Join(root, "apps", "playground", "public", "native-docx")
	if len(os.Args) > 1 {
		outDir = os.Args[1]
	}
	if err := os.MkdirAll(outDir, 0o755); err != nil {
		return err
	}
	fixture, err := buildPlaygroundFixture()
	if err != nil {
		return err
	}
	encoded := append([]byte(base64.StdEncoding.EncodeToString(fixture)), '\n')
	path := filepath.Join(outDir, "northstar-launch-brief.docx.b64")
	if err := os.WriteFile(path, encoded, 0o644); err != nil {
		return err
	}
	fmt.Fprintf(os.Stderr, "wrote %s\n", path)
	return nil
}

func repoRoot() (string, error) {
	dir, err := os.Getwd()
	if err != nil {
		return "", err
	}
	for {
		if _, err := os.Stat(filepath.Join(dir, "go", "docxpatch")); err == nil {
			if _, err := os.Stat(filepath.Join(dir, "apps", "playground")); err == nil {
				return dir, nil
			}
		}
		parent := filepath.Dir(dir)
		if parent == dir {
			return "", fmt.Errorf("run from the injoffice repository")
		}
		dir = parent
	}
}

func buildPlaygroundFixture() ([]byte, error) {
	return buildZip(map[string][]byte{
		"[Content_Types].xml": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Types xmlns="` + content + `"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
			`<Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`),
		"_rels/.rels": []byte(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
			`<Relationships xmlns="` + pkgRels + `"><Relationship Id="rOffice" Type="` + office + `" Target="word/document.xml"/></Relationships>`),
		"word/document.xml": []byte(documentXML()),
	})
}

func documentXML() string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
		`<w:document xmlns:w="` + wordNS + `" xmlns:w14="` + word14 + `"><w:body>` +
		paragraph("10000001", `<w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr>`, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:b/><w:color w:val="234F78"/><w:sz w:val="36"/></w:rPr><w:t>Northstar Launch Brief</w:t></w:r>`) +
		paragraph("10000002", `<w:pPr><w:jc w:val="center"/></w:pPr>`, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:i/><w:color w:val="5C6B7A"/></w:rPr><w:t>Decision memo | 14 October 2026</w:t></w:r>`) +
		paragraph("10000003", ``, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/></w:rPr><w:t>Northstar is prepared for a controlled beta launch. The team has validated the core workflow, assigned every open risk, and aligned customer support on the rollout plan.</w:t></w:r>`) +
		paragraph("10000004", `<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>`, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:b/><w:color w:val="234F78"/><w:sz w:val="26"/></w:rPr><w:t>Executive snapshot</w:t></w:r>`) +
		decisionTable() +
		paragraph("10000011", `<w:pPr><w:pStyle w:val="Heading2"/></w:pPr>`, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:b/><w:color w:val="234F78"/><w:sz w:val="26"/></w:rPr><w:t>This week's priorities</w:t></w:r>`) +
		priority("10000012", "1 · Close the loop", "Finish accessibility checks and publish the signed release checklist.") +
		priority("10000013", "2 · Rehearse the handoff", "Run the support escalation drill with product, engineering, and success leads.") +
		priority("10000014", "3 · Protect the signal", "Limit beta scope to the 18 confirmed design partners and review telemetry daily.") +
		paragraph("10000015", `<w:pPr><w:jc w:val="center"/></w:pPr>`, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:i/><w:color w:val="3B6D5D"/></w:rPr><w:t>Next checkpoint: launch council, Thursday at 09:30.</w:t></w:r>`) +
		`<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>` +
		`</w:body></w:document>`
}

func paragraph(id, properties, runs string) string {
	return `<w:p w14:paraId="` + id + `">` + properties + runs + `</w:p>`
}

func priority(id, label, detail string) string {
	return paragraph(id, ``, `<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:b/><w:color w:val="234F78"/></w:rPr><w:t>`+label+` — </w:t></w:r>`+
		`<w:r><w:rPr><w:rFonts w:ascii="DejaVu Sans"/></w:rPr><w:t>`+detail+`</w:t></w:r>`)
}

func decisionTable() string {
	return `<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblBorders>` +
		`<w:top w:val="single" w:sz="6" w:color="D7DEE8"/><w:left w:val="single" w:sz="6" w:color="D7DEE8"/><w:bottom w:val="single" w:sz="6" w:color="D7DEE8"/><w:right w:val="single" w:sz="6" w:color="D7DEE8"/><w:insideH w:val="single" w:sz="6" w:color="D7DEE8"/><w:insideV w:val="single" w:sz="6" w:color="D7DEE8"/>` +
		`</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="2800"/><w:gridCol w:w="2200"/><w:gridCol w:w="4360"/></w:tblGrid>` +
		tableRow(true, "10000005", "Decision", "10000006", "Owner", "10000007", "Evidence") +
		tableRow(false, "10000008", "Release scope", "10000009", "Maya Chen", "10000010", "Approved - 24 workflows validated") +
		tableRow(false, "10000016", "Risk posture", "10000017", "Omar Reed", "10000018", "2 medium items - owners assigned") +
		tableRow(false, "10000019", "Customer readiness", "10000020", "Sofia King", "10000021", "18 design partners confirmed") +
		`</w:tbl>`
}

func tableRow(header bool, firstID, first, secondID, second, thirdID, third string) string {
	rowProperties := ""
	shading := ""
	runProperties := `<w:rPr><w:rFonts w:ascii="DejaVu Sans"/></w:rPr>`
	if header {
		rowProperties = `<w:trPr><w:tblHeader/></w:trPr>`
		shading = `<w:shd w:fill="234F78"/>`
		runProperties = `<w:rPr><w:rFonts w:ascii="DejaVu Sans"/><w:b/><w:color w:val="FFFFFF"/></w:rPr>`
	}
	cell := func(width int, id, text string) string {
		return fmt.Sprintf(`<w:tc><w:tcPr><w:tcW w:w="%d" w:type="dxa"/>%s</w:tcPr>%s</w:tc>`, width, shading, paragraph(id, ``, `<w:r>`+runProperties+`<w:t>`+text+`</w:t></w:r>`))
	}
	return `<w:tr>` + rowProperties + cell(2800, firstID, first) + cell(2200, secondID, second) + cell(4360, thirdID, third) + `</w:tr>`
}

func buildZip(entries map[string][]byte) ([]byte, error) {
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	names := make([]string, 0, len(entries))
	for name := range entries {
		names = append(names, name)
	}
	sort.Strings(names)
	for _, name := range names {
		header := &zip.FileHeader{Name: name, Method: zip.Deflate}
		part, err := writer.CreateHeader(header)
		if err != nil {
			return nil, err
		}
		if _, err := part.Write(entries[name]); err != nil {
			return nil, err
		}
	}
	if err := writer.Close(); err != nil {
		return nil, err
	}
	return output.Bytes(), nil
}
