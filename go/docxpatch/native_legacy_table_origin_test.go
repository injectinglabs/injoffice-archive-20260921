package docxpatch

import (
	"crypto/sha256"
	"encoding/hex"
	"strings"
	"testing"
)

func legacyOriginParts() map[string]string {
	styles := `<w:styles xmlns:w="` + wordMLTransitional + `"><w:style w:type="table" w:styleId="Base"><w:tblPr><w:tblInd w:type="dxa" w:w="0"/><w:tblCellMar><w:left w:type="dxa" w:w="108"/></w:tblCellMar></w:tblPr></w:style><w:style w:type="table" w:styleId="Child"><w:basedOn w:val="Base"/></w:style></w:styles>`
	parts := resolvedStylesTestParts(styles)
	parts["word/document.xml"] = `<w:document xmlns:w="` + wordMLTransitional + `"><w:body><w:tbl><w:tblPr><w:tblStyle w:val="Child"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:type="dxa" w:w="4000"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/><w:sectPr/></w:body></w:document>`
	return parts
}

func TestLegacyTableOriginSourceEvidence(t *testing.T) {
	for _, strict := range []bool{false, true} {
		for _, name := range []string{"inherited", "direct override", "direct margin", "zero margin", "missing ancestor", "merged", "floating", "spacing", "missing indent", "missing margin", "default margin type", "duplicate indent", "malformed margin", "conditional", "cycle", "cell override", "style cell override", "row override", "rtl", "foreign"} {
			t.Run(name+map[bool]string{false: " transitional", true: " strict"}[strict], func(t *testing.T) {
				parts := legacyOriginParts()
				s := parts["word/styles.xml"]
				d := parts["word/document.xml"]
				switch name {
				case "direct margin", "zero margin":
					value := "216"
					if name == "zero margin" {
						value = "0"
					}
					d = strings.Replace(d, `<w:tblStyle w:val="Child"/>`, `<w:tblStyle w:val="Child"/><w:tblCellMar><w:left w:type="dxa" w:w="`+value+`"/></w:tblCellMar>`, 1)
				case "missing ancestor":
					s = strings.Replace(s, `w:val="Base"`, `w:val="Missing"`, 1)
				case "merged":
					d = strings.Replace(d, `<w:tcPr>`, `<w:tcPr><w:gridSpan w:val="2"/>`, 1)
				case "floating":
					d = strings.Replace(d, `<w:tblPr>`, `<w:tblPr><w:tblpPr w:tblpX="0"/>`, 1)
				case "spacing":
					d = strings.Replace(d, `<w:tblPr>`, `<w:tblPr><w:tblCellSpacing w:w="20" w:type="dxa"/>`, 1)
				case "direct override":
					d = strings.Replace(d, `<w:tblStyle w:val="Child"/>`, `<w:tblStyle w:val="Child"/><w:tblInd w:type="dxa" w:w="200"/>`, 1)
				case "missing indent":
					s = strings.Replace(s, `<w:tblInd w:type="dxa" w:w="0"/>`, "", 1)
				case "missing margin":
					s = strings.Replace(s, `<w:left w:type="dxa" w:w="108"/>`, "", 1)
				case "default margin type":
					s = strings.Replace(s, `<w:left w:type="dxa"`, `<w:left`, 1)
				case "duplicate indent":
					s = strings.Replace(s, `<w:tblInd w:type="dxa" w:w="0"/>`, `<w:tblInd w:type="dxa" w:w="0"/><w:tblInd w:type="dxa" w:w="1"/>`, 1)
				case "malformed margin":
					s = strings.Replace(s, `w:w="108"`, `w:w="oops"`, 1)
				case "conditional":
					s = strings.Replace(s, `w:styleId="Base">`, `w:styleId="Base"><w:tblStylePr w:type="firstRow"/>`, 1)
				case "cycle":
					s = strings.Replace(s, `w:styleId="Base">`, `w:styleId="Base"><w:basedOn w:val="Child"/>`, 1)
				case "cell override":
					d = strings.Replace(d, `<w:tcPr>`, `<w:tcPr><w:tcMar><w:left w:type="dxa" w:w="12"/></w:tcMar>`, 1)
				case "style cell override":
					s = strings.Replace(s, `w:styleId="Base">`, `w:styleId="Base"><w:tcPr><w:tcMar><w:left w:type="dxa" w:w="12"/></w:tcMar></w:tcPr>`, 1)
				case "row override":
					d = strings.Replace(d, `<w:tr>`, `<w:tr><w:tblPrEx><w:tblInd w:type="dxa" w:w="12"/></w:tblPrEx>`, 1)
				case "rtl":
					d = strings.Replace(d, `<w:tblPr>`, `<w:tblPr><w:bidiVisual/>`, 1)
				case "foreign":
					s = strings.Replace(s, `<w:tblInd `, `<w:tblInd xmlns:w="urn:foreign" `, 1)
				}
				parts["word/styles.xml"] = s
				parts["word/document.xml"] = d
				if strict {
					for k, v := range parts {
						parts[k] = strings.ReplaceAll(strings.ReplaceAll(v, wordMLTransitional, wordMLStrict), relBaseTransitional, relBaseStrict)
					}
				}
				data := buildNativeDOCX(t, nativeEntries(parts))
				before := string(data)
				result, err := ExtractNativeDocxApproximationEligibilityV1(data)
				if name == "foreign" {
					if err == nil {
						t.Fatal("spoofed namespace did not refuse")
					}
					return
				}
				if err != nil {
					t.Fatal(err)
				}
				positive := name == "inherited" || name == "direct override" || name == "direct margin"
				if !positive {
					if len(result.LegacyTableOrigins) != 0 {
						t.Fatalf("unexpected fact %#v", result.LegacyTableOrigins)
					}
					return
				}
				if len(result.LegacyTableOrigins) != 1 {
					t.Fatalf("missing evidence %#v", result)
				}
				fact := result.LegacyTableOrigins[0]
				wantPart := "word/styles.xml"
				wantIndent := int64(0)
				if name == "direct override" {
					wantPart = "word/document.xml"
					wantIndent = 200
				}
				digest := sha256.Sum256([]byte(parts[wantPart]))
				wantMargin := int64(108)
				marginPart := "word/styles.xml"
				indentPath := "/w:styles[1]/w:style[1]/w:tblPr[1]/w:tblInd[1]"
				marginPath := "/w:styles[1]/w:style[1]/w:tblPr[1]/w:tblCellMar[1]/w:left[1]"
				if name == "direct override" {
					indentPath = "/w:document[1]/w:body[1]/w:tbl[1]/w:tblPr[1]/w:tblInd[1]"
				}
				if name == "direct margin" {
					wantMargin = 216
					marginPart = "word/document.xml"
					marginPath = "/w:document[1]/w:body[1]/w:tbl[1]/w:tblPr[1]/w:tblCellMar[1]/w:left[1]"
				}
				marginDigest := sha256.Sum256([]byte(parts[marginPart]))
				if fact.SourceIndent.PartName != wantPart || fact.SourceIndent.Path != indentPath || fact.SourceMargin.Path != marginPath || fact.SourceMargin.PartName != marginPart || fact.SourceMargin.SHA256 != "sha256:"+hex.EncodeToString(marginDigest[:]) || fact.SourceIndent.SHA256 != "sha256:"+hex.EncodeToString(digest[:]) || fact.IndentTwips != wantIndent || fact.LeftMarginTwips != wantMargin || fact.PackageSHA256 != result.PackageSHA256 {
					t.Fatalf("wrong evidence %#v", fact)
				}
				if string(data) != before {
					t.Fatal("source changed")
				}
			})
		}
	}
}

func TestLegacyTableOriginModeGate(t *testing.T) {
	for _, mode := range []string{"12", "14", "15", "oops"} {
		parts := legacyOriginParts()
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/></Types>`, 1)
		parts["word/_rels/document.xml.rels"] = strings.Replace(parts["word/_rels/document.xml.rels"], `</Relationships>`, `<Relationship Id="settings" Type="`+relBaseTransitional+`settings" Target="settings.xml"/></Relationships>`, 1)
		parts["word/settings.xml"] = `<w:settings xmlns:w="` + wordMLTransitional + `"><w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="` + mode + `"/></w:compat></w:settings>`
		result, err := ExtractNativeDocxApproximationEligibilityV1(buildNativeDOCX(t, nativeEntries(parts)))
		if err != nil {
			t.Fatal(err)
		}
		if (len(result.LegacyTableOrigins) == 1) != (mode == "12") {
			t.Fatalf("mode%s: %#v", mode, result)
		}
	}
}
