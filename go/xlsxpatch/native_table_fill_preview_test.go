package xlsxpatch

import (
	"strings"
	"testing"
)

func TestTableFillHLSTint(t *testing.T) {
	for _, sample := range [][2]string{{"#156082", "#C0E6F5"}, {"#4F81BD", "#DCE6F1"}, {"#5B9BD5", "#DDEBF7"}, {"#000000", "#CCCCCC"}, {"#FFFFFF", "#FFFFFF"}} {
		if got := tableLightenHLS(sample[0], 0.8); got != sample[1] {
			t.Errorf("%s => %s, want %s", sample[0], got, sample[1])
		}
	}
}

func TestTableStyleIDProjectionBudget(t *testing.T) {
	table := NativeTablePreviewV1{FillPreview: &NativeTableFillPreviewV1{FillStyleIDs: make([]int, 4096), HeaderFontStyleIDs: make([]int, 4096)}}
	if !nativeTableStyleIDsWithinBudget([]NativeTablePreviewV1{table, table}) || nativeTableStyleIDsWithinBudget([]NativeTablePreviewV1{table, table, table}) {
		t.Fatal("cumulative style-ID budget not enforced at 16384")
	}
}

func TestNativeTableFillQualification(t *testing.T) {
	for _, blocked := range []string{"", "conditional", "custom", "dxf", "transformed-theme", "missing-theme", "nondefault-fill-zero", "invalid-boolean", "unknown-option", "option-child", "option-text", "explicit-no-fill", "ignored-direct-fill", "named-base-fill"} {
		parts := nativeWorkbookFixture(false)
		parts["Charts/chart1.xml"] = previewChartFixture()
		parts["Sheets/s1.xml"] = `<worksheet xmlns="` + spreadsheetMLTransitional + `" xmlns:r="` + officeRelNamespaceTransitional + `"><sheetData/><tableParts count="1"><tablePart r:id="table1"/></tableParts></worksheet>`
		parts["Sheets/_rels/s1.xml.rels"] = strings.Replace(parts["Sheets/_rels/s1.xml.rels"], `</Relationships>`, `<Relationship Id="table1" Type="`+officeRelNamespaceTransitional+`/table" Target="../Tables/table.xml"/></Relationships>`, 1)
		parts["Tables/table.xml"] = `<table xmlns="` + spreadsheetMLTransitional + `" displayName="Original" ref="A1:C4"><tableStyleInfo name="TableStyleMedium2" showRowStripes="1"/></table>`
		parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`, `<Override PartName="/Tables/table.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/><Override PartName="/theme/theme1.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
		parts["Book/_rels/Workbook.xml.rels"] = strings.Replace(parts["Book/_rels/Workbook.xml.rels"], `</Relationships>`, `<Relationship Id="theme" Type="`+relTypeThemeTransitional+`" Target="../theme/theme1.xml"/></Relationships>`, 1)
		parts["theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLNamespace + `"><a:themeElements><a:clrScheme><a:accent1><a:srgbClr val="156082"/></a:accent1></a:clrScheme></a:themeElements></a:theme>`
		switch blocked {
		case "explicit-no-fill":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`, `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyFill="1"/>`, 1)
		case "ignored-direct-fill":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`, `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0" applyFill="0"/>`, 1)
		case "named-base-fill":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>`, `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="2" borderId="0"/>`, 1)
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`, `<xf numFmtId="0" fontId="0" fillId="2" borderId="0" xfId="0"/>`, 1)
		case "invalid-boolean":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `showRowStripes="1"`, `showRowStripes="yes"`, 1)
		case "unknown-option":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `showRowStripes="1"`, `showRowStripes="1" future="1"`, 1)
		case "option-child":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `showRowStripes="1"/>`, `showRowStripes="1"><future/></tableStyleInfo>`, 1)
		case "option-text":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `showRowStripes="1"/>`, `showRowStripes="1">unknown</tableStyleInfo>`, 1)
		case "conditional":
			parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</worksheet>`, `<conditionalFormatting sqref="A1"><cfRule type="expression" priority="1"><formula>1</formula></cfRule></conditionalFormatting></worksheet>`, 1)
		case "custom":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `</styleSheet>`, `<tableStyles count="1"><tableStyle name="TableStyleMedium2" count="0"/></tableStyles></styleSheet>`, 1)
		case "dxf":
			parts["Tables/table.xml"] = strings.Replace(parts["Tables/table.xml"], `displayName=`, `dataDxfId="0" displayName=`, 1)
		case "transformed-theme":
			parts["theme/theme1.xml"] = strings.Replace(parts["theme/theme1.xml"], `val="156082"/>`, `val="156082"><a:tint val="40000"/></a:srgbClr>`, 1)
		case "missing-theme":
			parts["theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLNamespace + `"><a:themeElements/></a:theme>`
		case "nondefault-fill-zero":
			parts["Meta/Styles.style"] = strings.Replace(parts["Meta/Styles.style"], `<patternFill patternType="none"/>`, `<patternFill patternType="solid"><fgColor rgb="FFAABBCC"/></patternFill>`, 1)
		}
		got, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
		if blocked == "ignored-direct-fill" {
			if err == nil || !strings.Contains(err.Error(), "apply flag is false or absent") {
				t.Fatalf("existing mismatched apply-flag refusal changed: %v", err)
			}
			continue
		}
		if err != nil {
			t.Fatalf("%s: %v", blocked, err)
		}
		palette := got.Tables[0].FillPreview
		if blocked == "explicit-no-fill" || blocked == "named-base-fill" {
			if palette == nil || len(palette.FillStyleIDs) != 0 {
				t.Fatalf("%s overwrote direct/inherited fill: %+v", blocked, palette)
			}
			continue
		}
		if blocked != "" {
			if palette != nil {
				t.Fatalf("%s falsely qualified: %+v", blocked, palette)
			}
			continue
		}
		if palette == nil || palette.Header != "#156082" || palette.Stripe != "#C0E6F5" {
			t.Fatalf("lost qualified source palette: %+v", palette)
		}
		if len(palette.HeaderFontStyleIDs) != 1 || palette.HeaderFontStyleIDs[0] != 0 {
			t.Fatalf("explicit font style incorrectly qualified: %+v", palette.HeaderFontStyleIDs)
		}
	}
}
