package xlsxpatch

import (
	"strings"
	"testing"
)

func TestExtractNativeWorkbookV2ProjectsDisplayIdentity(t *testing.T) {
	entries := nativeWorkbookFixture(false)
	entries["Book/Workbook.xml"] = strings.Replace(entries["Book/Workbook.xml"], `<bookViews>`, `<workbookPr date1904="0"/><bookViews>`, 1)
	entries["Book/_rels/Workbook.xml.rels"] = strings.Replace(entries["Book/_rels/Workbook.xml.rels"], `</Relationships>`,
		`<Relationship Id="rTheme" Type="`+relTypeThemeTransitional+`" Target="../theme/theme1.xml"/></Relationships>`, 1)
	entries["[Content_Types].xml"] = strings.Replace(entries["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/theme/theme1.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
	entries["theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLNamespace + `" name="Office"><a:themeElements><a:clrScheme name="Office">` +
		`<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="4472C4"><a:tint val="40000"/></a:srgbClr></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>` +
		`<a:minorFont><a:latin typeface="ThemeMinorFace"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>`
	entries["Meta/Styles.style"] = strings.Replace(entries["Meta/Styles.style"],
		`<alignment horizontal="center" vertical="center" wrapText="1"/>`,
		`<alignment horizontal="center" vertical="center" wrapText="1" shrinkToFit="1" textRotation="90"/>`, 1)

	workbook, err := ExtractNativeWorkbookV2(buildZip(t, entries))
	if err != nil {
		t.Fatal(err)
	}
	if workbook.Date1904 == nil || *workbook.Date1904 {
		t.Fatalf("date1904 was not projected as false: %#v", workbook.Date1904)
	}
	for _, capability := range workbook.Capabilities {
		if capability.Name == "native-cell-glyphs" {
			t.Fatal("Go extract advertised glyph-paint identity")
		}
	}
	style0 := workbook.Styles[0].Effective
	if style0.FontName == nil || *style0.FontName != "ThemeMinorFace" {
		t.Fatalf("theme minor latin typeface was not resolved: %#v", style0)
	}
	if workbook.NormalStyle == nil || workbook.NormalStyle.FontName != "ThemeMinorFace" {
		t.Fatalf("Normal style did not take the theme minor typeface: %#v", workbook.NormalStyle)
	}
	if style0.FontColor == nil || *style0.FontColor != "#000000" {
		t.Fatalf("theme dk1 lastClr RGB was not resolved: %#v", style0)
	}
	if containsNativeStyleUnsupported(style0.Unsupported, "font-color") {
		t.Fatalf("resolved theme color still marked unsupported: %#v", style0)
	}
	style1 := workbook.Styles[1].Effective
	if style1.ShrinkToFit == nil || !*style1.ShrinkToFit || style1.TextRotation == nil || *style1.TextRotation != 90 {
		t.Fatalf("shrink/rotation were not projected: %#v", style1)
	}
	if containsNativeStyleUnsupported(style1.Unsupported, "alignment-extended") {
		t.Fatalf("exact shrink/rotation still marked alignment-extended: %#v", style1)
	}
	shared := findNativeCellV2(t, workbook, "7", "A1")
	if shared.Value == nil || !shared.Value.Rich || len(shared.Value.Runs) != 2 || shared.Value.Runs[0].Text != "Rich " || shared.Value.Runs[1].Text != "Text" {
		t.Fatalf("rich runs were not projected: %#v", shared.Value)
	}
	if shared.Value.Runs[0].Bold == nil || !*shared.Value.Runs[0].Bold {
		t.Fatalf("bold run identity was lost: %#v", shared.Value.Runs[0])
	}
}

func TestExtractNativeWorkbookV2ReconstructsRunsAndTintedSchemeColors(t *testing.T) {
	entries := nativeWorkbookFixture(false)
	entries["Book/_rels/Workbook.xml.rels"] = strings.Replace(entries["Book/_rels/Workbook.xml.rels"], `</Relationships>`,
		`<Relationship Id="rTheme" Type="`+relTypeThemeTransitional+`" Target="../theme/theme1.xml"/></Relationships>`, 1)
	entries["[Content_Types].xml"] = strings.Replace(entries["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/theme/theme1.xml" ContentType="`+themePartContentType+`"/></Types>`, 1)
	entries["theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLNamespace + `" name="Office"><a:themeElements><a:clrScheme name="Office">` +
		`<a:dk1><a:sysClr val="windowText"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="4472C4"><a:tint val="40000"/></a:srgbClr></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme></a:themeElements></a:theme>`
	entries["Meta/Strings.xml"] = `<sst xmlns="` + spreadsheetMLTransitional + `" count="3" uniqueCount="3">` +
		`<si><r><t>Solo</t></r></si><si><t>Plain</t></si><si><t>Ruby</t><rPh sb="0" eb="4"><t>ル</t></rPh><phoneticPr fontId="0"/></si></sst>`
	entries["Meta/Styles.style"] = strings.Replace(entries["Meta/Styles.style"],
		`<color theme="1"/>`,
		`<color theme="4"/>`, 1)
	entries["Sheets/s1.xml"] = strings.Replace(entries["Sheets/s1.xml"],
		`<c r="A1" s="1" t="s"><v>0</v></c>`,
		`<c r="A1" s="0" t="s"><v>0</v></c>`, 1)
	entries["Sheets/s1.xml"] = strings.Replace(entries["Sheets/s1.xml"],
		`</row>`,
		`<c r="L1" t="s"><v>2</v></c></row>`, 1)

	workbook, err := ExtractNativeWorkbookV2(buildZip(t, entries))
	if err != nil {
		t.Fatal(err)
	}
	for _, capability := range workbook.Capabilities {
		if capability.Name == "native-cell-glyphs" {
			t.Fatal("Go extract advertised glyph-paint identity")
		}
	}
	solo := findNativeCellV2(t, workbook, "7", "A1")
	if solo.Value == nil || !solo.Value.Rich || len(solo.Value.Runs) != 1 || solo.Value.Runs[0].Text != "Solo" {
		t.Fatalf("SST <r> without rPr did not emit a reconstructible run: %#v", solo.Value)
	}
	ruby := findNativeCellV2(t, workbook, "7", "L1")
	if ruby.Value == nil || !ruby.Value.Rich || len(ruby.Value.Runs) != 1 || ruby.Value.Runs[0].Text != "Ruby" || ruby.Value.Text == nil || *ruby.Value.Text != "Ruby" {
		t.Fatalf("SST text plus phonetic did not emit a reconstructible run: %#v", ruby.Value)
	}
	style0 := workbook.Styles[0].Effective
	if style0.FontColor == nil || *style0.FontColor != "#8FAADC" {
		t.Fatalf("linear srgbClr tint was not applied: %#v", style0)
	}
	if containsNativeStyleUnsupported(style0.Unsupported, "font-color") {
		t.Fatalf("tinted scheme color still marked unsupported: %#v", style0)
	}

	tintedEntries := nativeWorkbookFixture(false)
	tintedEntries["Book/_rels/Workbook.xml.rels"] = entries["Book/_rels/Workbook.xml.rels"]
	tintedEntries["[Content_Types].xml"] = entries["[Content_Types].xml"]
	tintedEntries["theme/theme1.xml"] = `<a:theme xmlns:a="` + drawingMLNamespace + `" name="Office"><a:themeElements><a:clrScheme name="Office">` +
		`<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
		`<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>` +
		`<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
		`<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>` +
		`<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>` +
		`<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>` +
		`</a:clrScheme></a:themeElements></a:theme>`
	tintedEntries["Meta/Styles.style"] = strings.Replace(tintedEntries["Meta/Styles.style"],
		`<color theme="1"/>`,
		`<color theme="4" tint="0.4"/>`, 1)
	tinted, err := ExtractNativeWorkbookV2(buildZip(t, tintedEntries))
	if err != nil {
		t.Fatal(err)
	}
	if tinted.Styles[0].Effective.FontColor == nil || *tinted.Styles[0].Effective.FontColor != "#8FAADC" {
		t.Fatalf("spreadsheet theme tint was not applied linearly: %#v", tinted.Styles[0].Effective)
	}
}

func findNativeCellV2(t *testing.T, workbook *NativeWorkbookV2, sheet, ref string) NativeWorkbookCellV2 {
	t.Helper()
	for _, candidate := range workbook.Sheets {
		if candidate.ID != sheet {
			continue
		}
		for _, cell := range candidate.Cells {
			if cell.Ref == ref {
				return cell
			}
		}
	}
	t.Fatalf("cell %s!%s not found", sheet, ref)
	return NativeWorkbookCellV2{}
}
