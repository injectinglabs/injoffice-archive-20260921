package xlsxpatch

import (
	"strings"
	"testing"
)

func TestStoredRowPreviewQualification(t *testing.T) {
	for _, ns := range []string{spreadsheetMLTransitional, spreadsheetMLStrict} {
		original := `<worksheet xmlns="` + ns + `" xmlns:x14ac="` + nativeRowDescentNamespace + `"><sheetFormatPr defaultRowHeight="14.4" customHeight="false" x14ac:dyDescent="0.3"/><sheetData><row r="1" spans="1:3" x14ac:dyDescent="0.3"/><row r="2" ht="24" customHeight="true"/><row r="3" hidden="true"/></sheetData></worksheet>`
		got := previewNativeStoredRows([]byte(original), "xl/worksheets/s.xml")
		if len(got.Rows) != 32 || got.Rows[0].HeightPoints != 14.4 || got.Rows[1].HeightPoints != 24 || !got.Rows[2].Hidden || len(got.Warnings) != 2 {
			t.Fatalf("wrong stored geometry: %+v", got)
		}
		for _, variant := range []string{"automatic-default", "automatic-override", "automatic-inherited", "foreign-descent", "invalid-descent", "hex-default", "hex-height", "hex-descent", "unknown-attribute", "invalid-spans", "duplicate-row", "zero-height", "text", "data-attrs"} {
			source := original
			switch variant {
			case "automatic-default":
				source = strings.Replace(source, ` x14ac:dyDescent="0.3"`, "", 1)
			case "automatic-override":
				source = strings.Replace(source, `ht="24" customHeight="true"`, `ht="24" customHeight="false"`, 1)
			case "automatic-inherited":
				source = strings.Replace(source, `r="3" hidden="true"`, `r="3" customHeight="false"`, 1)
			case "foreign-descent":
				source = strings.ReplaceAll(source, nativeRowDescentNamespace, "urn:foreign")
			case "invalid-descent":
				source = strings.Replace(source, `dyDescent="0.3"`, `dyDescent="NaN"`, 1)
			case "hex-default":
				source = strings.Replace(source, `defaultRowHeight="14.4"`, `defaultRowHeight="0x1p2"`, 1)
			case "hex-height":
				source = strings.Replace(source, `ht="24"`, `ht="0x1p2"`, 1)
			case "hex-descent":
				source = strings.Replace(source, `dyDescent="0.3"`, `dyDescent="0x1p2"`, 1)
			case "unknown-attribute":
				source = strings.Replace(source, `r="1"`, `r="1" thickBot="1"`, 1)
			case "invalid-spans":
				source = strings.Replace(source, `spans="1:3"`, `spans="3:1"`, 1)
			case "duplicate-row":
				source = strings.Replace(source, `r="2"`, `r="1"`, 1)
			case "zero-height":
				source = strings.Replace(source, `customHeight="false"`, `customHeight="true" zeroHeight="true"`, 1)
			case "text":
				source = strings.Replace(source, `<sheetData>`, `<sheetData>unknown`, 1)
			case "data-attrs":
				source = strings.Replace(source, `<sheetData>`, `<sheetData future="1">`, 1)
			}
			if result := previewNativeStoredRows([]byte(source), "xl/worksheets/s.xml"); len(result.Rows) != 0 || len(result.Warnings) == 0 {
				t.Fatalf("%s falsely qualified: %+v", variant, result)
			}
		}
	}
}

func TestStoredRowPreviewExplicitDefaultAndZeroOverride(t *testing.T) {
	source := `<worksheet xmlns="` + spreadsheetMLTransitional + `"><sheetFormatPr defaultRowHeight="20" customHeight="1"/><sheetData><row r="1" ht="0" customHeight="1"/><row r="33" ht="40" customHeight="1"/></sheetData></worksheet>`
	got := previewNativeStoredRows([]byte(source), "sheet.xml")
	if len(got.Rows) != 32 || !got.Rows[0].Hidden || got.Rows[0].HeightPoints != 0 || got.Rows[31].HeightPoints != 20 {
		t.Fatalf("bad bounds or zero: %+v", got)
	}
}
