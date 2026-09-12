//go:build !js

package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func TestWASMInspectSavedPrintAreaMatchesGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	for _, tc := range []struct {
		name, ref       string
		available       bool
		fit             string
		fitAvailable    bool
		titles          string
		titlesAvailable bool
	}{
		{"rectangle", `Sheet1!$B$2:$D$4`, true, "", false, "", false},
		{"union", `Sheet1!$B$2:$D$4,Sheet1!$F$6`, false, "", false, "", false},
		{"fit", `Sheet1!$B$2:$D$4`, true, "1", true, "", false},
		{"fit-unbounded", `Sheet1!$B$2:$D$4`, true, "0", true, "", false},
		{"fit-invalid", `Sheet1!$B$2:$D$4`, true, "101", false, "", false},
		{"titles-both", `Sheet1!$B$2:$D$4`, true, "", false, `Sheet1!$1:$2,Sheet1!$A:$B`, true},
		{"titles-invalid", `Sheet1!$B$2:$D$4`, false, "", false, `Sheet1!$A$1:$B$2`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := map[string]string{
				"[Content_Types].xml":        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
				"_rels/.rels":                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
				"xl/workbook.xml":            `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="17" r:id="r1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">` + tc.ref + `</definedName></definedNames></workbook>`,
				"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
				"xl/worksheets/sheet1.xml":   `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>Inside</t></is></c></row></sheetData></worksheet>`,
			}
			if tc.fit != "" {
				parts["xl/worksheets/sheet1.xml"] = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetPr><pageSetUpPr fitToPage="true"/></sheetPr><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="` + tc.fit + `"/></worksheet>`
			}
			if tc.titles != "" {
				parts["xl/workbook.xml"] = strings.Replace(parts["xl/workbook.xml"], `</definedNames>`, `<definedName name="_xlnm.Print_Titles" localSheetId="0">`+tc.titles+`</definedName></definedNames>`, 1)
			}
			var buf bytes.Buffer
			z := zip.NewWriter(&buf)
			for name, raw := range parts {
				w, e := z.Create(name)
				if e != nil {
					t.Fatal(e)
				}
				if _, e = w.Write([]byte(raw)); e != nil {
					t.Fatal(e)
				}
			}
			if e := z.Close(); e != nil {
				t.Fatal(e)
			}
			input := filepath.Join(t.TempDir(), "area.xlsx")
			if e := os.WriteFile(input, buf.Bytes(), 0600); e != nil {
				t.Fatal(e)
			}
			want, e := xlsxpatch.InspectNativeWorkbookObjectsV1(buf.Bytes())
			if e != nil {
				t.Fatal(e)
			}
			if len(want.PrintAreas) != 1 || (want.PrintAreas[0].Status == "available") != tc.available {
				t.Fatalf("%+v", want.PrintAreas)
			}
			if len(want.PrintTitles) != 1 || (want.PrintTitles[0].Status == "available") != tc.titlesAvailable {
				t.Fatalf("%+v", want.PrintTitles)
			}
			if tc.fit != "" && (len(want.PageSettings) != 1 || (want.PageSettings[0].Status == "available") != tc.fitAvailable) {
				t.Fatalf("unexpected Go fit settings: %+v", want.PageSettings)
			}
			encoded, e := json.Marshal(want)
			if e != nil {
				t.Fatal(e)
			}
			cmd := exec.Command("node", script, "inspect", "--wasm", wasm, "--wasm-exec", wasmExec, "--input", input)
			got, e := cmd.Output()
			if e != nil {
				t.Fatalf("WASM inspect failed: %v\n%s", e, stderrFrom(e))
			}
			if !bytes.Equal(bytes.TrimSpace(got), encoded) {
				t.Fatalf("WASM and Go inspect differ: %s", got)
			}
		})
	}
}
