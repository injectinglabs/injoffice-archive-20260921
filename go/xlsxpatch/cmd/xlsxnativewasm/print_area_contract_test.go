//go:build !js

package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"os"
	"os/exec"
	"path/filepath"
	"testing"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func TestWASMInspectSavedPrintAreaMatchesGo(t *testing.T) {
	wasm, wasmExec, script := requireNodeHarness(t)
	for _, tc := range []struct {
		name, ref string
		available bool
	}{
		{"rectangle", `Sheet1!$B$2:$D$4`, true},
		{"union", `Sheet1!$B$2:$D$4,Sheet1!$F$6`, false},
	} {
		t.Run(tc.name, func(t *testing.T) {
			parts := map[string]string{
				"[Content_Types].xml":        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
				"_rels/.rels":                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
				"xl/workbook.xml":            `<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="17" r:id="r1"/></sheets><definedNames><definedName name="_xlnm.Print_Area" localSheetId="0">` + tc.ref + `</definedName></definedNames></workbook>`,
				"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
				"xl/worksheets/sheet1.xml":   `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="2"><c r="B2" t="inlineStr"><is><t>Inside</t></is></c></row></sheetData></worksheet>`,
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
