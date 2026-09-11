package xlsxhttp

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
	"net/http"
	"net/http/httptest"
	"testing"
)

func TestObjectsPreviewReadOnlyHTTP(t *testing.T) {
	const ss = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
	const rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
	files := map[string]string{
		"[Content_Types].xml":        `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>`,
		"_rels/.rels":                `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="` + rel + `officeDocument" Target="xl/workbook.xml"/></Relationships>`,
		"xl/workbook.xml":            `<workbook xmlns="` + ss + `" xmlns:r="` + rel[:len(rel)-1] + `"><sheets><sheet name="Original" sheetId="1" r:id="r1"/></sheets></workbook>`,
		"xl/_rels/workbook.xml.rels": `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="` + rel + `worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
		"xl/worksheets/sheet1.xml":   `<worksheet xmlns="` + ss + `"><sheetData/></worksheet>`,
	}
	var data bytes.Buffer
	z := zip.NewWriter(&data)
	for name, value := range files {
		f, err := z.Create(name)
		if err != nil {
			t.Fatal(err)
		}
		if _, err = f.Write([]byte(value)); err != nil {
			t.Fatal(err)
		}
	}
	if err := z.Close(); err != nil {
		t.Fatal(err)
	}
	original := bytes.Clone(data.Bytes())
	handler := NewHandler(nil)
	request := httptest.NewRequest(http.MethodPost, PreviewObjectsPath, bytes.NewReader(original))
	request.Header.Set("Content-Type", XLSXContentType)
	response := httptest.NewRecorder()
	handler.ServeHTTP(response, request)
	if response.Code != http.StatusOK {
		t.Fatalf("HTTP %d: %s", response.Code, response.Body.String())
	}
	var projection xlsxpatch.NativeWorkbookObjectsV1
	if err := json.Unmarshal(response.Body.Bytes(), &projection); err != nil {
		t.Fatal(err)
	}
	if projection.Protocol != "injoffice.xlsx.preview-objects" || projection.PackageSHA256 == "" || len(projection.Charts) != 0 {
		t.Fatalf("invalid projection %+v", projection)
	}
	if response.Header().Get(HeaderArtifactID) != "" || !bytes.Equal(original, data.Bytes()) {
		t.Fatal("read-only inspection persisted or changed source")
	}
	for _, method := range []string{http.MethodGet, http.MethodDelete} {
		r := httptest.NewRecorder()
		handler.ServeHTTP(r, httptest.NewRequest(method, PreviewObjectsPath, nil))
		if r.Code != http.StatusMethodNotAllowed {
			t.Fatal("unexpected method accepted")
		}
	}
}
