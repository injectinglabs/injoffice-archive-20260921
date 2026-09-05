package xlsxpatch

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"os"
	"strings"
	"testing"
)

func readExcelAuthoredFixture(t *testing.T, name, expectedSHA string) []byte {
	t.Helper()
	data, err := os.ReadFile("testdata/excel-authored/" + name)
	if err != nil {
		t.Fatal(err)
	}
	digest := sha256.Sum256(data)
	if actual := hex.EncodeToString(digest[:]); actual != expectedSHA {
		t.Fatalf("%s bytes changed: sha256=%s, want %s", name, actual, expectedSHA)
	}
	return data
}

func TestExcelAuthoredNativeFixturesKeepOfficeAuthority(t *testing.T) {
	happy := readExcelAuthoredFixture(t, "happy-tree.xlsx", "c08f0bb099770a475556af8a78d9d6e46296e9791bf9af56c70571c1fa3cb513")
	if app := readEntry(t, happy, "docProps/app.xml"); !strings.Contains(app, "<Application>Microsoft Excel Online</Application>") || !strings.Contains(app, "<AppVersion>16.0300</AppVersion>") {
		t.Fatalf("positive fixture lost Excel authoring provenance: %s", app)
	}
	happyWorkbook, err := ExtractNativeWorkbookV1(happy)
	if err != nil {
		t.Fatal(err)
	}
	if hasNativeWorkbookUnsupported(happyWorkbook, "WORKSHEET_ATTRIBUTES") || hasNativeWorkbookUnsupported(happyWorkbook, "SHEET_VIEW_GEOMETRY") {
		t.Fatalf("known Excel revision metadata or selection-only view was treated as unmodeled appearance authority: %#v", happyWorkbook.Unsupported)
	}
	happyEncoded, err := EncodeNativeWorkbookV1(happyWorkbook)
	if err != nil {
		t.Fatal(err)
	}
	happyExpected, err := os.ReadFile("testdata/native-xlsx-v1/valid/excel-authored-happy-tree.json")
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(append(happyEncoded, '\n'), happyExpected) {
		t.Fatal("committed TypeScript bridge is not the exact Go Extract/Encode output")
	}
	if decoded, decodeErr := DecodeNativeWorkbookV1(happyEncoded); decodeErr != nil || decoded.Revision != happyWorkbook.Revision {
		t.Fatalf("positive Excel fixture did not survive Go decode: decoded=%#v err=%v", decoded, decodeErr)
	}

	email := readExcelAuthoredFixture(t, "email-chart-table.xlsx", "1da5a2011f4f1e6f2e4407f148e814cae47e93d97335c6819d582b6cf433e1be")
	if app := readEntry(t, email, "docProps/app.xml"); !strings.Contains(app, "<Application>Microsoft Excel Online</Application>") || !strings.Contains(app, "<AppVersion>16.0300</AppVersion>") {
		t.Fatalf("fixture lost Excel authoring provenance: %s", app)
	}
	emailWorkbook, err := ExtractNativeWorkbookV1(email)
	if err != nil {
		t.Fatal(err)
	}
	if len(emailWorkbook.Styles) != 12 {
		t.Fatalf("Excel style count = %d, want 12", len(emailWorkbook.Styles))
	}
	explicitCurrency := `"$"#,##0.00_);[Red]\("$"#,##0.00\)`
	if format := nativeOptionalString(emailWorkbook.Styles[5].Effective.NumberFormat); format != explicitCurrency {
		t.Fatalf("Excel-authored explicit built-in ID override = %q, want %q", format, explicitCurrency)
	}
	border := emailWorkbook.Styles[1].Effective.Border
	if border == nil || border.Origin != "styles-record" || border.BorderID == nil || *border.BorderID != 1 || !nativeWorkbookSHA.MatchString(nativeOptionalString(border.RecordSHA256)) || border.Top == nil || border.Top.Style != "thin" || border.Top.Color != "#95B3D7" || border.Bottom == nil || border.Bottom.Style != "thin" || border.Bottom.Color != "#95B3D7" {
		t.Fatalf("Excel-authored direct-RGB border was not projected exactly: %#v", border)
	}
	fill := emailWorkbook.Styles[1].Effective.Fill
	if fill == nil || fill.Origin != "styles-record" || fill.FillID == nil || *fill.FillID != 3 || fill.Color == nil || *fill.Color != "#D9E1F2" || !nativeWorkbookSHA.MatchString(nativeOptionalString(fill.RecordSHA256)) {
		t.Fatalf("Excel-authored direct-RGB solid fill was not projected exactly: %#v", fill)
	}
	emailEncoded, err := EncodeNativeWorkbookV1(emailWorkbook)
	if err != nil {
		t.Fatal(err)
	}
	if decoded, decodeErr := DecodeNativeWorkbookV1(emailEncoded); decodeErr != nil || decoded.Source.PackageSHA256 != emailWorkbook.Source.PackageSHA256 {
		t.Fatalf("unmodified Excel-authored fixture did not survive encode/decode: decoded=%#v err=%v", decoded, decodeErr)
	}
	conditional := readExcelAuthoredFixture(t, "conditional-formatting-samples.xlsx", "fa17f45f47e0766f13b9cbbf9e63b83962ea7852606bfe0b33807fbfbae5ec64")
	if app := readEntry(t, conditional, "docProps/app.xml"); !strings.Contains(app, "<Application>Microsoft Excel</Application>") {
		t.Fatalf("fixture lost desktop Excel authoring provenance: %s", app)
	}
	conditionalWorkbook, err := ExtractNativeWorkbookV1(conditional)
	if err != nil {
		t.Fatal(err)
	}
	if !hasNativeWorkbookUnsupported(conditionalWorkbook, "UNMODELED_WORKSHEET_FEATURE") || !hasNativeWorkbookUnsupported(conditionalWorkbook, "STYLE_TABLE_OPAQUE_CONTENT") {
		t.Fatalf("Excel-authored unmodeled authority was not refused: %#v", conditionalWorkbook.Unsupported)
	}
	first, err := EncodeNativeWorkbookV1(conditionalWorkbook)
	if err != nil {
		t.Fatal(err)
	}
	second, err := EncodeNativeWorkbookV1(conditionalWorkbook)
	if err != nil || !bytes.Equal(first, second) {
		t.Fatalf("native extraction is nondeterministic: err=%v", err)
	}
	decoded, err := DecodeNativeWorkbookV1(first)
	if err != nil || decoded.Revision != conditionalWorkbook.Revision || decoded.Source.PackageSHA256 != conditionalWorkbook.Source.PackageSHA256 {
		t.Fatalf("native contract lost package identity: decoded=%#v err=%v", decoded, err)
	}
}
