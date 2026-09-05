package officecompat_test

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"os"
	"path/filepath"
	"sort"
	"strings"
	"testing"

	"github.com/injectinglabs/injoffice/go/officecompat/corpus"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const xlsxCorpusRoot = "corpus"

func TestRoundTripXLSXNativeCompatibilityCorpus(t *testing.T) {
	expectedPaths, err := filepath.Glob(filepath.Join(xlsxCorpusRoot, "generated", "expected", "xlsx-*.json"))
	if err != nil {
		t.Fatal(err)
	}
	sort.Strings(expectedPaths)
	if len(expectedPaths) != 8 {
		t.Fatalf("XLSX corpus has %d fixtures, want 8; run go run ./cmd/corpusgen", len(expectedPaths))
	}

	for _, expectedPath := range expectedPaths {
		expectedPath := expectedPath
		t.Run(strings.TrimSuffix(filepath.Base(expectedPath), ".json"), func(t *testing.T) {
			expectation := readXLSXExpectation(t, expectedPath)
			packagePath := filepath.Join(xlsxCorpusRoot, "generated", "packages", expectation.FixtureID+".xlsx")
			packageBytes, err := os.ReadFile(packagePath)
			if err != nil {
				t.Fatal(err)
			}

			workbook, extractErr := xlsxpatch.ExtractNativeWorkbookV1(packageBytes)
			if expectation.Outcome == "refused" {
				if extractErr == nil || workbook != nil {
					t.Fatalf("refused fixture produced workbook: %#v", workbook)
				}
				if got := classifyXLSXRefusal(extractErr); got != expectation.Refusal.Class {
					t.Fatalf("refusal class = %q, want %q: %v", got, expectation.Refusal.Class, extractErr)
				}
				if !strings.Contains(extractErr.Error(), expectation.Refusal.Contains) {
					t.Fatalf("refusal %q does not contain %q", extractErr, expectation.Refusal.Contains)
				}
				return
			}

			if extractErr != nil {
				t.Fatal(extractErr)
			}
			if issues := xlsxpatch.ValidateNativeWorkbookV1(workbook); len(issues) != 0 {
				t.Fatalf("extracted native contract is invalid: %+v", issues)
			}
			assertXLSXPackageProvenance(t, packageBytes, workbook)
			assertXLSXFixtureSemantics(t, expectation.FixtureID, workbook)

			actual, err := xlsxpatch.EncodeNativeWorkbookV1(workbook)
			if err != nil {
				t.Fatal(err)
			}
			var expected bytes.Buffer
			if err := json.Compact(&expected, expectation.Native); err != nil {
				t.Fatalf("compact expected native JSON: %v", err)
			}
			if !bytes.Equal(actual, expected.Bytes()) {
				t.Fatalf("native JSON drift\nactual: %s\nexpected: %s", actual, expected.Bytes())
			}

			decoded, err := xlsxpatch.DecodeNativeWorkbookV1(actual)
			if err != nil {
				t.Fatal(err)
			}
			roundTrip, err := xlsxpatch.EncodeNativeWorkbookV1(decoded)
			if err != nil || !bytes.Equal(roundTrip, actual) {
				t.Fatalf("native Encode -> Decode -> Encode drift: err=%v", err)
			}
			again, err := xlsxpatch.ExtractNativeWorkbookV1(packageBytes)
			if err != nil {
				t.Fatal(err)
			}
			againJSON, err := xlsxpatch.EncodeNativeWorkbookV1(again)
			if err != nil || !bytes.Equal(againJSON, actual) {
				t.Fatalf("repeated extraction is not deterministic: err=%v", err)
			}
		})
	}
}

func TestXLSXCorpusSpecsContainNoLegacyReconstructionAuthority(t *testing.T) {
	forbidden := []string{"mammoth", "luckyexcel", "domparser", "jszip", "innerhtml", "<html"}
	paths, err := filepath.Glob(filepath.Join(xlsxCorpusRoot, "specs", "xlsx", "*.json"))
	if err != nil {
		t.Fatal(err)
	}
	if len(paths) != 8 {
		t.Fatalf("XLSX spec count = %d, want 8", len(paths))
	}
	for _, path := range paths {
		data, err := os.ReadFile(path)
		if err != nil {
			t.Fatal(err)
		}
		lower := strings.ToLower(string(data))
		for _, token := range forbidden {
			if strings.Contains(lower, token) {
				t.Errorf("%s contains forbidden legacy reconstruction token %q", path, token)
			}
		}
		if !strings.Contains(string(data), `"license": "CC0-1.0"`) {
			t.Errorf("%s does not carry CC0-1.0 provenance", path)
		}
	}
}

func readXLSXExpectation(t *testing.T, path string) corpus.Expectation {
	t.Helper()
	data, err := os.ReadFile(path)
	if err != nil {
		t.Fatal(err)
	}
	var expectation corpus.Expectation
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	if err := decoder.Decode(&expectation); err != nil {
		t.Fatal(err)
	}
	if expectation.Protocol != corpus.ExpectationProtocol || expectation.Format != "xlsx" || expectation.FixtureID == "" {
		t.Fatalf("invalid XLSX expectation envelope: %+v", expectation)
	}
	return expectation
}

func classifyXLSXRefusal(err error) string {
	message := err.Error()
	switch {
	case strings.Contains(message, "compression-ratio limit"), strings.Contains(message, "resource bound"), strings.Contains(message, "exceeds"):
		return "resource-limit"
	case strings.Contains(message, "OPC parts"), strings.Contains(message, "ZIP"), strings.Contains(message, "relationship"), strings.Contains(message, "content type"):
		return "invalid-package"
	default:
		return "invalid-native-ooxml"
	}
}

func assertXLSXPackageProvenance(t *testing.T, packageBytes []byte, workbook *xlsxpatch.NativeWorkbookV1) {
	t.Helper()
	digest := sha256.Sum256(packageBytes)
	wantSHA := "sha256:" + hex.EncodeToString(digest[:])
	if workbook.Source.PackageSHA256 != wantSHA || workbook.Revision != "rev:"+strings.TrimPrefix(wantSHA, "sha256:") || workbook.Source.Authority != "exact-package-bytes" {
		t.Fatalf("package provenance mismatch: %+v", workbook.Source)
	}
	if workbook.Source.Dialect != "strict" && workbook.Source.Dialect != "transitional" {
		t.Fatalf("unexpected dialect %q", workbook.Source.Dialect)
	}

	reader, err := zip.NewReader(bytes.NewReader(packageBytes), int64(len(packageBytes)))
	if err != nil {
		t.Fatal(err)
	}
	parts := make(map[string][]byte, len(reader.File))
	for _, file := range reader.File {
		rc, err := file.Open()
		if err != nil {
			t.Fatal(err)
		}
		payload, readErr := io.ReadAll(rc)
		closeErr := rc.Close()
		if readErr != nil || closeErr != nil {
			t.Fatalf("read %s: %v / %v", file.Name, readErr, closeErr)
		}
		parts[file.Name] = payload
	}
	for _, passthrough := range workbook.PassthroughParts {
		payload, found := parts[passthrough.PartName]
		if !found || passthrough.ByteLength == nil || *passthrough.ByteLength != int64(len(payload)) || passthrough.Policy != "preserve-exact" {
			t.Fatalf("passthrough provenance mismatch: %+v", passthrough)
		}
		partDigest := sha256.Sum256(payload)
		if passthrough.SHA256 != "sha256:"+hex.EncodeToString(partDigest[:]) {
			t.Fatalf("passthrough digest mismatch for %s", passthrough.PartName)
		}
	}
}

func assertXLSXFixtureSemantics(t *testing.T, fixtureID string, workbook *xlsxpatch.NativeWorkbookV1) {
	t.Helper()
	switch fixtureID {
	case "xlsx-transitional-common":
		if workbook.Source.WorkbookPart != "xl/workbook.xml" || nativeCell(t, workbook, "C1").Value == nil || *nativeCell(t, workbook, "C1").Value.Lexical != "001.2300" || nativeCell(t, workbook, "G1").Formula.Cached == nil || *nativeCell(t, workbook, "G1").Formula.Cached.Lexical != "cached_x0020_text" || nativeCell(t, workbook, "B2").Value != nil || nativeCell(t, workbook, "C2").OOXMLType == nil || *nativeCell(t, workbook, "C2").OOXMLType != "n" {
			t.Fatal("Transitional lexical or absent-vs-explicit fidelity drift")
		}
	case "xlsx-strict-common":
		if workbook.Source.WorkbookPart != "Book/Workbook.xml" || workbook.Sheets[0].PartName != "Sheets/S1.xml" || workbook.Source.Dialect != "strict" {
			t.Fatalf("Strict actual part spelling drift: %+v / %+v", workbook.Source, workbook.Sheets[0])
		}
	case "xlsx-relational-preservation":
		if len(workbook.Sheets) != 3 ||
			workbook.Sheets[0].ID != "4" || workbook.Sheets[0].Name != "Inputs" || workbook.Sheets[0].Order != 0 || workbook.Sheets[0].State != "visible" || workbook.Sheets[0].PartName != "xl/worksheets/input.xml" ||
			workbook.Sheets[1].ID != "8" || workbook.Sheets[1].Name != "Hidden Data" || workbook.Sheets[1].Order != 1 || workbook.Sheets[1].State != "hidden" || workbook.Sheets[1].PartName != "xl/worksheets/hidden.xml" ||
			workbook.Sheets[2].ID != "12" || workbook.Sheets[2].Name != "Audit" || workbook.Sheets[2].Order != 2 || workbook.Sheets[2].State != "veryHidden" || workbook.Sheets[2].PartName != "xl/worksheets/audit.xml" {
			t.Fatalf("multi-sheet identity/order/state drift: %#v", workbook.Sheets)
		}
		formula := workbook.Sheets[0].Cells[2].Formula
		if formula == nil || formula.Text != `'Hidden Data'!A1*B1` || formula.Cached == nil || formula.Cached.Lexical == nil || *formula.Cached.Lexical != "3.1500" || formula.Cached.Storage != "formula-string" {
			t.Fatalf("cross-sheet formula/cache fidelity drift: %#v", formula)
		}
		hiddenValue := workbook.Sheets[1].Cells[0].Value
		if hiddenValue == nil || hiddenValue.Lexical == nil || *hiddenValue.Lexical != "00042" {
			t.Fatalf("hidden-sheet numeric lexical drift: %#v", hiddenValue)
		}
		unsupportedCounts := map[string]int{}
		for _, unsupported := range workbook.Unsupported {
			unsupportedCounts[unsupported.Code]++
			if unsupported.Preservation != "preserve-exact" {
				t.Fatalf("relationship/package diagnostic lost exact-preservation authority: %+v", unsupported)
			}
		}
		for code, count := range map[string]int{"EXTERNAL_RELATIONSHIP": 2, "EXTERNAL_LINK_CONTENT": 2, "HYPERLINKS": 1, "UNMODELED_WORKBOOK_FEATURE": 1} {
			if unsupportedCounts[code] != count {
				t.Errorf("unsupported %s count = %d, want %d", code, unsupportedCounts[code], count)
			}
		}
		wantPassthrough := []string{
			"customXml/_rels/item1.xml.rels",
			"customXml/item1.xml",
			"customXml/itemProps1.xml",
			"docProps/core.xml",
			"xl/externalLinks/_rels/externalLink1.xml.rels",
			"xl/externalLinks/externalLink1.xml",
			"xl/worksheets/_rels/input.xml.rels",
		}
		if len(workbook.PassthroughParts) != len(wantPassthrough) {
			t.Fatalf("passthrough relationship graph has %d parts, want %d: %#v", len(workbook.PassthroughParts), len(wantPassthrough), workbook.PassthroughParts)
		}
		for index, want := range wantPassthrough {
			part := workbook.PassthroughParts[index]
			if part.PartName != want || part.Policy != "preserve-exact" || part.ByteLength == nil || *part.ByteLength <= 0 || !strings.HasPrefix(part.SHA256, "sha256:") {
				t.Errorf("passthrough part %d = %#v, want exact %s evidence", index, part, want)
			}
		}
	case "xlsx-unsupported-preserve":
		sheet := &workbook.Sheets[0]
		if len(sheet.MergedRanges) != 1 || sheet.MergedRanges[0] != (xlsxpatch.NativeWorkbookMergedRangeV1{
			Ref: "A1:A2", Row: 0, Column: 0, EndRow: 1, EndColumn: 0, Editable: false,
		}) {
			t.Fatalf("merged-range projection drift: %#v", sheet.MergedRanges)
		}
		anchor, covered := nativeCell(t, workbook, "A1"), nativeCell(t, workbook, "A2")
		if anchor.Value == nil || covered.Value != nil || covered.Formula != nil || anchor.Editable || covered.Editable || anchor.StyleID != covered.StyleID {
			t.Fatalf("top-left/blank-covered/shared-style semantics drift: anchor=%#v covered=%#v", anchor, covered)
		}
		if len(sheet.Rows) != 1 || sheet.Rows[0].Row != 1 || !sheet.Rows[0].Hidden || len(sheet.Columns) != 1 || sheet.Columns[0].Column != 0 || sheet.Columns[0].EndColumn != 0 || !sheet.Columns[0].Hidden {
			t.Fatalf("hidden dimensions changed merged grid coordinates: rows=%#v columns=%#v", sheet.Rows, sheet.Columns)
		}
		codes := map[string]bool{}
		for _, unsupported := range workbook.Unsupported {
			codes[unsupported.Code] = true
			if unsupported.Preservation != "preserve-exact" {
				t.Fatalf("unsupported item lost exact preservation: %+v", unsupported)
			}
		}
		for _, code := range []string{"RICH_SHARED_STRING", "RICH_CELL_STRING", "FORMULA_SHARED", "FORMULA_ARRAY", "FORMULA_DATATABLE", "FORMULA_GROUP_RANGE", "MERGED_CELLS", "CONDITIONAL_FORMATTING", "EXTERNAL_RELATIONSHIP", "CHART_CONTENT", "OPAQUE_PACKAGE_PART"} {
			if !codes[code] {
				t.Errorf("missing unsupported diagnostic %s", code)
			}
		}
		if sheet.Editable || sheet.RefusalCode == nil || *sheet.RefusalCode != "FORMULA_GROUPS" || len(workbook.PassthroughParts) != 3 {
			t.Fatalf("unsupported scope was not explicitly refused/preserved: %+v", sheet)
		}
	default:
		t.Fatalf("unexpected accepted XLSX fixture %q", fixtureID)
	}
}

func nativeCell(t *testing.T, workbook *xlsxpatch.NativeWorkbookV1, ref string) *xlsxpatch.NativeWorkbookCellV1 {
	t.Helper()
	for sheetIndex := range workbook.Sheets {
		for cellIndex := range workbook.Sheets[sheetIndex].Cells {
			cell := &workbook.Sheets[sheetIndex].Cells[cellIndex]
			if cell.Ref == ref {
				return cell
			}
		}
	}
	t.Fatal(fmt.Sprintf("cell %s not found", ref))
	return nil
}
