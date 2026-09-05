package officecompat_test

import (
	"archive/zip"
	"bytes"
	"crypto/sha256"
	"encoding/json"
	"fmt"
	"os"
	"path/filepath"
	"reflect"
	"sort"
	"strings"
	"testing"
	"time"

	"github.com/injectinglabs/injoffice/go/docxpatch"
	"github.com/injectinglabs/injoffice/go/officecompat"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const performanceRepeats = 7

var performanceFixtureSHA256 = map[string]string{
	"docx": "47eff90d4080fef9dcc39d66b75f11ff17d052b7c9602b758ea4761e909ce209",
	"pptx": "09041bde998c7f18db24a7f9c9aae250fce53784265f928fa8c0fcc97b990ff4",
	"xlsx": "d5b0e821886fd43d7d484310e1a6c86ead6af813ce5c47bb2a2f33a31eea6bb2",
}

// These ceilings are intentionally generous regression guards, not speed
// claims. They were rounded to more than 2x the largest of repeated local
// Node 22/Go 1.23 arm64 runs documented in
// docs/qualification/native-office-performance-v1.json. Wall-clock samples
// are report-only because hosted-runner scheduling is not a correctness signal.
var allocationBudgets = map[string]float64{
	"docx/extract": 12_000,
	"docx/mutate":  30_000,
	"pptx/extract": 60_000,
	"pptx/mutate":  200_000,
	"xlsx/extract": 12_000,
	"xlsx/mutate":  60_000,
	"opc/inspect":  5_000,
}

type performanceFixture struct {
	format string
	bytes  []byte
}

type performanceSnapshot struct {
	format                string
	inventory             officecompat.Inventory
	canonicalNative       []byte
	objects               int
	pages                 int
	cells                 int
	shapes                int
	mutationBytes         []byte
	mutationBeforeObjects int
	mutationBeforeCells   int
	mutationBeforeShapes  int
	mutationObjects       int
	mutationCells         int
	mutationShapes        int
}

func loadPerformanceFixtures(tb testing.TB) []performanceFixture {
	tb.Helper()
	formats := []string{"docx", "pptx", "xlsx"}
	fixtures := make([]performanceFixture, 0, len(formats))
	for _, format := range formats {
		path := filepath.Join("corpus", "generated", "packages", format+"-transitional-common."+format)
		data, err := os.ReadFile(path)
		if err != nil {
			tb.Fatalf("read pinned %s performance fixture: %v", format, err)
		}
		digest := sha256.Sum256(data)
		if got := fmt.Sprintf("%x", digest[:]); got != performanceFixtureSHA256[format] {
			tb.Fatalf("pinned %s performance fixture digest = %s, want %s", format, got, performanceFixtureSHA256[format])
		}
		fixtures = append(fixtures, performanceFixture{format: format, bytes: data})
	}
	return fixtures
}

func TestNativeOfficePerformanceQualification(t *testing.T) {
	for _, fixture := range loadPerformanceFixtures(t) {
		fixture := fixture
		t.Run(fixture.format, func(t *testing.T) {
			first := snapshotPerformanceFixture(t, fixture)
			assertExpectedPerformanceCounts(t, first)
			for run := 1; run < performanceRepeats; run++ {
				next := snapshotPerformanceFixture(t, fixture)
				if !reflect.DeepEqual(first, next) {
					t.Fatalf("qualification run %d was nondeterministic\nfirst: %#v\nnext:  %#v", run+1, first, next)
				}
			}
		})
	}
}

func TestNativeOfficePerformanceAllocationBudgets(t *testing.T) {
	for _, fixture := range loadPerformanceFixtures(t) {
		fixture := fixture
		t.Run(fixture.format, func(t *testing.T) {
			extractName := fixture.format + "/extract"
			extractAllocs := testing.AllocsPerRun(5, func() {
				if _, err := extractPerformanceNative(fixture); err != nil {
					panic(err)
				}
			})
			assertAllocationBudget(t, extractName, extractAllocs)

			mutateAllocs := testing.AllocsPerRun(3, func() {
				if _, err := mutatePerformanceNative(fixture); err != nil {
					panic(err)
				}
			})
			assertAllocationBudget(t, fixture.format+"/mutate", mutateAllocs)
		})
	}

	synthetic := generatedOPCFixture(t, 48, 32)
	opcLimits := tightPerformanceLimits(t, synthetic)
	opcAllocs := testing.AllocsPerRun(7, func() {
		if _, err := officecompat.InspectWithLimits(synthetic, opcLimits); err != nil {
			panic(err)
		}
	})
	assertAllocationBudget(t, "opc/inspect", opcAllocs)
}

func TestPerformanceHarnessBudgetSelfTest(t *testing.T) {
	if err := allocationBudgetError("self-test", 10, 10); err != nil {
		t.Fatalf("exact budget should pass: %v", err)
	}
	if err := allocationBudgetError("self-test", 10.01, 10); err == nil || !strings.Contains(err.Error(), "algorithmic/resource regression") {
		t.Fatalf("over-budget sample was not rejected: %v", err)
	}
	if err := allocationBudgetError("self-test", -1, 10); err == nil {
		t.Fatal("negative allocation sample was accepted")
	}
	if err := allocationBudgetError("self-test", 1, 0); err == nil {
		t.Fatal("zero allocation budget was accepted")
	}
}

func TestPerformanceBudgetArtifactMatchesHarness(t *testing.T) {
	data, err := os.ReadFile(filepath.Join("..", "..", "docs", "qualification", "native-office-performance-v1.json"))
	if err != nil {
		t.Fatal(err)
	}
	var artifact struct {
		Protocol string `json:"protocol"`
		Baseline struct {
			Fixtures map[string]struct {
				SHA256 string `json:"sha256"`
			} `json:"fixtures"`
		} `json:"baseline"`
		Allocations map[string]float64 `json:"go_allocation_ceilings_per_op"`
	}
	if err := json.Unmarshal(data, &artifact); err != nil {
		t.Fatal(err)
	}
	if artifact.Protocol != "injoffice.office.performance-budgets/v1" || !reflect.DeepEqual(artifact.Allocations, allocationBudgets) {
		t.Fatalf("performance budget artifact drifted from executable harness: protocol=%q allocations=%v", artifact.Protocol, artifact.Allocations)
	}
	for format, digest := range performanceFixtureSHA256 {
		if artifact.Baseline.Fixtures[format].SHA256 != digest {
			t.Fatalf("%s fixture pin drifted between artifact and harness", format)
		}
	}
}

func TestPerformanceFixturesNearAndOverResourceLimits(t *testing.T) {
	for _, fixture := range loadPerformanceFixtures(t) {
		fixture := fixture
		t.Run(fixture.format, func(t *testing.T) {
			limits := tightPerformanceLimits(t, fixture.bytes)
			inventory, err := officecompat.InspectWithLimits(fixture.bytes, limits)
			if err != nil || len(inventory.Parts) != limits.MaxParts {
				t.Fatalf("near-limit fixture was refused or lost parts: parts=%d limit=%d err=%v", len(inventory.Parts), limits.MaxParts, err)
			}
			over := limits
			over.MaxParts--
			if _, err := officecompat.InspectWithLimits(fixture.bytes, over); err == nil || !strings.Contains(err.Error(), "entries") {
				t.Fatalf("over-limit fixture did not fail at the part boundary: %v", err)
			}
		})
	}

	for _, parts := range []int{64, 65} {
		data := generatedOPCFixture(t, parts, 24)
		limits := tightPerformanceLimits(t, data)
		limits.MaxParts = 64
		_, err := officecompat.InspectWithLimits(data, limits)
		if parts == 64 && err != nil {
			t.Fatalf("generated near-limit OPC fixture was refused: %v", err)
		}
		if parts == 65 && (err == nil || !strings.Contains(err.Error(), "entries")) {
			t.Fatalf("generated over-limit OPC fixture was accepted: %v", err)
		}
	}
}

func TestPerformanceSyntheticComplexityGuard(t *testing.T) {
	// Doubling flat OPC parts should stay comfortably below a 3x allocation
	// multiplier. The absolute ceiling also catches a bypass hidden by ratios.
	small := generatedOPCFixture(t, 32, 64)
	large := generatedOPCFixture(t, 64, 64)
	measure := func(data []byte) float64 {
		limits := tightPerformanceLimits(t, data)
		return testing.AllocsPerRun(7, func() {
			if _, err := officecompat.InspectWithLimits(data, limits); err != nil {
				panic(err)
			}
		})
	}
	smallAllocs, largeAllocs := measure(small), measure(large)
	if largeAllocs > smallAllocs*3 || largeAllocs > allocationBudgets["opc/inspect"] {
		t.Fatalf("flat OPC inspection allocation growth is explosive: 32 parts=%.0f, 64 parts=%.0f", smallAllocs, largeAllocs)
	}
}

func snapshotPerformanceFixture(tb testing.TB, fixture performanceFixture) performanceSnapshot {
	tb.Helper()
	inventory, err := officecompat.InspectWithLimits(fixture.bytes, tightPerformanceLimits(tb, fixture.bytes))
	if err != nil {
		tb.Fatalf("inspect %s: %v", fixture.format, err)
	}
	native, err := extractPerformanceNative(fixture)
	if err != nil {
		tb.Fatalf("extract %s: %v", fixture.format, err)
	}
	mutated, err := mutatePerformanceNative(fixture)
	if err != nil {
		tb.Fatalf("mutate/reopen %s: %v", fixture.format, err)
	}
	native.mutationBytes = mutated.mutationBytes
	native.mutationBeforeObjects = mutated.mutationBeforeObjects
	native.mutationBeforeCells = mutated.mutationBeforeCells
	native.mutationBeforeShapes = mutated.mutationBeforeShapes
	native.mutationObjects = mutated.mutationObjects
	native.mutationCells = mutated.mutationCells
	native.mutationShapes = mutated.mutationShapes
	native.inventory = inventory
	return native
}

func extractPerformanceNative(fixture performanceFixture) (performanceSnapshot, error) {
	switch fixture.format {
	case "xlsx":
		workbook, err := xlsxpatch.ExtractNativeWorkbookV1(fixture.bytes)
		if err != nil {
			return performanceSnapshot{}, err
		}
		encoded, err := json.Marshal(workbook)
		if err != nil {
			return performanceSnapshot{}, err
		}
		cells := 0
		for _, sheet := range workbook.Sheets {
			cells += len(sheet.Cells)
		}
		return performanceSnapshot{format: fixture.format, canonicalNative: encoded, objects: len(workbook.Sheets) + cells + len(workbook.Styles), cells: cells}, nil
	case "pptx":
		deck, err := pptxpatch.ExtractNativePPTX(fixture.bytes, performancePPTXOptions(nil))
		if err != nil {
			return performanceSnapshot{}, err
		}
		encoded, err := pptxpatch.MarshalNativePPTXJSON(deck)
		if err != nil {
			return performanceSnapshot{}, err
		}
		elements, shapes := countPPTXElements(deck)
		return performanceSnapshot{format: fixture.format, canonicalNative: encoded, objects: len(deck.Slides) + elements, pages: len(deck.Slides), shapes: shapes}, nil
	case "docx":
		document, err := docxpatch.ExtractNativeDocumentV1(fixture.bytes)
		if err != nil {
			return performanceSnapshot{}, err
		}
		encoded, err := docxpatch.EncodeNativeDocumentV1(document)
		if err != nil {
			return performanceSnapshot{}, err
		}
		objects, cells := countDOCXObjects(document)
		return performanceSnapshot{format: fixture.format, canonicalNative: encoded, objects: objects, pages: len(document.Sections), cells: cells}, nil
	default:
		return performanceSnapshot{}, fmt.Errorf("unsupported performance format %q", fixture.format)
	}
}

func mutatePerformanceNative(fixture performanceFixture) (performanceSnapshot, error) {
	switch fixture.format {
	case "xlsx":
		mutationSource, err := os.ReadFile(filepath.Join("corpus", "generated", "packages", "xlsx-relational-preservation.xlsx"))
		if err != nil {
			return performanceSnapshot{}, err
		}
		mutationDigest := sha256.Sum256(mutationSource)
		if got := fmt.Sprintf("%x", mutationDigest[:]); got != "6fa5d6029e501d23228ccb1a7e0581c87fa1afebda4ac3445fbfc96e29311617" {
			return performanceSnapshot{}, fmt.Errorf("xlsx mutation fixture digest = %s", got)
		}
		before, err := xlsxpatch.ExtractNativeWorkbookV1(mutationSource)
		if err != nil {
			return performanceSnapshot{}, err
		}
		beforeCells := 0
		for _, sheet := range before.Sheets {
			beforeCells += len(sheet.Cells)
		}
		result, err := xlsxpatch.ApplyNativeWorkbookMutationTransactionV1(mutationSource, xlsxpatch.NativeWorkbookMutationTransactionV1{
			ExpectedRevision: before.Revision,
			Cells:            []xlsxpatch.CellMutation{{OperationID: "qualification-cell", SheetID: before.Sheets[0].ID, Kind: xlsxpatch.CellSetValue, Cell: xlsxpatch.CellRef{Row: 0, Column: 1}, Value: 0.125}},
		})
		if err != nil {
			return performanceSnapshot{}, err
		}
		cells := 0
		for _, sheet := range result.Workbook.Sheets {
			cells += len(sheet.Cells)
		}
		return performanceSnapshot{
			mutationBytes:         result.Package,
			mutationBeforeObjects: len(before.Sheets) + beforeCells + len(before.Styles), mutationBeforeCells: beforeCells,
			mutationObjects: len(result.Workbook.Sheets) + cells + len(result.Workbook.Styles), mutationCells: cells,
		}, nil
	case "pptx":
		before, err := pptxpatch.ExtractNativePPTX(fixture.bytes, performancePPTXOptions(nil))
		if err != nil {
			return performanceSnapshot{}, err
		}
		target := before.Slides[0].Elements[0]
		paragraphs := append([]pptxpatch.NativeParagraph(nil), (*target.Paragraphs)...)
		paragraphs[0].Runs = append([]pptxpatch.NativeTextRun(nil), paragraphs[0].Runs...)
		text := " Qualified "
		paragraphs[0].Runs[0].Text = &text
		produced, err := pptxpatch.ApplyNativePPTXMutations(fixture.bytes, pptxpatch.NativePPTXMutationRequest{
			ExpectedSourceRevision: *before.SourceRevision,
			Operations:             []pptxpatch.NativePPTXMutation{{OperationID: "qualification-text", Kind: pptxpatch.NativePPTXReplaceText, ElementID: target.ID, ExpectedFingerprintSHA256: target.Source.FingerprintSHA256, Paragraphs: &paragraphs}},
		})
		if err != nil {
			return performanceSnapshot{}, err
		}
		afterOptions := performancePPTXOptions(&before)
		after, err := pptxpatch.ExtractNativePPTX(produced, afterOptions)
		if err != nil {
			return performanceSnapshot{}, err
		}
		elements, shapes := countPPTXElements(after)
		beforeElements, beforeShapes := countPPTXElements(before)
		return performanceSnapshot{mutationBytes: produced, mutationBeforeObjects: len(before.Slides) + beforeElements, mutationBeforeShapes: beforeShapes, mutationObjects: len(after.Slides) + elements, mutationShapes: shapes}, nil
	case "docx":
		before, err := docxpatch.ExtractNativeDocumentV1(fixture.bytes)
		if err != nil {
			return performanceSnapshot{}, err
		}
		target := before.Body.Blocks[1].Table.Rows[0].Cells[0].Paragraphs[0].Runs[0]
		result, err := docxpatch.ApplyNativeTextMutationsV1(fixture.bytes, before.Source.PackageSHA256, []docxpatch.NativeDOCXTextMutationV1{{
			TargetKind: "run", TargetID: target.ID, ExpectedXMLSHA256: target.Anchor.XMLSHA256, Text: "Qualified",
		}})
		if err != nil {
			return performanceSnapshot{}, err
		}
		objects, cells := countDOCXObjects(result.Document)
		beforeObjects, beforeCells := countDOCXObjects(before)
		return performanceSnapshot{mutationBytes: result.Package, mutationBeforeObjects: beforeObjects, mutationBeforeCells: beforeCells, mutationObjects: objects, mutationCells: cells}, nil
	default:
		return performanceSnapshot{}, fmt.Errorf("unsupported performance format %q", fixture.format)
	}
}

func countPPTXElements(deck pptxpatch.NativePPTXDeck) (elements, shapes int) {
	var visit func([]pptxpatch.NativeElement)
	visit = func(items []pptxpatch.NativeElement) {
		for _, item := range items {
			elements++
			if item.Kind == pptxpatch.NativeElementKindShape {
				shapes++
			}
			if len(item.Children) != 0 {
				visit(item.Children)
			}
		}
	}
	for _, slide := range deck.Slides {
		visit(slide.Elements)
	}
	return elements, shapes
}

func countDOCXObjects(document *docxpatch.NativeDocumentV1) (objects, cells int) {
	objects = len(document.Sections)
	for _, block := range document.Body.Blocks {
		objects++
		if block.Paragraph != nil {
			objects += len(block.Paragraph.Runs)
		}
		if block.Table != nil {
			for _, row := range block.Table.Rows {
				objects++
				for _, cell := range row.Cells {
					cells++
					objects++
					for _, paragraph := range cell.Paragraphs {
						objects += 1 + len(paragraph.Runs)
					}
				}
			}
		}
	}
	return objects, cells
}

func assertExpectedPerformanceCounts(t *testing.T, snapshot performanceSnapshot) {
	t.Helper()
	want := map[string]struct{ objects, pages, cells, shapes int }{
		"docx": {objects: 9, pages: 1, cells: 1, shapes: 0},
		"pptx": {objects: 6, pages: 1, cells: 0, shapes: 4},
		"xlsx": {objects: 13, pages: 0, cells: 10, shapes: 0},
	}[snapshot.format]
	if snapshot.objects != want.objects || snapshot.pages != want.pages || snapshot.cells != want.cells || snapshot.shapes != want.shapes {
		t.Fatalf("native output cardinality changed: got objects/pages/cells/shapes=%d/%d/%d/%d, want %d/%d/%d/%d", snapshot.objects, snapshot.pages, snapshot.cells, snapshot.shapes, want.objects, want.pages, want.cells, want.shapes)
	}
	if snapshot.mutationObjects != snapshot.mutationBeforeObjects || snapshot.mutationCells != snapshot.mutationBeforeCells || snapshot.mutationShapes != snapshot.mutationBeforeShapes {
		t.Fatalf("mutation round trip silently lost objects: before objects/cells/shapes=%d/%d/%d, after=%d/%d/%d", snapshot.mutationBeforeObjects, snapshot.mutationBeforeCells, snapshot.mutationBeforeShapes, snapshot.mutationObjects, snapshot.mutationCells, snapshot.mutationShapes)
	}
}

func performancePPTXOptions(previous *pptxpatch.NativePPTXDeck) pptxpatch.NativePPTXExtractOptions {
	return pptxpatch.NativePPTXExtractOptions{
		Previous: previous,
		TokenFactory: pptxpatch.NativePassthroughTokenFactoryFunc(func(request pptxpatch.NativePassthroughTokenRequest) (string, error) {
			digest := sha256.Sum256([]byte(request.OwnerPart + "\x00" + request.ObjectID + "\x00" + request.Reason))
			return fmt.Sprintf("qualification-%x", digest[:12]), nil
		}),
	}
}

func assertAllocationBudget(t *testing.T, name string, measured float64) {
	t.Helper()
	budget, ok := allocationBudgets[name]
	if !ok {
		t.Fatalf("missing allocation budget for %s", name)
	}
	if err := allocationBudgetError(name, measured, budget); err != nil {
		t.Fatal(err)
	}
}

func allocationBudgetError(name string, measured, budget float64) error {
	if measured < 0 || budget <= 0 {
		return fmt.Errorf("invalid allocation measurement for %s: measured=%.2f budget=%.2f", name, measured, budget)
	}
	if measured > budget {
		return fmt.Errorf("algorithmic/resource regression in %s: %.0f allocations/op exceeds generous %.0f budget", name, measured, budget)
	}
	return nil
}

func tightPerformanceLimits(tb testing.TB, data []byte) officecompat.Limits {
	tb.Helper()
	inventory, err := officecompat.Inspect(data)
	if err != nil {
		tb.Fatalf("derive tight performance limits: %v", err)
	}
	var expanded, largest uint64
	for _, part := range inventory.Parts {
		expanded += part.Size
		if part.Size > largest {
			largest = part.Size
		}
	}
	return officecompat.Limits{
		MaxPackageBytes: uint64(len(data)), MaxParts: len(inventory.Parts), MaxPartBytes: largest,
		MaxExpandedBytes: expanded, MaxCompressionRatio: officecompat.MaxCompressionRatio,
		CompressionRatioSlack: 1 << 20, MaxXMLBytes: largest, MaxXMLDepth: 128,
		MaxXMLTokens: 100_000, MaxXMLAttributes: 100_000,
	}
}

func generatedOPCFixture(tb testing.TB, parts, payloadBytes int) []byte {
	tb.Helper()
	if parts < 2 {
		tb.Fatal("generated OPC fixture needs root parts")
	}
	entries := map[string]string{
		"[Content_Types].xml": `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/></Types>`,
		"_rels/.rels":         `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`,
	}
	for index := 2; index < parts; index++ {
		entries[fmt.Sprintf("parts/p%05d.bin", index)] = strings.Repeat("x", payloadBytes)
	}
	return buildQualificationZip(tb, entries)
}

func BenchmarkNativeOfficePerformanceQualification(b *testing.B) {
	for _, fixture := range loadPerformanceFixtures(b) {
		fixture := fixture
		b.Run(fixture.format+"/extract", func(b *testing.B) {
			b.ReportAllocs()
			b.SetBytes(int64(len(fixture.bytes)))
			for i := 0; i < b.N; i++ {
				if _, err := extractPerformanceNative(fixture); err != nil {
					b.Fatal(err)
				}
			}
		})
		b.Run(fixture.format+"/mutate", func(b *testing.B) {
			b.ReportAllocs()
			b.SetBytes(int64(len(fixture.bytes)))
			for i := 0; i < b.N; i++ {
				if _, err := mutatePerformanceNative(fixture); err != nil {
					b.Fatal(err)
				}
			}
		})
	}
}

// buildQualificationZip is shared with the structural tests in this external
// test package. Keep a local wrapper so the performance fixtures remain
// generated, compact, and independent of ZIP timestamps.
func buildQualificationZip(tb testing.TB, entries map[string]string) []byte {
	tb.Helper()
	names := make([]string, 0, len(entries))
	for name := range entries {
		names = append(names, name)
	}
	sort.Strings(names)
	var output bytes.Buffer
	writer := zip.NewWriter(&output)
	for _, name := range names {
		header := &zip.FileHeader{Name: name, Method: zip.Store}
		header.SetModTime(time.Date(1980, 1, 1, 0, 0, 0, 0, time.UTC))
		header.SetMode(0o644)
		part, err := writer.CreateHeader(header)
		if err != nil {
			tb.Fatalf("create generated OPC part %s: %v", name, err)
		}
		if _, err := part.Write([]byte(entries[name])); err != nil {
			tb.Fatalf("write generated OPC part %s: %v", name, err)
		}
	}
	if err := writer.Close(); err != nil {
		tb.Fatalf("close generated OPC fixture: %v", err)
	}
	return output.Bytes()
}
