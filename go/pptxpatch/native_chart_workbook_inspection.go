package pptxpatch

import (
	"encoding/base64"
	"encoding/xml"
	"fmt"
)

const nativeChartWorkbookInspectionMaxRecords = 64
const nativeChartWorkbookInspectionMaxResources = 8
const nativeChartWorkbookInspectionMaxBytes = 16 * 1024 * 1024

// Inspection is a read-only source projection. It contains no mutation tokens.
type NativePPTXChartWorkbookInspection struct {
	Protocol       string                             `json:"protocol"`
	PackageSHA256  string                             `json:"package_sha256"`
	SourceRevision string                             `json:"source_revision"`
	Charts         []NativePPTXInspectedWorkbookChart `json:"charts"`
	Workbooks      []NativePPTXInspectedChartWorkbook `json:"workbooks"`
	Omissions      []NativePPTXTableOmission          `json:"omissions"`
}
type NativePPTXInspectedChartWorkbook struct {
	Part        string `json:"part"`
	SHA256      string `json:"sha256"`
	ByteLength  int64  `json:"byteLength"`
	BytesBase64 string `json:"bytesBase64"`
}
type NativePPTXInspectedWorkbookChart struct {
	SlideID     string `json:"slide_id"`
	SlideIndex  int    `json:"slide_index"`
	SlidePart   string `json:"slide_part"`
	SlideSHA256 string `json:"slide_sha256"`
	ElementID   string `json:"element_id"`
	ObjectID    string `json:"object_id"`
	// FrameSHA256 is the digest of the exact original p:graphicFrame XML bytes.
	FrameSHA256         string                         `json:"frame_sha256"`
	ChartRelationshipID string                         `json:"chart_relationship_id"`
	ChartPart           string                         `json:"chart_part"`
	ChartSHA256         string                         `json:"chart_sha256"`
	Source              NativePPTXWorkbookChartSource  `json:"source"`
	Workbook            NativePPTXChartWorkbookBinding `json:"workbook"`
}
type NativePPTXWorkbookChartSource struct {
	BubbleScale     *int64                          `json:"bubbleScale,omitempty"`
	SizeRepresents  string                          `json:"sizeRepresents,omitempty"`
	Family          string                          `json:"family"`
	BarDirection    string                          `json:"barDirection,omitempty"`
	GapWidth        *int64                          `json:"gapWidth,omitempty"`
	XAxis           NativeLiteralBarAxis            `json:"xAxis"`
	YAxis           NativeLiteralBarAxis            `json:"yAxis"`
	Series          []NativePPTXWorkbookChartSeries `json:"series"`
	PlotVisibleOnly bool                            `json:"plotVisibleOnly"`
	DispBlanksAs    *string                         `json:"dispBlanksAs,omitempty"`
}
type NativePPTXWorkbookChartSeries struct {
	Index             int64                             `json:"index"`
	Order             int64                             `json:"order"`
	Title             *string                           `json:"title,omitempty"`
	TitleReference    *NativePPTXChartWorkbookReference `json:"titleReference,omitempty"`
	CategoryReference *NativePPTXChartWorkbookReference `json:"categoryReference,omitempty"`
	XReference        *NativePPTXChartWorkbookReference `json:"xReference,omitempty"`
	SizeReference     *NativePPTXChartWorkbookReference `json:"sizeReference,omitempty"`
	ValueReference    *NativePPTXChartWorkbookReference `json:"valueReference"`
	Colors            []string                          `json:"colors,omitempty"`
	Color             string                            `json:"color,omitempty"`
	WidthEMU          *int64                            `json:"widthEmu,omitempty"`
}
type NativePPTXChartWorkbookReference struct {
	Kind         string                       `json:"kind"`
	Formula      string                       `json:"formula"`
	Range        NativePPTXChartWorkbookRange `json:"range"`
	CachePresent bool                         `json:"cachePresent"`
}
type NativePPTXChartWorkbookRange struct {
	Sheet       string `json:"sheet"`
	StartRow    int64  `json:"startRow"`
	StartColumn int64  `json:"startColumn"`
	EndRow      int64  `json:"endRow"`
	EndColumn   int64  `json:"endColumn"`
	Count       int64  `json:"count"`
}
type NativePPTXChartWorkbookBinding struct {
	RelationshipID string `json:"relationshipId"`
	Part           string `json:"part"`
	SHA256         string `json:"sha256"`
	ByteLength     int64  `json:"byteLength"`
	AutoUpdate     *bool  `json:"autoUpdate,omitempty"`
}

func nativeWorkbookPublicReference(ref *nativeChartReference) *NativePPTXChartWorkbookReference {
	if ref == nil {
		return nil
	}
	r := ref.Range
	return &NativePPTXChartWorkbookReference{Kind: ref.Kind, Formula: ref.Formula, CachePresent: ref.CachePresent, Range: NativePPTXChartWorkbookRange{Sheet: r.Sheet, StartRow: r.StartRow, StartColumn: r.StartColumn, EndRow: r.EndRow, EndColumn: r.EndColumn, Count: r.Count}}
}
func nativeWorkbookPublicSource(source *nativeChartWorkbookSource) NativePPTXWorkbookChartSource {
	family := map[string]string{"barChart": "bar", "lineChart": "line", "scatterChart": "scatter", "bubbleChart": "bubble"}[source.Family]
	result := NativePPTXWorkbookChartSource{Family: family, BubbleScale: source.BubbleScale, SizeRepresents: source.SizeRepresents, Series: []NativePPTXWorkbookChartSeries{}, DispBlanksAs: source.DispBlanksAs}
	result.XAxis = nativeLiteralChartAxis(source.XAxis, source.XAxis.Min != "")
	result.YAxis = nativeLiteralChartAxis(source.YAxis, source.YAxis.Min != "")
	if family == "bar" {
		result.BarDirection = source.Direction
		result.GapWidth = &source.GapWidth
	}
	for _, s := range source.Series {
		out := NativePPTXWorkbookChartSeries{Index: s.Index, Order: s.Order, Title: s.Title, TitleReference: nativeWorkbookPublicReference(s.TitleReference), CategoryReference: nativeWorkbookPublicReference(s.CategoryReference), XReference: nativeWorkbookPublicReference(s.XReference), ValueReference: nativeWorkbookPublicReference(s.ValueReference), SizeReference: nativeWorkbookPublicReference(s.SizeReference), Colors: s.Colors, Color: s.Color}
		if family == "line" || family == "scatter" {
			out.WidthEMU = &s.Width
		}
		result.Series = append(result.Series, out)
	}
	return result
}

// InspectNativePPTXChartWorkbooks reuses strict PPTX extraction and relationship
// routing. Only actual projected native chart frames are candidates. Refused
// source remains in the original package; this does not turn inspection into
// authoring authority or parse the embedded SpreadsheetML.
func InspectNativePPTXChartWorkbooks(data []byte) (NativePPTXChartWorkbookInspection, error) {
	result := NativePPTXChartWorkbookInspection{Protocol: "pptx-chart-workbook-inspection-v1", Charts: []NativePPTXInspectedWorkbookChart{}, Workbooks: []NativePPTXInspectedChartWorkbook{}, Omissions: []NativePPTXTableOmission{}}
	if len(data) == 0 || len(data) > NativePPTXMaxPackageBytes {
		return result, fmt.Errorf("chart workbook inspection: invalid package size")
	}
	result.PackageSHA256 = nativeSHA256(data)
	resources := map[string]bool{}
	totalBytes := 0
	_, err := extractNativePPTXWithInspection(data, NativePPTXExtractOptions{}, func(extractor *nativeExtractor, deck NativePPTXDeck) error {
		result.SourceRevision = *deck.SourceRevision
		for slideIndex, slide := range deck.Slides {
			if slide.Source == nil {
				return fmt.Errorf("chart workbook inspection: missing slide source")
			}
			slideBytes := extractor.pkg.parts[slide.Source.PartName]
			slideRoot, e := parseNativeXML(slideBytes, slide.Source.PartName)
			if e != nil {
				return e
			}
			dialect, e := nativeDialectForPresentation(xml.Name{Space: slideRoot.Name.Space, Local: "presentation"})
			if e != nil {
				return e
			}
			var walk func([]NativeElement) error
			walk = func(elements []NativeElement) error {
				for _, element := range elements {
					if element.Kind == NativeElementKindGroup {
						if e := walk(element.Children); e != nil {
							return e
						}
						continue
					}
					if element.Kind != NativeElementKindChart || element.Chart == nil {
						continue
					}
					if len(result.Charts)+len(result.Omissions) >= nativeChartWorkbookInspectionMaxRecords {
						return fmt.Errorf("chart workbook inspection: chart record budget exceeded")
					}
					if element.Source == nil {
						return fmt.Errorf("chart workbook inspection: missing chart frame source")
					}
					chart := element.Chart
					payload := extractor.pkg.parts[chart.ChartPart]
					source := extractNativeChartWorkbookSource(payload, chart.ChartPart, dialect)
					omit := func(reason string) {
						result.Omissions = append(result.Omissions, NativePPTXTableOmission{SlideID: slide.ID, ObjectID: element.Source.ObjectID, Reason: reason})
					}
					if source == nil {
						omit("Chart source is outside the explicit embedded-workbook profile.")
						continue
					}
					binding, ok, e := extractor.extractNativeChartWorkbookBinding(source.ExternalData, chart.ChartPart, dialect)
					if e != nil {
						return e
					}
					if !ok {
						omit("Chart externalData does not bind a qualified local XLSX package.")
						continue
					}
					if !resources[binding.Part] {
						if len(resources) >= nativeChartWorkbookInspectionMaxResources {
							return fmt.Errorf("chart workbook inspection: resource count budget exceeded")
						}
						bytes := extractor.pkg.parts[binding.Part]
						if len(bytes) > nativeChartWorkbookInspectionMaxBytes-totalBytes {
							return fmt.Errorf("chart workbook inspection: aggregate workbook byte budget exceeded")
						}
						totalBytes += len(bytes)
						resources[binding.Part] = true
						result.Workbooks = append(result.Workbooks, NativePPTXInspectedChartWorkbook{Part: binding.Part, SHA256: binding.SHA256, ByteLength: binding.ByteLength, BytesBase64: base64.StdEncoding.EncodeToString(bytes)})
					}
					result.Charts = append(result.Charts, NativePPTXInspectedWorkbookChart{SlideID: slide.ID, SlideIndex: slideIndex, SlidePart: slide.Source.PartName, SlideSHA256: nativeSHA256(slideBytes), ElementID: element.ID, ObjectID: element.Source.ObjectID, FrameSHA256: element.Source.FingerprintSHA256, ChartRelationshipID: chart.RelationshipID, ChartPart: chart.ChartPart, ChartSHA256: nativeSHA256(payload), Source: nativeWorkbookPublicSource(source), Workbook: NativePPTXChartWorkbookBinding{RelationshipID: binding.RelationshipID, Part: binding.Part, SHA256: binding.SHA256, ByteLength: binding.ByteLength, AutoUpdate: binding.AutoUpdate}})
				}
				return nil
			}
			if e := walk(slide.Elements); e != nil {
				return e
			}
		}
		return nil
	})
	return result, err
}
