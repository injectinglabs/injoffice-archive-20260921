package pptxpatch

import (
	"encoding/xml"
	"errors"
	"fmt"
)

const nativeTableInspectionMaxTables = 256
const nativeTableInspectionMaxCells = 4096
const nativeTableInspectionMaxText = 65536

// NativePPTXTableInspection contains plain source content, never HTML, rendered
// Office styling, mutation anchors, or published preservation capabilities.
type NativePPTXTableInspection struct {
	Protocol       string                     `json:"protocol"`
	PackageSHA256  string                     `json:"package_sha256"`
	SourceRevision string                     `json:"source_revision"`
	Tables         []NativePPTXInspectedTable `json:"tables"`
	Omissions      []NativePPTXTableOmission  `json:"omissions"`
}

type NativePPTXInspectionRect struct {
	X      int64 `json:"x"`
	Y      int64 `json:"y"`
	Width  int64 `json:"width"`
	Height int64 `json:"height"`
}

type NativePPTXInspectedCell struct {
	Row        int                      `json:"row"`
	Column     int                      `json:"column"`
	Rect       NativePPTXInspectionRect `json:"rect"`
	Paragraphs []string                 `json:"paragraphs"`
}

type NativePPTXInspectedTable struct {
	SlideID           string                    `json:"slide_id"`
	SlideIndex        int                       `json:"slide_index"`
	PartName          string                    `json:"part_name"`
	PartSHA256        string                    `json:"part_sha256"`
	SlideSourceSHA256 string                    `json:"slide_source_sha256"`
	ObjectID          string                    `json:"object_id"`
	SourceSHA256      string                    `json:"source_sha256"`
	Rect              NativePPTXInspectionRect  `json:"rect"`
	Cells             []NativePPTXInspectedCell `json:"cells"`
	Warnings          []string                  `json:"warnings"`
}

type NativePPTXTableOmission struct {
	SlideID  string `json:"slide_id"`
	ObjectID string `json:"object_id"`
	Reason   string `json:"reason"`
}

// InspectNativePPTXTables performs the same bounded strict package extraction
// before inspecting top-level table source. It does not make refused content
// editable, resolve external resources, or turn inspection into a native deck.
func InspectNativePPTXTables(data []byte) (NativePPTXTableInspection, error) {
	if len(data) == 0 || len(data) > NativePPTXMaxPackageBytes {
		return NativePPTXTableInspection{}, fmt.Errorf("table inspection: invalid package size")
	}
	result := NativePPTXTableInspection{Protocol: "pptx-table-content-inspection-v1", PackageSHA256: nativeSHA256(data), Tables: []NativePPTXInspectedTable{}, Omissions: []NativePPTXTableOmission{}}
	reserve := func() error {
		if len(result.Tables)+len(result.Omissions) >= nativeTableInspectionMaxTables {
			return &nativeInspectionBudgetError{"table/omission"}
		}
		return nil
	}
	_, err := extractNativePPTXWithInspection(data, NativePPTXExtractOptions{}, func(extractor *nativeExtractor, deck NativePPTXDeck) error {
		result.SourceRevision = *deck.SourceRevision
		budget := &nativeInspectionBudget{}
		for slideIndex, slide := range deck.Slides {
			if slide.Source == nil {
				return fmt.Errorf("table inspection: missing source slide")
			}
			part := slide.Source.PartName
			payload := extractor.pkg.parts[part]
			root, err := parseNativeXML(payload, part)
			if err != nil {
				return err
			}
			if hasNativeSemanticAttrs(root) {
				if err := reserve(); err != nil {
					return err
				}
				result.Omissions = append(result.Omissions, NativePPTXTableOmission{SlideID: slide.ID, ObjectID: "slide", Reason: "Slide visibility or root source properties are outside the table inspection profile."})
				continue
			}
			d := nativeExtractDialect{presentation: root.Name.Space, drawing: nsDrawingTransitional}
			if root.Name.Space == nsPresentationStrict {
				d.drawing = nsDrawingStrict
			}
			common, err := nativeSingleton(root, d.presentation, "cSld", true)
			if err != nil {
				return err
			}
			tree, err := nativeSingleton(common, d.presentation, "spTree", true)
			if err != nil {
				return err
			}
			if hasNativeSemanticAttrs(common) || hasNativeSemanticAttrs(tree) {
				if err := reserve(); err != nil {
					return err
				}
				result.Omissions = append(result.Omissions, NativePPTXTableOmission{SlideID: slide.ID, ObjectID: "slide", Reason: "Common-slide source properties are outside the table inspection profile."})
				continue
			}
			rootGroup, err := nativeSingleton(tree, d.presentation, "grpSpPr", true)
			if err != nil {
				return err
			}
			if !inspectTableRootGroup(rootGroup, d) {
				if err := reserve(); err != nil {
					return err
				}
				result.Omissions = append(result.Omissions, NativePPTXTableOmission{SlideID: slide.ID, ObjectID: "slide", Reason: "Slide root group geometry is outside the table inspection profile."})
				continue
			}
			for _, node := range tree.Children {
				if node.Name != (xml.Name{Space: d.presentation, Local: "graphicFrame"}) {
					continue
				}
				table, objectID, err := inspectNativeTableSource(node, d, budget)
				if err != nil {
					var fatal *nativeInspectionBudgetError
					if errors.As(err, &fatal) {
						return err
					}
					if err := reserve(); err != nil {
						return err
					}
					result.Omissions = append(result.Omissions, NativePPTXTableOmission{SlideID: slide.ID, ObjectID: objectID, Reason: err.Error()})
					continue
				}
				if table == nil {
					continue
				}
				if err := reserve(); err != nil {
					return err
				}
				raw, err := rawNativeNode(payload, node)
				if err != nil {
					return err
				}
				table.SlideID, table.SlideIndex, table.PartName = slide.ID, slideIndex, part
				table.PartSHA256, table.SourceSHA256 = nativeSHA256(payload), nativeSHA256(raw)
				table.SlideSourceSHA256 = slide.Source.FingerprintSHA256
				result.Tables = append(result.Tables, *table)
			}
		}
		return nil
	})
	if err != nil {
		return NativePPTXTableInspection{}, err
	}
	return result, nil
}
