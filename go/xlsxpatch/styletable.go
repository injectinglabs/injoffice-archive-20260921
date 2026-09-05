package xlsxpatch

import (
	"bytes"
	"encoding/xml"
	"fmt"
	"io"
	"strconv"
)

const (
	relTypeStylesTransitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
	relTypeStylesStrict       = "http://purl.oclc.org/ooxml/officeDocument/relationships/styles"
	maxStyleTableRecords      = 100_000
)

type styleTableEntry struct {
	span     xmlSpan
	start    xml.StartElement
	children []styleTableNode
}

type styleTableNode struct {
	span     xmlSpan
	start    xml.StartElement
	children []styleTableNode
}

type styleTableContainer struct {
	xmlSpan
	local   string
	entries []styleTableEntry
	count   int
}

type styleTableIndex struct {
	root         xmlSpan
	namespace    string
	numFmts      *styleTableContainer
	fonts        *styleTableContainer
	fills        *styleTableContainer
	borders      *styleTableContainer
	cellStyleXfs *styleTableContainer
	cellXfs      *styleTableContainer
}

type activeStyleEntry struct {
	entry styleTableEntry
	depth int
	stack []activeStyleNode
}

type activeStyleNode struct {
	node  styleTableNode
	depth int
}

func parseStyleTable(data []byte) (styleTableIndex, error) {
	index := styleTableIndex{}
	decoder := xml.NewDecoder(bytes.NewReader(data))
	depth := 0
	rootClosed := false
	var activeContainer *styleTableContainer
	containerDepth := 0
	var activeEntry *activeStyleEntry
	declaredRecords := 0
	parsedRecords := 0
	for {
		before := int(decoder.InputOffset())
		token, err := decoder.Token()
		after := int(decoder.InputOffset())
		if err == io.EOF {
			break
		}
		if err != nil {
			return styleTableIndex{}, fmt.Errorf("parse styles XML: %w", err)
		}
		switch token := token.(type) {
		case xml.StartElement:
			depth++
			if depth == 1 {
				if index.root.startTagEnd != 0 {
					return styleTableIndex{}, fmt.Errorf("styles XML has multiple root elements")
				}
				if token.Name.Local != "styleSheet" || (token.Name.Space != spreadsheetMLTransitional && token.Name.Space != spreadsheetMLStrict) {
					return styleTableIndex{}, fmt.Errorf("root element is not a supported SpreadsheetML styleSheet")
				}
				index.namespace = token.Name.Space
				index.root = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				continue
			}
			if depth == 2 && token.Name.Space == index.namespace {
				container, expectedChild := styleContainerFor(&index, token.Name.Local)
				if container == nil {
					continue
				}
				if container.startTagEnd != 0 {
					return styleTableIndex{}, fmt.Errorf("duplicate %s style table", token.Name.Local)
				}
				count, err := styleTableCount(token)
				if err != nil {
					return styleTableIndex{}, fmt.Errorf("%s: %w", token.Name.Local, err)
				}
				if count > maxStyleTableRecords-declaredRecords {
					return styleTableIndex{}, fmt.Errorf("style table counts exceed the %d-record cumulative safety limit", maxStyleTableRecords)
				}
				declaredRecords += count
				container.xmlSpan = xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}
				container.local = expectedChild
				container.count = count
				activeContainer = container
				containerDepth = depth
				continue
			}
			if activeContainer != nil && depth == containerDepth+1 {
				if token.Name.Space != index.namespace || token.Name.Local != activeContainer.local {
					return styleTableIndex{}, fmt.Errorf("%s has unsupported direct child %q", styleContainerName(index, activeContainer), rawQName(data, before, after))
				}
				activeEntry = &activeStyleEntry{
					entry: styleTableEntry{span: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, start: token},
					depth: depth,
				}
				continue
			}
			if activeEntry != nil && depth > activeEntry.depth {
				activeEntry.stack = append(activeEntry.stack, activeStyleNode{
					node:  styleTableNode{span: xmlSpan{qname: rawQName(data, before, after), start: before, startTagEnd: after}, start: token},
					depth: depth,
				})
			}
		case xml.EndElement:
			if activeEntry != nil && len(activeEntry.stack) > 0 {
				last := len(activeEntry.stack) - 1
				active := activeEntry.stack[last]
				if depth == active.depth {
					active.node.span.endStart, active.node.span.end = before, after
					activeEntry.stack = activeEntry.stack[:last]
					if len(activeEntry.stack) == 0 {
						activeEntry.entry.children = append(activeEntry.entry.children, active.node)
					} else {
						parent := &activeEntry.stack[len(activeEntry.stack)-1]
						parent.node.children = append(parent.node.children, active.node)
					}
				}
			}
			if activeEntry != nil && depth == activeEntry.depth && token.Name.Space == index.namespace && token.Name.Local == activeContainer.local {
				if len(activeEntry.stack) != 0 {
					return styleTableIndex{}, fmt.Errorf("%s record has incomplete nested markup", activeContainer.local)
				}
				activeEntry.entry.span.endStart, activeEntry.entry.span.end = before, after
				if err := claimStyleTableRecord(&parsedRecords); err != nil {
					return styleTableIndex{}, err
				}
				activeContainer.entries = append(activeContainer.entries, activeEntry.entry)
				activeEntry = nil
			}
			if activeContainer != nil && depth == containerDepth && token.Name.Space == index.namespace && token.Name.Local == styleContainerName(index, activeContainer) {
				activeContainer.endStart, activeContainer.end = before, after
				if len(activeContainer.entries) != activeContainer.count {
					return styleTableIndex{}, fmt.Errorf("%s count=%d does not match %d direct records", token.Name.Local, activeContainer.count, len(activeContainer.entries))
				}
				activeContainer = nil
				containerDepth = 0
			}
			if depth == 1 && token.Name.Space == index.namespace && token.Name.Local == "styleSheet" {
				index.root.endStart, index.root.end = before, after
				rootClosed = true
			}
			depth--
		case xml.CharData:
			if activeContainer != nil && activeEntry == nil && depth == containerDepth && len(bytes.TrimSpace(token)) != 0 {
				return styleTableIndex{}, fmt.Errorf("%s has unsupported text content", styleContainerName(index, activeContainer))
			}
		case xml.ProcInst:
			if depth == 0 && index.root.startTagEnd == 0 && token.Target == "xml" {
				continue
			}
			return styleTableIndex{}, fmt.Errorf("styles XML has unsupported processing instruction %q", token.Target)
		case xml.Directive:
			return styleTableIndex{}, fmt.Errorf("styles XML has unsupported directive")
		case xml.Comment:
			return styleTableIndex{}, fmt.Errorf("styles XML has unsupported comment")
		}
	}
	if depth != 0 || !rootClosed || index.root.end == 0 {
		return styleTableIndex{}, fmt.Errorf("missing complete styleSheet root")
	}
	for name, container := range map[string]*styleTableContainer{
		"fonts": index.fonts, "fills": index.fills, "borders": index.borders,
		"cellStyleXfs": index.cellStyleXfs, "cellXfs": index.cellXfs,
	} {
		if container == nil || container.end == 0 || container.count == 0 {
			return styleTableIndex{}, fmt.Errorf("styleSheet requires a non-empty %s table", name)
		}
	}
	if index.numFmts != nil && index.numFmts.end == 0 {
		return styleTableIndex{}, fmt.Errorf("missing complete numFmts table")
	}
	totalRecords := len(index.fonts.entries) + len(index.fills.entries) + len(index.borders.entries) + len(index.cellStyleXfs.entries) + len(index.cellXfs.entries)
	if index.numFmts != nil {
		totalRecords += len(index.numFmts.entries)
	}
	if totalRecords > maxStyleTableRecords {
		return styleTableIndex{}, fmt.Errorf("style tables contain %d records, exceeding the %d-record safety limit", totalRecords, maxStyleTableRecords)
	}
	return index, nil
}

func claimStyleTableRecord(total *int) error {
	if *total >= maxStyleTableRecords {
		return fmt.Errorf("style tables exceed the %d-record cumulative safety limit", maxStyleTableRecords)
	}
	*total++
	return nil
}

func styleContainerFor(index *styleTableIndex, local string) (*styleTableContainer, string) {
	switch local {
	case "numFmts":
		if index.numFmts == nil {
			index.numFmts = &styleTableContainer{}
		}
		return index.numFmts, "numFmt"
	case "fonts":
		if index.fonts == nil {
			index.fonts = &styleTableContainer{}
		}
		return index.fonts, "font"
	case "fills":
		if index.fills == nil {
			index.fills = &styleTableContainer{}
		}
		return index.fills, "fill"
	case "borders":
		if index.borders == nil {
			index.borders = &styleTableContainer{}
		}
		return index.borders, "border"
	case "cellStyleXfs":
		if index.cellStyleXfs == nil {
			index.cellStyleXfs = &styleTableContainer{}
		}
		return index.cellStyleXfs, "xf"
	case "cellXfs":
		if index.cellXfs == nil {
			index.cellXfs = &styleTableContainer{}
		}
		return index.cellXfs, "xf"
	default:
		return nil, ""
	}
}

func styleContainerName(index styleTableIndex, container *styleTableContainer) string {
	switch container {
	case index.numFmts:
		return "numFmts"
	case index.fonts:
		return "fonts"
	case index.fills:
		return "fills"
	case index.borders:
		return "borders"
	case index.cellStyleXfs:
		return "cellStyleXfs"
	case index.cellXfs:
		return "cellXfs"
	default:
		return "style table"
	}
}

func styleTableCount(start xml.StartElement) (int, error) {
	raw, found, err := unqualifiedXMLAttribute(start, "count")
	if err != nil {
		return 0, err
	}
	if !found || raw == "" {
		return 0, fmt.Errorf("requires a count attribute")
	}
	for _, character := range raw {
		if character < '0' || character > '9' {
			return 0, fmt.Errorf("count=%q is not an unsigned integer", raw)
		}
	}
	count, err := strconv.Atoi(raw)
	if err != nil {
		return 0, fmt.Errorf("count=%q is outside platform limits", raw)
	}
	if count > maxStyleTableRecords {
		return 0, fmt.Errorf("count=%d exceeds the %d-record safety limit", count, maxStyleTableRecords)
	}
	return count, nil
}
