package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"math"
	"regexp"
	"strconv"
	"strings"
	"unicode"
)

const (
	NativeXLSXProtocol        = "injoffice.xlsx.native"
	NativeXLSXVersion         = 1
	NativeXLSXMaxJSONBytes    = 128 * 1024 * 1024
	NativeXLSXMaxSheets       = 1_024
	NativeXLSXMaxCells        = 1_000_000
	NativeXLSXMaxRows         = excelMaxRows
	NativeXLSXMaxColumns      = excelMaxColumns
	NativeXLSXMaxInventory    = 10_000
	NativeXLSXMaxMergedRanges = 100_000
	NativeXLSXMaxTextLength   = 32 * 1024 * 1024
)

// NativeWorkbookSourceV1 identifies the exact source archive and the actual
// spelling of its routed workbook part. Dialect is "transitional" or "strict".
type NativeWorkbookSourceV1 struct {
	PackageSHA256 string `json:"package_sha256"`
	WorkbookPart  string `json:"workbook_part"`
	Dialect       string `json:"dialect"`
	Authority     string `json:"authority"`
}

// NativeWorkbookCapabilityV1 is intentionally wire-compatible with the DOCX
// native contract's capability inventory.
type NativeWorkbookCapabilityV1 struct {
	Name   string  `json:"name"`
	Level  string  `json:"level"`
	Detail *string `json:"detail,omitempty"`
}

// NativeWorkbookPassthroughPartV1 fingerprints the exact bytes that remain
// authoritative for content outside the bounded native projection.
type NativeWorkbookPassthroughPartV1 struct {
	PartName    string `json:"part_name"`
	ContentType string `json:"content_type"`
	ByteLength  *int64 `json:"byte_length"`
	SHA256      string `json:"sha256"`
	Policy      string `json:"policy"`
}

// NativeWorkbookUnsupportedV1 records content that can be displayed or
// preserved but is not safe for the v1 native mutation vocabulary. The full
// item, including its bounded explanatory Message, must remain exact across a
// native save; ID is the stable key used for bidirectional inventory checks.
type NativeWorkbookUnsupportedV1 struct {
	ID           string  `json:"id"`
	Code         string  `json:"code"`
	Capability   string  `json:"capability"`
	ScopeID      string  `json:"scope_id"`
	PartName     *string `json:"part_name,omitempty"`
	CellRef      *string `json:"cell_ref,omitempty"`
	RangeRef     *string `json:"range_ref,omitempty"`
	Preservation string  `json:"preservation"`
	Message      string  `json:"message"`
}

// NativeWorkbookValueV1 keeps the OOXML lexical form. Numeric and date values
// are deliberately strings here so import never rounds or coerces them.
type NativeWorkbookValueV1 struct {
	Kind    string                    `json:"kind"`
	Storage string                    `json:"storage"`
	Lexical *string                   `json:"lexical,omitempty"`
	Text    *string                   `json:"text,omitempty"`
	Rich    bool                      `json:"rich"`
	runs    []NativeWorkbookRichRunV2 `json:"-"`
}

type NativeWorkbookFormulaV1 struct {
	Text        string                 `json:"text"`
	Type        string                 `json:"type"`
	Ref         *string                `json:"ref,omitempty"`
	SharedIndex *uint32                `json:"shared_index,omitempty"`
	Cached      *NativeWorkbookValueV1 `json:"cached,omitempty"`
}

type NativeWorkbookCellV1 struct {
	Row       int                      `json:"row"`
	Column    int                      `json:"column"`
	Ref       string                   `json:"ref"`
	OOXMLType *string                  `json:"ooxml_type,omitempty"`
	StyleID   uint32                   `json:"style_id"`
	Value     *NativeWorkbookValueV1   `json:"value,omitempty"`
	Formula   *NativeWorkbookFormulaV1 `json:"formula,omitempty"`
	Editable  bool                     `json:"editable"`
}

// NativeWorkbookMergedRangeV1 is a renderer-neutral, source-authoritative
// merged grid rectangle. Coordinates are zero-based and inclusive; Ref is the
// same rectangle in canonical bounded A1 notation. Structural mutation remains
// refused until an exact atomic native writer is proven.
type NativeWorkbookMergedRangeV1 struct {
	Ref       string `json:"ref"`
	Row       int    `json:"row"`
	Column    int    `json:"column"`
	EndRow    int    `json:"end_row"`
	EndColumn int    `json:"end_column"`
	Editable  bool   `json:"editable"`
}

type NativeWorkbookRowDimensionV1 struct {
	Row          int      `json:"row"`
	HeightPoints *float64 `json:"height_points,omitempty"`
	Hidden       bool     `json:"hidden"`
	CustomHeight bool     `json:"custom_height"`
	StyleID      *uint32  `json:"style_id,omitempty"`
}

// NativeWorkbookSheetFormatV1 is the exact worksheet default geometry from
// sheetFormatPr. It is absent when the source omits sheetFormatPr; renderers
// must not invent defaults in that case.
type NativeWorkbookSheetFormatV1 struct {
	BaseColumnWidth        *uint32  `json:"base_column_width,omitempty"`
	DefaultColumnWidth     *float64 `json:"default_column_width,omitempty"`
	DefaultRowHeightPoints float64  `json:"default_row_height_points"`
	CustomHeight           bool     `json:"custom_height"`
	ZeroHeight             bool     `json:"zero_height"`
}

// NativeWorkbookNormalStyleV1 identifies the built-in Normal cell-style XF
// and its exact font record. It is omitted unless styles.xml contains one
// unambiguous builtinId=0 declaration with explicit font name and size.
type NativeWorkbookNormalStyleV1 struct {
	StyleXFID        uint32  `json:"style_xf_id"`
	FontID           uint32  `json:"font_id"`
	FontName         string  `json:"font_name"`
	FontSizePoints   float64 `json:"font_size_points"`
	FontBold         bool    `json:"font_bold"`
	FontItalic       bool    `json:"font_italic"`
	FontRecordSHA256 string  `json:"font_record_sha256"`
}

type NativeWorkbookColumnDimensionV1 struct {
	Column      int      `json:"column"`
	EndColumn   int      `json:"end_column"`
	Width       *float64 `json:"width,omitempty"`
	Hidden      bool     `json:"hidden"`
	CustomWidth bool     `json:"custom_width"`
	BestFit     bool     `json:"best_fit"`
	StyleID     *uint32  `json:"style_id,omitempty"`
}

// NativeWorkbookBorderSideV1 retains the exact SpreadsheetML line-style token
// and direct opaque RGB color. Physical stroke width/dash phase is deliberately
// not inferred from the qualitative OOXML token.
type NativeWorkbookBorderSideV1 struct {
	Style string `json:"style"`
	Color string `json:"color"`
}

// NativeWorkbookBorderV1 is either the implicit no-border default used when a
// package has no styles part, or one exact supported borders-table record.
// Unsupported/theme/indexed/diagonal records are omitted and classified on the
// containing effective style instead of being flattened approximately.
type NativeWorkbookBorderV1 struct {
	Origin       string                      `json:"origin"`
	BorderID     *uint32                     `json:"border_id,omitempty"`
	RecordSHA256 *string                     `json:"record_sha256,omitempty"`
	Left         *NativeWorkbookBorderSideV1 `json:"left,omitempty"`
	Right        *NativeWorkbookBorderSideV1 `json:"right,omitempty"`
	Top          *NativeWorkbookBorderSideV1 `json:"top,omitempty"`
	Bottom       *NativeWorkbookBorderSideV1 `json:"bottom,omitempty"`
}

// NativeWorkbookFillV1 binds the useful direct-RGB fill to the exact raw
// fills-table record that supplied it. An implicit default exists only when the
// package has no styles part.
type NativeWorkbookFillV1 struct {
	Origin       string  `json:"origin"`
	FillID       *uint32 `json:"fill_id,omitempty"`
	RecordSHA256 *string `json:"record_sha256,omitempty"`
	Color        *string `json:"color,omitempty"`
}

// NativeWorkbookEffectiveStyleV1 is the useful subset shared with
// @injoffice/sheets style.patch. Omitted color/name/size fields either inherit
// no explicit value or are accompanied by an Unsupported code.
type NativeWorkbookEffectiveStyleV1 struct {
	NumberFormat        *string                 `json:"number_format,omitempty"`
	FontName            *string                 `json:"font_name,omitempty"`
	FontSizePoints      *float64                `json:"font_size_points,omitempty"`
	Bold                *bool                   `json:"bold,omitempty"`
	Italic              *bool                   `json:"italic,omitempty"`
	FontColor           *string                 `json:"font_color,omitempty"`
	FillColor           *string                 `json:"fill_color,omitempty"`
	Fill                *NativeWorkbookFillV1   `json:"fill,omitempty"`
	Border              *NativeWorkbookBorderV1 `json:"border,omitempty"`
	HorizontalAlignment *string                 `json:"horizontal_alignment,omitempty"`
	VerticalAlignment   *string                 `json:"vertical_alignment,omitempty"`
	WrapText            *bool                   `json:"wrap_text,omitempty"`
	Projection          string                  `json:"projection"`
	Unsupported         []string                `json:"unsupported"`
}

type NativeWorkbookStyleV1 struct {
	ID                  uint32                         `json:"id"`
	Effective           NativeWorkbookEffectiveStyleV1 `json:"effective"`
	RawProjectionSHA256 string                         `json:"raw_projection_sha256"`
}

type NativeWorkbookSheetV1 struct {
	ID           string                            `json:"id"`
	Name         string                            `json:"name"`
	Order        int                               `json:"order"`
	State        string                            `json:"state"`
	PartName     string                            `json:"part_name"`
	SheetFormat  *NativeWorkbookSheetFormatV1      `json:"sheet_format,omitempty"`
	Rows         []NativeWorkbookRowDimensionV1    `json:"rows"`
	Columns      []NativeWorkbookColumnDimensionV1 `json:"columns"`
	Cells        []NativeWorkbookCellV1            `json:"cells"`
	MergedRanges []NativeWorkbookMergedRangeV1     `json:"merged_ranges"`
	Editable     bool                              `json:"editable"`
	RefusalCode  *string                           `json:"refusal_code,omitempty"`
}

type NativeWorkbookV1 struct {
	Protocol         string                            `json:"protocol"`
	Version          int                               `json:"version"`
	DocumentID       string                            `json:"document_id"`
	Revision         string                            `json:"revision"`
	Source           NativeWorkbookSourceV1            `json:"source"`
	NormalStyle      *NativeWorkbookNormalStyleV1      `json:"normal_style,omitempty"`
	Sheets           []NativeWorkbookSheetV1           `json:"sheets"`
	Styles           []NativeWorkbookStyleV1           `json:"styles"`
	Capabilities     []NativeWorkbookCapabilityV1      `json:"capabilities"`
	PassthroughParts []NativeWorkbookPassthroughPartV1 `json:"passthrough_parts"`
	Unsupported      []NativeWorkbookUnsupportedV1     `json:"unsupported"`
}

type NativeWorkbookValidationIssue struct {
	Code    string `json:"code"`
	Path    string `json:"path"`
	Message string `json:"message"`
}

type NativeWorkbookValidationError struct {
	Issues []NativeWorkbookValidationIssue `json:"issues"`
}

func (err *NativeWorkbookValidationError) Error() string {
	if len(err.Issues) == 0 {
		return "invalid native XLSX workbook"
	}
	return fmt.Sprintf("invalid native XLSX workbook at %s: %s", err.Issues[0].Path, err.Issues[0].Message)
}

var (
	nativeWorkbookIDPattern = regexp.MustCompile(`^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$`)
	nativeWorkbookSHA       = regexp.MustCompile(`^sha256:[0-9a-f]{64}$`)
	nativeWorkbookRevision  = regexp.MustCompile(`^rev:[0-9a-f]{64}$`)
)

func nativeWorkbookIssue(issues *[]NativeWorkbookValidationIssue, code, path, message string) {
	if len(*issues) >= 100 {
		return
	}
	*issues = append(*issues, NativeWorkbookValidationIssue{Code: code, Path: path, Message: message})
}

// ValidateNativeWorkbookV1 validates a generated contract or Previous input.
// It is intentionally strict about identities, coordinates and collection
// bounds so Previous cannot redirect or amplify a subsequent extraction.
func ValidateNativeWorkbookV1(workbook *NativeWorkbookV1) []NativeWorkbookValidationIssue {
	var issues []NativeWorkbookValidationIssue
	if workbook == nil {
		nativeWorkbookIssue(&issues, "REQUIRED", "", "workbook is required")
		return issues
	}
	if workbook.Protocol != NativeXLSXProtocol {
		nativeWorkbookIssue(&issues, "UNSUPPORTED_PROTOCOL", "/protocol", "unsupported protocol")
	}
	if workbook.Version != NativeXLSXVersion {
		nativeWorkbookIssue(&issues, "UNSUPPORTED_VERSION", "/version", "unsupported version")
	}
	if !nativeWorkbookIDPattern.MatchString(workbook.DocumentID) {
		nativeWorkbookIssue(&issues, "INVALID_ID", "/document_id", "invalid document identity")
	}
	if !nativeWorkbookRevision.MatchString(workbook.Revision) {
		nativeWorkbookIssue(&issues, "INVALID_REVISION", "/revision", "revision must contain a full SHA-256")
	}
	if !nativeWorkbookSHA.MatchString(workbook.Source.PackageSHA256) {
		nativeWorkbookIssue(&issues, "INVALID_SHA256", "/source/package_sha256", "invalid package SHA-256")
	}
	if nativeWorkbookRevision.MatchString(workbook.Revision) && nativeWorkbookSHA.MatchString(workbook.Source.PackageSHA256) && strings.TrimPrefix(workbook.Revision, "rev:") != strings.TrimPrefix(workbook.Source.PackageSHA256, "sha256:") {
		nativeWorkbookIssue(&issues, "INVALID_REVISION", "/revision", "revision must fingerprint the same exact package bytes as source.package_sha256")
	}
	if workbook.Source.WorkbookPart == "" {
		nativeWorkbookIssue(&issues, "REQUIRED", "/source/workbook_part", "workbook part is required")
	} else if _, err := canonicalOPCPartKey(workbook.Source.WorkbookPart); err != nil {
		nativeWorkbookIssue(&issues, "INVALID_PART", "/source/workbook_part", err.Error())
	}
	if workbook.Source.Dialect != "transitional" && workbook.Source.Dialect != "strict" {
		nativeWorkbookIssue(&issues, "INVALID_VALUE", "/source/dialect", "dialect must be transitional or strict")
	}
	if workbook.NormalStyle != nil {
		normal := workbook.NormalStyle
		if normal.FontName == "" || strings.TrimSpace(normal.FontName) != normal.FontName || utf16Length(normal.FontName) > maxFontNameLength {
			nativeWorkbookIssue(&issues, "INVALID_VALUE", "/normal_style/font_name", "Normal-style font name must be a bounded non-empty value")
		}
		if math.IsNaN(normal.FontSizePoints) || math.IsInf(normal.FontSizePoints, 0) || normal.FontSizePoints <= 0 || normal.FontSizePoints > maxRowHeightPoints {
			nativeWorkbookIssue(&issues, "INVALID_NUMBER", "/normal_style/font_size_points", "Normal-style font size must be finite and within 0..409.5 points")
		}
		if !nativeWorkbookSHA.MatchString(normal.FontRecordSHA256) {
			nativeWorkbookIssue(&issues, "INVALID_SHA256", "/normal_style/font_record_sha256", "Normal-style font record requires a full SHA-256")
		}
	}
	if len(workbook.Sheets) > NativeXLSXMaxSheets {
		nativeWorkbookIssue(&issues, "LIMIT_EXCEEDED", "/sheets", fmt.Sprintf("exceeds %d sheets", NativeXLSXMaxSheets))
	}
	seenSheets, seenParts, seenNames := map[string]bool{}, map[string]bool{}, map[string]bool{}
	cellCount, mergedRangeCount := 0, 0
	for index, sheet := range workbook.Sheets {
		path := "/sheets/" + strconv.Itoa(index)
		if sheet.Order != index {
			nativeWorkbookIssue(&issues, "INVALID_ORDER", path+"/order", "sheet order must match array order")
		}
		if !validNativeSheetID(sheet.ID) {
			nativeWorkbookIssue(&issues, "INVALID_ID", path+"/id", "sheet id must be an OOXML unsigned integer in 1..4294967295")
		} else if seenSheets[sheet.ID] {
			nativeWorkbookIssue(&issues, "DUPLICATE_ID", path+"/id", "duplicate sheet id")
		}
		seenSheets[sheet.ID] = true
		if sheet.Name == "" {
			nativeWorkbookIssue(&issues, "REQUIRED", path+"/name", "sheet name is required")
		}
		nameKey := nativeSheetNameCaseKey(sheet.Name)
		if strings.HasPrefix(sheet.Name, "'") || strings.HasSuffix(sheet.Name, "'") {
			nativeWorkbookIssue(&issues, "INVALID_VALUE", path+"/name", "sheet names cannot begin or end with an apostrophe")
		}
		if seenNames[nameKey] {
			nativeWorkbookIssue(&issues, "DUPLICATE_ID", path+"/name", "sheet names must be case-insensitively unique")
		}
		seenNames[nameKey] = true
		if sheet.State != "visible" && sheet.State != "hidden" && sheet.State != "veryHidden" {
			nativeWorkbookIssue(&issues, "INVALID_VALUE", path+"/state", "invalid sheet state")
		}
		partKey, err := canonicalOPCPartKey(sheet.PartName)
		if err != nil {
			nativeWorkbookIssue(&issues, "INVALID_PART", path+"/part_name", err.Error())
		} else if seenParts[partKey] {
			nativeWorkbookIssue(&issues, "DUPLICATE_PART", path+"/part_name", "multiple sheets route to the same worksheet part")
		}
		seenParts[partKey] = true
		if sheet.SheetFormat != nil {
			format := sheet.SheetFormat
			if math.IsNaN(format.DefaultRowHeightPoints) || math.IsInf(format.DefaultRowHeightPoints, 0) || format.DefaultRowHeightPoints <= 0 || format.DefaultRowHeightPoints > maxRowHeightPoints {
				nativeWorkbookIssue(&issues, "INVALID_NUMBER", path+"/sheet_format/default_row_height_points", "default row height must be finite and within 0..409.5 points")
			}
			if format.BaseColumnWidth != nil && *format.BaseColumnWidth > uint32(maxColumnWidth) {
				nativeWorkbookIssue(&issues, "INVALID_NUMBER", path+"/sheet_format/base_column_width", "base column width is outside Excel bounds")
			}
			if format.DefaultColumnWidth != nil && (math.IsNaN(*format.DefaultColumnWidth) || math.IsInf(*format.DefaultColumnWidth, 0) || *format.DefaultColumnWidth <= 0 || *format.DefaultColumnWidth > maxColumnWidth) {
				nativeWorkbookIssue(&issues, "INVALID_NUMBER", path+"/sheet_format/default_column_width", "default column width must be finite and within 0..255")
			}
		}
		lastRowDimension := -1
		for rowIndex, row := range sheet.Rows {
			rowPath := path + "/rows/" + strconv.Itoa(rowIndex)
			if row.Row < 0 || row.Row >= excelMaxRows || row.Row <= lastRowDimension {
				nativeWorkbookIssue(&issues, "INVALID_ORDER", rowPath+"/row", "row dimensions must be unique, ordered, and within Excel bounds")
			}
			lastRowDimension = row.Row
			if row.HeightPoints != nil && (math.IsNaN(*row.HeightPoints) || math.IsInf(*row.HeightPoints, 0) || *row.HeightPoints < 0 || *row.HeightPoints > maxRowHeightPoints) {
				nativeWorkbookIssue(&issues, "INVALID_NUMBER", rowPath+"/height_points", "row height is outside native mutation bounds")
			}
			if row.StyleID != nil && int(*row.StyleID) >= len(workbook.Styles) {
				nativeWorkbookIssue(&issues, "INVALID_REFERENCE", rowPath+"/style_id", "row style id is outside styles")
			}
		}
		lastColumn := -1
		for columnIndex, column := range sheet.Columns {
			columnPath := path + "/columns/" + strconv.Itoa(columnIndex)
			if column.Column < 0 || column.EndColumn < column.Column || column.EndColumn >= excelMaxColumns || column.Column <= lastColumn {
				nativeWorkbookIssue(&issues, "INVALID_ORDER", columnPath, "column ranges must be non-overlapping, ordered, and within Excel bounds")
			}
			lastColumn = column.EndColumn
			if column.Width != nil && (math.IsNaN(*column.Width) || math.IsInf(*column.Width, 0) || *column.Width < 0 || *column.Width > maxColumnWidth) {
				nativeWorkbookIssue(&issues, "INVALID_NUMBER", columnPath+"/width", "column width is outside native mutation bounds")
			}
			if column.StyleID != nil && int(*column.StyleID) >= len(workbook.Styles) {
				nativeWorkbookIssue(&issues, "INVALID_REFERENCE", columnPath+"/style_id", "column style id is outside styles")
			}
		}
		lastCell := -1
		for cellIndex, cell := range sheet.Cells {
			cellPath := path + "/cells/" + strconv.Itoa(cellIndex)
			if cell.Row < 0 || cell.Row >= excelMaxRows || cell.Column < 0 || cell.Column >= excelMaxColumns {
				nativeWorkbookIssue(&issues, "OUT_OF_RANGE", cellPath, "cell is outside Excel bounds")
			}
			position := cell.Row*excelMaxColumns + cell.Column
			if position <= lastCell {
				nativeWorkbookIssue(&issues, "INVALID_ORDER", cellPath, "cells must be unique and row-major ordered")
			}
			lastCell = position
			if expected := cellReference(cell.Row, cell.Column); cell.Ref != expected {
				nativeWorkbookIssue(&issues, "INVALID_REFERENCE", cellPath+"/ref", "cell reference does not match coordinates")
			}
			if int(cell.StyleID) >= len(workbook.Styles) {
				nativeWorkbookIssue(&issues, "INVALID_REFERENCE", cellPath+"/style_id", "cell style id is outside styles")
			}
			if cell.Value != nil && cell.Formula != nil {
				nativeWorkbookIssue(&issues, "INVALID_UNION", cellPath, "cell cannot contain both a literal value and a formula")
			}
			if cell.OOXMLType != nil && *cell.OOXMLType != "n" && *cell.OOXMLType != "s" && *cell.OOXMLType != "str" && *cell.OOXMLType != "inlineStr" && *cell.OOXMLType != "b" && *cell.OOXMLType != "e" && *cell.OOXMLType != "d" {
				nativeWorkbookIssue(&issues, "INVALID_VALUE", cellPath+"/ooxml_type", "invalid OOXML cell type")
			}
			if cell.Value != nil {
				if !validNativeValueKind(cell.Value.Kind) || !validNativeValueStorage(cell.Value.Storage) {
					nativeWorkbookIssue(&issues, "INVALID_VALUE", cellPath+"/value", "invalid native cell value kind/storage")
				}
			}
			if cell.Formula != nil && cell.Formula.Type != "normal" && cell.Formula.Type != "shared" && cell.Formula.Type != "array" && cell.Formula.Type != "dataTable" {
				nativeWorkbookIssue(&issues, "INVALID_VALUE", cellPath+"/formula/type", "invalid formula type")
			}
			cellCount++
		}
		lastMergedPosition := -1
		validMergedRanges := make([]NativeWorkbookMergedRangeV1, 0, len(sheet.MergedRanges))
		for mergedIndex, merged := range sheet.MergedRanges {
			mergedPath := path + "/merged_ranges/" + strconv.Itoa(mergedIndex)
			if merged.Row < 0 || merged.Row >= excelMaxRows || merged.Column < 0 || merged.Column >= excelMaxColumns || merged.EndRow < merged.Row || merged.EndRow >= excelMaxRows || merged.EndColumn < merged.Column || merged.EndColumn >= excelMaxColumns {
				nativeWorkbookIssue(&issues, "OUT_OF_RANGE", mergedPath, "merged range coordinates must be forward and within Excel bounds")
			} else {
				validMergedRanges = append(validMergedRanges, merged)
				position := merged.Row*excelMaxColumns + merged.Column
				if position <= lastMergedPosition {
					nativeWorkbookIssue(&issues, "INVALID_ORDER", mergedPath, "merged ranges must be unique and row-major ordered by their top-left cell")
				}
				lastMergedPosition = position
				if merged.Row == merged.EndRow && merged.Column == merged.EndColumn {
					nativeWorkbookIssue(&issues, "INVALID_REFERENCE", mergedPath, "merged range must span at least two cells")
				}
				expected := cellReference(merged.Row, merged.Column) + ":" + cellReference(merged.EndRow, merged.EndColumn)
				if merged.Ref != expected {
					nativeWorkbookIssue(&issues, "INVALID_REFERENCE", mergedPath+"/ref", "merged range ref must be canonical bounded A1 and exactly match its coordinates")
				}
			}
			if merged.Editable {
				nativeWorkbookIssue(&issues, "INVALID_VALUE", mergedPath+"/editable", "merged ranges are structurally mutation-refused in native v1")
			}
			mergedRangeCount++
		}
		if err := rejectOverlappingNativeMergedRanges(validMergedRanges); err != nil {
			nativeWorkbookIssue(&issues, "INVALID_UNION", path+"/merged_ranges", err.Error())
		}
	}
	if cellCount > NativeXLSXMaxCells {
		nativeWorkbookIssue(&issues, "LIMIT_EXCEEDED", "/sheets", fmt.Sprintf("exceeds %d cells", NativeXLSXMaxCells))
	}
	if mergedRangeCount > NativeXLSXMaxMergedRanges {
		nativeWorkbookIssue(&issues, "LIMIT_EXCEEDED", "/sheets", fmt.Sprintf("exceeds %d merged ranges", NativeXLSXMaxMergedRanges))
	}
	if len(workbook.Styles) == 0 || len(workbook.Styles) > maxStyleTableRecords {
		nativeWorkbookIssue(&issues, "LIMIT_EXCEEDED", "/styles", fmt.Sprintf("styles must contain 1..%d entries", maxStyleTableRecords))
	}
	for index, style := range workbook.Styles {
		if uint64(style.ID) != uint64(index) {
			nativeWorkbookIssue(&issues, "INVALID_ORDER", "/styles/"+strconv.Itoa(index)+"/id", "style ids must be dense and ordered")
		}
		for name, value := range map[string]*float64{"font_size_points": style.Effective.FontSizePoints} {
			if value != nil && (math.IsNaN(*value) || math.IsInf(*value, 0)) {
				nativeWorkbookIssue(&issues, "INVALID_NUMBER", "/styles/"+strconv.Itoa(index)+"/effective/"+name, "must be finite")
			}
		}
	}
	if len(workbook.PassthroughParts) > NativeXLSXMaxInventory || len(workbook.Unsupported) > NativeXLSXMaxInventory {
		nativeWorkbookIssue(&issues, "LIMIT_EXCEEDED", "/passthrough_parts", fmt.Sprintf("inventory exceeds %d entries", NativeXLSXMaxInventory))
	}
	seenInventory := map[string]bool{}
	for index, part := range workbook.PassthroughParts {
		path := "/passthrough_parts/" + strconv.Itoa(index)
		key, err := canonicalOPCPartKey(part.PartName)
		if err != nil {
			nativeWorkbookIssue(&issues, "INVALID_PART", path+"/part_name", err.Error())
		} else if seenInventory[key] {
			nativeWorkbookIssue(&issues, "DUPLICATE_PART", path+"/part_name", "duplicate passthrough part")
		}
		seenInventory[key] = true
		if !nativeWorkbookSHA.MatchString(part.SHA256) {
			nativeWorkbookIssue(&issues, "INVALID_SHA256", path+"/sha256", "invalid part SHA-256")
		}
		if part.ByteLength == nil || *part.ByteLength < 0 || part.ContentType == "" || part.Policy != "preserve-exact" {
			nativeWorkbookIssue(&issues, "INVALID_VALUE", path, "invalid passthrough metadata")
		}
	}
	validateNativeWorkbookDeep(workbook, &issues)
	return issues
}

func validNativeSheetID(value string) bool {
	canonical, err := canonicalNativeSheetID(value)
	return err == nil && canonical == value
}

// nativeSheetNameCaseKey is a locale-independent, per-scalar Unicode casing
// closure shared with the TypeScript binding. Upper-then-lower joins special
// lowercase forms such as Greek final sigma; U+0130 is joined explicitly so
// dotted-I names are refused under the Office "in any locale" uniqueness rule.
func nativeSheetNameCaseKey(value string) string {
	var key strings.Builder
	for _, character := range value {
		if character == '\u0130' {
			key.WriteRune('i')
			continue
		}
		key.WriteRune(unicode.ToLower(unicode.ToUpper(character)))
	}
	return key.String()
}

func canonicalNativeSheetID(value string) (string, error) {
	if value == "" {
		return "", fmt.Errorf("empty sheet id")
	}
	for _, character := range value {
		if character < '0' || character > '9' {
			return "", fmt.Errorf("sheet id %q is not an unsigned integer", value)
		}
	}
	parsed, err := strconv.ParseUint(value, 10, 32)
	if err != nil || parsed == 0 {
		return "", fmt.Errorf("sheet id %q is outside 1..4294967295", value)
	}
	return strconv.FormatUint(parsed, 10), nil
}

func validNativeValueKind(value string) bool {
	return value == "number" || value == "boolean" || value == "error" || value == "date" || value == "string"
}

func validNativeValueStorage(value string) bool {
	return value == "number" || value == "boolean" || value == "error" || value == "date" || value == "formula-string" || value == "shared" || value == "inline"
}

// EncodeNativeWorkbookV1 validates before emitting a deterministic JSON value.
func EncodeNativeWorkbookV1(workbook *NativeWorkbookV1) ([]byte, error) {
	if issues := ValidateNativeWorkbookV1(workbook); len(issues) != 0 {
		return nil, &NativeWorkbookValidationError{Issues: issues}
	}
	data, err := json.Marshal(workbook)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: encode native XLSX v1: %w", err)
	}
	if len(data) > NativeXLSXMaxJSONBytes {
		return nil, fmt.Errorf("xlsxpatch: encode native XLSX v1: JSON exceeds %d bytes", NativeXLSXMaxJSONBytes)
	}
	return data, nil
}

// DecodeNativeWorkbookV1 rejects unknown fields and validates Previous input.
func DecodeNativeWorkbookV1(data []byte) (*NativeWorkbookV1, error) {
	if len(data) == 0 || len(data) > NativeXLSXMaxJSONBytes {
		return nil, fmt.Errorf("xlsxpatch: decode native XLSX v1: JSON size must be 1..%d bytes", NativeXLSXMaxJSONBytes)
	}
	if err := rejectDuplicateNativeWorkbookJSONKeys(data); err != nil {
		return nil, fmt.Errorf("xlsxpatch: decode native XLSX v1: %w", err)
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	var workbook NativeWorkbookV1
	if err := decoder.Decode(&workbook); err != nil {
		return nil, fmt.Errorf("xlsxpatch: decode native XLSX v1: %w", err)
	}
	if err := requireNativeWorkbookJSONEOF(decoder); err != nil {
		return nil, err
	}
	if issues := ValidateNativeWorkbookV1(&workbook); len(issues) != 0 {
		return nil, &NativeWorkbookValidationError{Issues: issues}
	}
	return &workbook, nil
}

func requireNativeWorkbookJSONEOF(decoder *json.Decoder) error {
	var trailing any
	if err := decoder.Decode(&trailing); err != io.EOF {
		if err == nil {
			return fmt.Errorf("xlsxpatch: decode native XLSX v1: trailing JSON value")
		}
		return fmt.Errorf("xlsxpatch: decode native XLSX v1: %w", err)
	}
	return nil
}

func nativeWorkbookString(value string) *string  { return &value }
func nativeWorkbookBool(value bool) *bool        { return &value }
func nativeWorkbookFloat(value float64) *float64 { return &value }
func nativeWorkbookInt64(value int64) *int64     { return &value }

func validNativeWorkbookIdentifier(value string) bool {
	return nativeWorkbookIDPattern.MatchString(value) && strings.TrimSpace(value) == value
}
