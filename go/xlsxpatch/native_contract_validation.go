package xlsxpatch

import (
	"bytes"
	"crypto/sha256"
	"encoding/hex"
	"encoding/json"
	"fmt"
	"io"
	"math"
	"regexp"
	"sort"
	"strconv"
	"strings"
	"time"
	"unicode/utf8"
)

const (
	nativeWorkbookMaxMetadataLength = 8 * 1024
	nativeWorkbookMaxJSONDepth      = 256
	nativeWorkbookMaxJSONTokens     = 4_000_000
)

var (
	nativeWorkbookColorPattern      = regexp.MustCompile(`^#[0-9A-F]{6}$`)
	nativeWorkbookUnsupportedID     = regexp.MustCompile(`^unsupported:[0-9a-f]{64}$`)
	nativeWorkbookUnsupportedCode   = regexp.MustCompile(`^[A-Z][A-Z0-9_]{0,127}$`)
	nativeWorkbookCapabilityName    = regexp.MustCompile(`^[a-z][a-z0-9-]{0,127}$`)
	nativeWorkbookISODate           = regexp.MustCompile(`^[0-9]{4}-[0-9]{2}-[0-9]{2}(?:Z|[+-][0-9]{2}:[0-9]{2})?$`)
	nativeWorkbookISOLocalDateTime  = regexp.MustCompile(`^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?$`)
	nativeWorkbookISOOffsetDateTime = regexp.MustCompile(`^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}(?:\.[0-9]+)?(?:Z|[+-][0-9]{2}:[0-9]{2})$`)
)

type nativeWorkbookValidationState struct {
	issues    *[]NativeWorkbookValidationIssue
	textBytes int
}

func (state *nativeWorkbookValidationState) issue(code, path, message string) {
	nativeWorkbookIssue(state.issues, code, path, message)
}

func (state *nativeWorkbookValidationState) text(path, value string, maximum int, required bool) {
	if !utf8.ValidString(value) {
		state.issue("INVALID_TEXT", path, "must be valid UTF-8")
		return
	}
	if required && value == "" {
		state.issue("REQUIRED", path, "must not be empty")
	}
	if utf8.RuneCountInString(value) > maximum {
		state.issue("LIMIT_EXCEEDED", path, fmt.Sprintf("exceeds %d Unicode scalar values", maximum))
	}
	state.textBytes += len(value)
}

func (state *nativeWorkbookValidationState) ooxmlText(path, value string, maximum int, required bool) {
	state.text(path, value, maximum, required)
	if utf16Length(value) > maximum {
		state.issue("LIMIT_EXCEEDED", path, fmt.Sprintf("exceeds its OOXML limit of %d UTF-16 code units", maximum))
	}
}

func validateNativeWorkbookDeep(workbook *NativeWorkbookV1, issues *[]NativeWorkbookValidationIssue) {
	state := nativeWorkbookValidationState{issues: issues}
	if workbook.Source.Authority != "exact-package-bytes" {
		state.issue("INVALID_VALUE", "/source/authority", "authority must be exact-package-bytes")
	}
	state.text("/document_id", workbook.DocumentID, 256, true)
	state.text("/source/workbook_part", workbook.Source.WorkbookPart, nativeWorkbookMaxMetadataLength, true)
	if workbook.Sheets == nil {
		state.issue("REQUIRED", "/sheets", "must be a non-null collection")
	}
	if len(workbook.Sheets) == 0 {
		state.issue("REQUIRED", "/sheets", "must contain at least one sheet")
	}
	if workbook.Styles == nil || workbook.Capabilities == nil || workbook.PassthroughParts == nil || workbook.Unsupported == nil {
		state.issue("REQUIRED", "/", "styles, capabilities, passthrough_parts, and unsupported must be non-null collections")
	}

	totalRows, totalColumns, totalCells, totalMergedRanges := 0, 0, 0, 0
	for sheetIndex := range workbook.Sheets {
		sheet := &workbook.Sheets[sheetIndex]
		path := "/sheets/" + strconv.Itoa(sheetIndex)
		state.text(path+"/id", sheet.ID, 10, true)
		state.ooxmlText(path+"/name", sheet.Name, 31, true)
		if strings.ContainsAny(sheet.Name, `[]:*?/\`) || strings.IndexFunc(sheet.Name, func(r rune) bool { return r < 0x20 }) >= 0 {
			state.issue("INVALID_VALUE", path+"/name", "sheet name contains an invalid Excel character")
		}
		state.text(path+"/part_name", sheet.PartName, nativeWorkbookMaxMetadataLength, true)
		if sheet.Rows == nil || sheet.Columns == nil || sheet.Cells == nil || sheet.MergedRanges == nil {
			state.issue("REQUIRED", path, "rows, columns, cells, and merged_ranges must be non-null collections")
		}
		if sheet.Editable && sheet.RefusalCode != nil {
			state.issue("INVALID_UNION", path+"/refusal_code", "editable sheets cannot carry a refusal code")
		}
		if !sheet.Editable && (sheet.RefusalCode == nil || *sheet.RefusalCode == "") {
			state.issue("REQUIRED", path+"/refusal_code", "non-editable sheets require a refusal code")
		}
		if sheet.RefusalCode != nil {
			state.text(path+"/refusal_code", *sheet.RefusalCode, 128, true)
			if !nativeWorkbookUnsupportedCode.MatchString(*sheet.RefusalCode) {
				state.issue("INVALID_VALUE", path+"/refusal_code", "invalid refusal code")
			}
		}
		if sheet.SheetFormat != nil {
			formatPath := path + "/sheet_format"
			format := sheet.SheetFormat
			if format.DefaultRowHeightPoints == 0 && math.Signbit(format.DefaultRowHeightPoints) {
				state.issue("INVALID_NUMBER", formatPath+"/default_row_height_points", "negative zero is not canonical")
			}
			if format.DefaultColumnWidth != nil && *format.DefaultColumnWidth == 0 && math.Signbit(*format.DefaultColumnWidth) {
				state.issue("INVALID_NUMBER", formatPath+"/default_column_width", "negative zero is not canonical")
			}
		}
		totalRows += len(sheet.Rows)
		totalColumns += len(sheet.Columns)
		totalCells += len(sheet.Cells)
		totalMergedRanges += len(sheet.MergedRanges)
		if len(sheet.Rows) > NativeXLSXMaxRows || len(sheet.Columns) > NativeXLSXMaxColumns || len(sheet.Cells) > NativeXLSXMaxCells || len(sheet.MergedRanges) > NativeXLSXMaxMergedRanges {
			state.issue("LIMIT_EXCEEDED", path, "sheet collection exceeds its resource bound")
		}
		for index := range sheet.Rows {
			row := &sheet.Rows[index]
			rowPath := path + "/rows/" + strconv.Itoa(index)
			if row.CustomHeight && row.HeightPoints == nil {
				state.issue("INVALID_UNION", rowPath, "custom_height requires height_points")
			}
			if row.HeightPoints != nil && *row.HeightPoints == 0 && math.Signbit(*row.HeightPoints) {
				state.issue("INVALID_NUMBER", rowPath+"/height_points", "negative zero is not canonical")
			}
		}
		for index := range sheet.Columns {
			column := &sheet.Columns[index]
			columnPath := path + "/columns/" + strconv.Itoa(index)
			if column.CustomWidth && column.Width == nil {
				state.issue("INVALID_UNION", columnPath, "custom_width requires width")
			}
			if column.Width != nil && *column.Width == 0 && math.Signbit(*column.Width) {
				state.issue("INVALID_NUMBER", columnPath+"/width", "negative zero is not canonical")
			}
		}
		for cellIndex := range sheet.Cells {
			validateNativeWorkbookCell(&state, path+"/cells/"+strconv.Itoa(cellIndex), &sheet.Cells[cellIndex], sheet)
		}
	}
	if totalRows > NativeXLSXMaxRows || totalColumns > NativeXLSXMaxCells || totalCells > NativeXLSXMaxCells || totalMergedRanges > NativeXLSXMaxMergedRanges {
		state.issue("LIMIT_EXCEEDED", "/sheets", "cumulative row, column, cell, or merged-range inventory exceeds its resource bound")
	}

	validateNativeWorkbookStyles(&state, workbook.Styles)
	validateNativeWorkbookCapabilities(&state, workbook.Capabilities)
	validateNativeWorkbookPassthrough(&state, workbook.PassthroughParts)
	validateNativeWorkbookUnsupported(&state, workbook)
	if state.textBytes > NativeXLSXMaxTextLength {
		state.issue("LIMIT_EXCEEDED", "/", fmt.Sprintf("contract strings exceed %d cumulative bytes", NativeXLSXMaxTextLength))
	}
}

func validateNativeWorkbookCell(state *nativeWorkbookValidationState, path string, cell *NativeWorkbookCellV1, sheet *NativeWorkbookSheetV1) {
	state.text(path+"/ref", cell.Ref, 16, true)
	storage := "n"
	if cell.OOXMLType != nil {
		storage = *cell.OOXMLType
	}
	if cell.Value != nil {
		validateNativeWorkbookValue(state, path+"/value", cell.Value, storage)
	}
	if cell.Formula == nil {
		if cell.Value == nil && (storage == "s" || storage == "inlineStr") {
			state.issue("REQUIRED", path+"/value", "shared and inline string cells require a value")
		}
		return
	}
	formula := cell.Formula
	formulaPath := path + "/formula"
	state.ooxmlText(formulaPath+"/text", formula.Text, maxFormulaLen-1, formula.Type == "normal" || formula.Type == "array")
	if formula.Cached != nil {
		validateNativeWorkbookValue(state, formulaPath+"/cached", formula.Cached, storage)
	}
	if storage == "s" || storage == "inlineStr" {
		state.issue("INVALID_UNION", path+"/ooxml_type", "formula cells cannot use shared or inline string storage")
	}
	if formula.Type != "normal" && cell.Editable {
		state.issue("INVALID_VALUE", path+"/editable", "group formula cells must be mutation-refused")
	}
	switch formula.Type {
	case "normal":
		if formula.Ref != nil || formula.SharedIndex != nil {
			state.issue("INVALID_UNION", formulaPath, "normal formulas cannot carry ref or shared_index")
		}
	case "shared":
		if formula.SharedIndex == nil {
			state.issue("REQUIRED", formulaPath+"/shared_index", "shared formulas require shared_index")
		}
		if formula.Text == "" && formula.Ref != nil {
			state.issue("INVALID_UNION", formulaPath+"/ref", "shared followers cannot carry a range")
		}
		if formula.Text != "" && formula.Ref == nil {
			state.issue("REQUIRED", formulaPath+"/ref", "shared masters require a range")
		}
	case "array", "dataTable":
		if formula.SharedIndex != nil {
			state.issue("INVALID_UNION", formulaPath+"/shared_index", "array and data-table formulas cannot carry shared_index")
		}
		if formula.Ref == nil {
			state.issue("REQUIRED", formulaPath+"/ref", "formula group requires a range")
		}
	}
	if formula.Ref != nil {
		state.text(formulaPath+"/ref", *formula.Ref, 64, true)
		minimumRow, minimumColumn, maximumRow, maximumColumn, err := parseDimensionReference(*formula.Ref)
		if err != nil {
			state.issue("INVALID_REFERENCE", formulaPath+"/ref", err.Error())
		} else if canonical, canonicalErr := canonicalNativeRangeReference(*formula.Ref); canonicalErr != nil || canonical != *formula.Ref {
			state.issue("INVALID_REFERENCE", formulaPath+"/ref", "formula range must use canonical bounded A1 notation")
		} else if cell.Row < minimumRow || cell.Row > maximumRow || cell.Column < minimumColumn || cell.Column > maximumColumn {
			state.issue("INVALID_REFERENCE", formulaPath+"/ref", "formula cell is outside its group range")
		}
	}
	if formula.Type != "normal" && sheet.Editable {
		state.issue("INVALID_VALUE", path+"/editable", "a sheet containing group formulas must be mutation-refused")
	}
}

func validateNativeWorkbookValue(state *nativeWorkbookValidationState, path string, value *NativeWorkbookValueV1, ooxmlType string) {
	state.text(path+"/kind", value.Kind, 32, true)
	state.text(path+"/storage", value.Storage, 32, true)
	requireLexical, requireText := true, false
	expectedKind, expectedType := "", ""
	allowRich := false
	switch value.Storage {
	case "number":
		expectedKind, expectedType = "number", "n"
	case "boolean":
		expectedKind, expectedType = "boolean", "b"
	case "error":
		expectedKind, expectedType = "error", "e"
	case "date":
		expectedKind, expectedType = "date", "d"
	case "formula-string":
		expectedKind, expectedType, requireText = "string", "str", true
	case "shared":
		expectedKind, expectedType, requireText, allowRich = "string", "s", true, true
	case "inline":
		expectedKind, expectedType, requireLexical, requireText, allowRich = "string", "inlineStr", false, true, true
	default:
		state.issue("INVALID_VALUE", path+"/storage", "invalid native value storage")
	}
	if value.Kind != expectedKind {
		state.issue("INVALID_UNION", path+"/kind", "kind does not match storage")
	}
	if ooxmlType == "" {
		ooxmlType = "n"
	}
	if ooxmlType != expectedType {
		state.issue("INVALID_UNION", path, "value storage does not match the cell OOXML type")
	}
	if requireLexical != (value.Lexical != nil) {
		state.issue("INVALID_UNION", path+"/lexical", "lexical presence does not match storage")
	}
	if requireText != (value.Text != nil) {
		state.issue("INVALID_UNION", path+"/text", "text presence does not match storage")
	}
	if value.Rich && !allowRich {
		state.issue("INVALID_UNION", path+"/rich", "rich is only valid for shared or inline strings")
	}
	if value.Lexical != nil {
		state.text(path+"/lexical", *value.Lexical, maxCellTextLen, true)
		validateNativeWorkbookLexical(state, path+"/lexical", value.Storage, *value.Lexical)
	}
	if value.Text != nil {
		state.ooxmlText(path+"/text", *value.Text, maxCellTextLen, false)
	}
	if value.Storage == "formula-string" && value.Lexical != nil && value.Text != nil {
		decoded, err := decodeSpreadsheetString(*value.Lexical)
		if err != nil || decoded != *value.Text {
			state.issue("INVALID_UNION", path, "formula-string lexical and decoded text disagree")
		}
	}
}

func validateNativeWorkbookLexical(state *nativeWorkbookValidationState, path, storage, lexical string) {
	switch storage {
	case "number":
		if _, err := finiteNativeFloat(lexical, -math.MaxFloat64, math.MaxFloat64); err != nil {
			state.issue("INVALID_VALUE", path, "invalid finite OOXML numeric lexical")
		}
	case "boolean":
		if lexical != "0" && lexical != "1" && lexical != "false" && lexical != "true" {
			state.issue("INVALID_VALUE", path, "invalid OOXML boolean lexical")
		}
	case "error":
		if lexical == "" || strings.IndexFunc(lexical, func(r rune) bool { return r < 0x20 && r != '\t' }) >= 0 {
			state.issue("INVALID_VALUE", path, "invalid OOXML error lexical")
		}
	case "date":
		if !validNativeISODateTime(lexical) {
			state.issue("INVALID_VALUE", path, "date cells require an ISO-8601 date or dateTime lexical")
		}
	case "shared":
		if _, err := strconv.ParseUint(lexical, 10, 32); err != nil || strings.TrimSpace(lexical) != lexical {
			state.issue("INVALID_VALUE", path, "shared-string lexical must be an unsigned decimal index")
		}
	}
}

func validNativeISODateTime(value string) bool {
	if nativeWorkbookISODate.MatchString(value) {
		layout := "2006-01-02"
		if len(value) > len("2006-01-02") {
			layout = "2006-01-02Z07:00"
		}
		_, err := time.Parse(layout, value)
		return err == nil
	}
	if nativeWorkbookISOLocalDateTime.MatchString(value) {
		layout := "2006-01-02T15:04:05"
		if strings.Contains(value, ".") {
			layout = "2006-01-02T15:04:05.999999999"
		}
		_, err := time.Parse(layout, value)
		return err == nil
	}
	if nativeWorkbookISOOffsetDateTime.MatchString(value) {
		_, err := time.Parse(time.RFC3339Nano, value)
		return err == nil
	}
	return false
}

func validateNativeWorkbookStyles(state *nativeWorkbookValidationState, styles []NativeWorkbookStyleV1) {
	allowedUnsupported := map[string]bool{"alignment-extended": true, "border": true, "fill": true, "font-color": true, "horizontal-alignment": true, "number-format": true, "vertical-alignment": true}
	allowedBorderStyles := map[string]bool{
		"dashDot": true, "dashDotDot": true, "dashed": true, "dotted": true, "double": true, "hair": true,
		"medium": true, "mediumDashDot": true, "mediumDashDotDot": true, "mediumDashed": true,
		"slantDashDot": true, "thick": true, "thin": true,
	}
	for index := range styles {
		entry := &styles[index]
		style := &entry.Effective
		path := "/styles/" + strconv.Itoa(index) + "/effective"
		if !nativeWorkbookSHA.MatchString(entry.RawProjectionSHA256) {
			state.issue("INVALID_SHA256", "/styles/"+strconv.Itoa(index)+"/raw_projection_sha256", "style raw projection requires a full SHA-256")
		} else if digest, err := nativeRawStyleProjectionDigest(*style); err != nil || digest != entry.RawProjectionSHA256 {
			state.issue("INVALID_SHA256", "/styles/"+strconv.Itoa(index)+"/raw_projection_sha256", "style raw projection digest does not match the canonical effective-style bytes")
		}
		if style.Unsupported == nil {
			state.issue("REQUIRED", path+"/unsupported", "must be a non-null collection")
		}
		last := ""
		for unsupportedIndex, code := range style.Unsupported {
			codePath := path + "/unsupported/" + strconv.Itoa(unsupportedIndex)
			state.text(codePath, code, 64, true)
			if !allowedUnsupported[code] {
				state.issue("INVALID_VALUE", codePath, "unsupported style projection code")
			}
			if code <= last {
				state.issue("INVALID_ORDER", codePath, "unsupported style codes must be unique and sorted")
			}
			last = code
		}
		if style.Projection != "full" && style.Projection != "partial" {
			state.issue("INVALID_VALUE", path+"/projection", "projection must be full or partial")
		}
		if (style.Projection == "full") != (len(style.Unsupported) == 0) {
			state.issue("INVALID_UNION", path, "projection must match unsupported components")
		}
		borderUnsupported := false
		fillUnsupported := false
		for _, code := range style.Unsupported {
			if code == "border" {
				borderUnsupported = true
			}
			if code == "fill" {
				fillUnsupported = true
			}
		}
		if fillUnsupported != (style.Fill == nil) {
			state.issue("INVALID_UNION", path+"/fill", "fill provenance must be present exactly when fill projection is supported")
		}
		if style.Fill != nil {
			fill := style.Fill
			if fill.Origin != "implicit-default" && fill.Origin != "styles-record" {
				state.issue("INVALID_VALUE", path+"/fill/origin", "fill origin must be implicit-default or styles-record")
			}
			if fill.Origin == "implicit-default" {
				if fill.FillID != nil || fill.RecordSHA256 != nil || fill.Color != nil || style.FillColor != nil {
					state.issue("INVALID_UNION", path+"/fill", "implicit-default fill cannot carry a styles-table record or color")
				}
			} else if fill.FillID == nil || fill.RecordSHA256 == nil {
				state.issue("REQUIRED", path+"/fill", "styles-record fill requires fill_id and record_sha256")
			}
			if fill.RecordSHA256 != nil && !nativeWorkbookSHA.MatchString(*fill.RecordSHA256) {
				state.issue("INVALID_SHA256", path+"/fill/record_sha256", "fill record requires a full SHA-256")
			}
			if (fill.Color == nil) != (style.FillColor == nil) || (fill.Color != nil && style.FillColor != nil && *fill.Color != *style.FillColor) {
				state.issue("INVALID_UNION", path+"/fill_color", "effective fill color must exactly match fill provenance")
			}
		}
		if borderUnsupported != (style.Border == nil) {
			state.issue("INVALID_UNION", path+"/border", "border must be present exactly when border projection is supported")
		}
		if style.Border != nil {
			border := style.Border
			if border.Origin != "implicit-default" && border.Origin != "styles-record" {
				state.issue("INVALID_VALUE", path+"/border/origin", "border origin must be implicit-default or styles-record")
			}
			if border.Origin == "implicit-default" {
				if border.BorderID != nil || border.RecordSHA256 != nil || border.Left != nil || border.Right != nil || border.Top != nil || border.Bottom != nil {
					state.issue("INVALID_UNION", path+"/border", "implicit-default border cannot carry a styles-table record or visible sides")
				}
			} else if border.BorderID == nil || border.RecordSHA256 == nil {
				state.issue("REQUIRED", path+"/border", "styles-record border requires border_id and record_sha256")
			}
			if border.RecordSHA256 != nil && !nativeWorkbookSHA.MatchString(*border.RecordSHA256) {
				state.issue("INVALID_SHA256", path+"/border/record_sha256", "border record requires a full SHA-256")
			}
			for _, side := range []struct {
				name string
				side *NativeWorkbookBorderSideV1
			}{{"left", border.Left}, {"right", border.Right}, {"top", border.Top}, {"bottom", border.Bottom}} {
				if side.side == nil {
					continue
				}
				if !allowedBorderStyles[side.side.Style] {
					state.issue("INVALID_VALUE", path+"/border/"+side.name+"/style", "unsupported OOXML border style token")
				}
				if !nativeWorkbookColorPattern.MatchString(side.side.Color) {
					state.issue("INVALID_VALUE", path+"/border/"+side.name+"/color", "border color must be canonical #RRGGBB")
				}
			}
		}
		if style.NumberFormat != nil {
			state.ooxmlText(path+"/number_format", *style.NumberFormat, maxNumberFormatLength, true)
		}
		if style.FontName != nil {
			state.ooxmlText(path+"/font_name", *style.FontName, maxFontNameLength, true)
		}
		if style.FontSizePoints != nil && (math.IsNaN(*style.FontSizePoints) || math.IsInf(*style.FontSizePoints, 0) || *style.FontSizePoints <= 0 || *style.FontSizePoints > maxFontSizePoints) {
			state.issue("INVALID_NUMBER", path+"/font_size_points", "font size is outside native mutation bounds")
		}
		for _, field := range []struct {
			name  string
			color *string
		}{{name: "font_color", color: style.FontColor}, {name: "fill_color", color: style.FillColor}} {
			if field.color != nil && !nativeWorkbookColorPattern.MatchString(*field.color) {
				state.issue("INVALID_VALUE", path+"/"+field.name, "color must be canonical #RRGGBB")
			}
		}
		if style.HorizontalAlignment != nil && *style.HorizontalAlignment != "general" && *style.HorizontalAlignment != "left" && *style.HorizontalAlignment != "center" && *style.HorizontalAlignment != "right" {
			state.issue("INVALID_VALUE", path+"/horizontal_alignment", "unsupported horizontal alignment")
		}
		if style.VerticalAlignment != nil && *style.VerticalAlignment != "top" && *style.VerticalAlignment != "middle" && *style.VerticalAlignment != "bottom" {
			state.issue("INVALID_VALUE", path+"/vertical_alignment", "unsupported vertical alignment")
		}
	}
}

func validateNativeWorkbookCapabilities(state *nativeWorkbookValidationState, capabilities []NativeWorkbookCapabilityV1) {
	expected := []struct{ name, level string }{
		{name: "native-ooxml-parse", level: "read-only"},
		{name: "native-v1-mutations", level: "partial"},
		{name: "unsupported-content", level: "preserve-exact"},
	}
	if len(capabilities) != len(expected) {
		state.issue("INVALID_VALUE", "/capabilities", "capabilities must contain the canonical v1 inventory")
	}
	seen := map[string]bool{}
	for index := range capabilities {
		capability := &capabilities[index]
		path := "/capabilities/" + strconv.Itoa(index)
		state.text(path+"/name", capability.Name, 128, true)
		if !nativeWorkbookCapabilityName.MatchString(capability.Name) || seen[capability.Name] {
			state.issue("INVALID_VALUE", path+"/name", "capability name is invalid or duplicated")
		}
		seen[capability.Name] = true
		if index >= len(expected) || capability.Name != expected[index].name || capability.Level != expected[index].level {
			state.issue("INVALID_VALUE", path, "capability name, level, or order is not canonical for v1")
		}
		if capability.Detail != nil {
			state.text(path+"/detail", *capability.Detail, nativeWorkbookMaxMetadataLength, true)
		}
	}
}

func validateNativeWorkbookPassthrough(state *nativeWorkbookValidationState, parts []NativeWorkbookPassthroughPartV1) {
	seen, previous := map[string]bool{}, ""
	for index := range parts {
		part := &parts[index]
		path := "/passthrough_parts/" + strconv.Itoa(index)
		state.text(path+"/part_name", part.PartName, nativeWorkbookMaxMetadataLength, true)
		state.text(path+"/content_type", part.ContentType, 1024, true)
		key, err := canonicalOPCPartKey(part.PartName)
		if err == nil {
			if seen[key] {
				state.issue("DUPLICATE_PART", path+"/part_name", "duplicate case/escape-equivalent passthrough part")
			}
			seen[key] = true
		}
		orderKey := asciiLower(part.PartName)
		if index > 0 && orderKey <= previous {
			state.issue("INVALID_ORDER", path+"/part_name", "passthrough parts must be uniquely ASCII-case-sorted")
		}
		previous = orderKey
		if strings.TrimSpace(part.ContentType) != part.ContentType || strings.ContainsAny(part.ContentType, "\r\n\x00") {
			state.issue("INVALID_VALUE", path+"/content_type", "invalid content type")
		}
		if part.ByteLength == nil || *part.ByteLength < 0 || *part.ByteLength > NativeXLSXMaxPartBytes {
			state.issue("INVALID_NUMBER", path+"/byte_length", "byte length is outside package bounds")
		}
	}
}

func validateNativeWorkbookUnsupported(state *nativeWorkbookValidationState, workbook *NativeWorkbookV1) {
	allowedCapabilities := map[string]bool{
		"cell-markup": true, "cell-metadata": true, "charts": true, "comments": true, "conditional-formatting": true,
		"data-validation": true, "dimensions": true, "drawings": true, "embeddings": true,
		"extensions": true, "external-links": true, "formula-groups": true, "hyperlinks": true,
		"macros": true, "merges": true, "opaque-parts": true, "pivots": true, "protection": true,
		"rich-text": true, "styles": true, "tables": true, "workbook-features": true, "worksheet-features": true,
	}
	sheetsByID := map[string]NativeWorkbookSheetV1{}
	cellsBySheet := map[string]map[string]NativeWorkbookCellV1{}
	expectedCellCodes := map[string]map[string]map[string]bool{}
	expectedFormulaRanges := map[string]map[string]bool{}
	mergedRangesBySheet := map[string][]nativeFormulaGroupRange{}
	mergedTopLeftBySheet := map[string]map[string]bool{}
	observedStyleCodes := make([]map[string]bool, len(workbook.Styles))
	for index := range observedStyleCodes {
		observedStyleCodes[index] = map[string]bool{}
	}
	for _, sheet := range workbook.Sheets {
		sheetsByID[sheet.ID] = sheet
		cellsBySheet[sheet.ID] = map[string]NativeWorkbookCellV1{}
		expectedCellCodes[sheet.ID] = map[string]map[string]bool{}
		expectedFormulaRanges[sheet.ID] = map[string]bool{}
		mergedTopLeftBySheet[sheet.ID] = map[string]bool{}
		for _, merged := range sheet.MergedRanges {
			if merged.Row < 0 || merged.Row >= excelMaxRows || merged.Column < 0 || merged.Column >= excelMaxColumns || merged.EndRow < merged.Row || merged.EndRow >= excelMaxRows || merged.EndColumn < merged.Column || merged.EndColumn >= excelMaxColumns {
				continue
			}
			mergedRangesBySheet[sheet.ID] = append(mergedRangesBySheet[sheet.ID], nativeFormulaGroupRange{minimumRow: merged.Row, minimumColumn: merged.Column, maximumRow: merged.EndRow, maximumColumn: merged.EndColumn, ref: merged.Ref})
			mergedTopLeftBySheet[sheet.ID][cellReference(merged.Row, merged.Column)] = true
		}
		for _, cell := range sheet.Cells {
			cellsBySheet[sheet.ID][cell.Ref] = cell
			expectedCellCodes[sheet.ID][cell.Ref] = map[string]bool{}
			if cell.Value != nil && cell.Value.Rich {
				expectedCellCodes[sheet.ID][cell.Ref]["RICH_CELL_STRING"] = true
			}
			if cell.Formula != nil && cell.Formula.Type != "normal" {
				code := "FORMULA_" + strings.ToUpper(cell.Formula.Type)
				expectedCellCodes[sheet.ID][cell.Ref][code] = true
				if cell.Formula.Ref != nil {
					expectedFormulaRanges[sheet.ID][*cell.Formula.Ref] = true
				}
			}
		}
	}
	cellReasons := map[string]map[string]bool{}
	observedCellCodes := map[string]map[string]map[string]bool{}
	observedFormulaRanges := map[string]map[string]bool{}
	rangeReasons := map[string][]nativeFormulaGroupRange{}
	sheetReasons := map[string]map[string]bool{}
	mergedSourceSheets := map[string]bool{}
	for _, sheet := range workbook.Sheets {
		id := sheet.ID
		cellReasons[id], sheetReasons[id] = map[string]bool{}, map[string]bool{}
		observedCellCodes[id], observedFormulaRanges[id] = map[string]map[string]bool{}, map[string]bool{}
	}
	seenID, seenLocation := map[string]bool{}, map[string]bool{}
	for index := range workbook.Unsupported {
		item := &workbook.Unsupported[index]
		path := "/unsupported/" + strconv.Itoa(index)
		for _, field := range []struct{ suffix, value string }{
			{suffix: "id", value: item.ID}, {suffix: "code", value: item.Code}, {suffix: "capability", value: item.Capability},
			{suffix: "scope_id", value: item.ScopeID}, {suffix: "preservation", value: item.Preservation}, {suffix: "message", value: item.Message},
		} {
			state.text(path+"/"+field.suffix, field.value, nativeWorkbookMaxMetadataLength, true)
		}
		if !nativeWorkbookUnsupportedID.MatchString(item.ID) || seenID[item.ID] {
			state.issue("INVALID_VALUE", path+"/id", "unsupported id is invalid or duplicated")
		}
		seenID[item.ID] = true
		if !nativeWorkbookUnsupportedCode.MatchString(item.Code) || !nativeWorkbookCapabilityName.MatchString(item.Capability) || !allowedCapabilities[item.Capability] {
			state.issue("INVALID_VALUE", path, "unsupported code or capability is invalid")
		}
		class, classified := nativeUnsupportedClassificationForCode(item.Code)
		if !classified || class.capability != item.Capability {
			state.issue("INVALID_VALUE", path, "unsupported code and capability are not a canonical v1 pair")
		}
		if item.Preservation != "preserve-exact" {
			state.issue("INVALID_VALUE", path+"/preservation", "preservation must be preserve-exact")
		}
		if item.ScopeID != "workbook" && !strings.HasPrefix(item.ScopeID, "sheet:") && !strings.HasPrefix(item.ScopeID, "style:") {
			state.issue("INVALID_VALUE", path+"/scope_id", "unsupported scope")
		}
		var scopedSheet *NativeWorkbookSheetV1
		scopedStyle := -1
		if strings.HasPrefix(item.ScopeID, "sheet:") {
			id := strings.TrimPrefix(item.ScopeID, "sheet:")
			if sheet, found := sheetsByID[id]; found {
				scopedSheet = &sheet
			} else {
				state.issue("INVALID_REFERENCE", path+"/scope_id", "sheet scope does not exist")
			}
		}
		if strings.HasPrefix(item.ScopeID, "style:") {
			styleLexical := strings.TrimPrefix(item.ScopeID, "style:")
			styleIndex, err := strconv.Atoi(styleLexical)
			if err != nil || styleIndex < 0 || styleIndex >= len(workbook.Styles) || styleLexical != strconv.Itoa(styleIndex) {
				state.issue("INVALID_REFERENCE", path+"/scope_id", "style scope does not exist")
			} else {
				scopedStyle = styleIndex
			}
		}
		if classified && class.scope == "style" && scopedStyle >= 0 {
			observedStyleCodes[scopedStyle][item.Code] = true
		}
		if item.PartName == nil {
			state.issue("REQUIRED", path+"/part_name", "unsupported content must identify its source-authoritative part")
		}
		if item.PartName != nil {
			state.text(path+"/part_name", *item.PartName, nativeWorkbookMaxMetadataLength, true)
			if _, err := canonicalOPCPartKey(*item.PartName); err != nil {
				state.issue("INVALID_PART", path+"/part_name", err.Error())
			}
			if scopedSheet != nil && item.Code != "SHEET_DECLARATION_ATTRIBUTES" && classified && class.scope == "sheet" && *item.PartName != scopedSheet.PartName {
				state.issue("INVALID_REFERENCE", path+"/part_name", "sheet-scoped unsupported content must use the exact worksheet part spelling")
			}
			if item.Code == "SHEET_DECLARATION_ATTRIBUTES" && *item.PartName != workbook.Source.WorkbookPart {
				state.issue("INVALID_REFERENCE", path+"/part_name", "sheet-declaration unsupported content must identify the exact workbook part")
			}
		}
		if item.CellRef != nil && item.RangeRef != nil {
			state.issue("INVALID_UNION", path, "unsupported location cannot contain both cell_ref and range_ref")
		}
		if (item.CellRef != nil || item.RangeRef != nil) && scopedSheet == nil {
			state.issue("INVALID_REFERENCE", path, "cell/range locations require a valid sheet scope")
		}
		if item.CellRef != nil {
			state.text(path+"/cell_ref", *item.CellRef, 16, true)
			row, column, err := parseCellReference(*item.CellRef)
			if err != nil || cellReference(row, column) != *item.CellRef {
				state.issue("INVALID_REFERENCE", path+"/cell_ref", "cell_ref must be a canonical bounded A1 reference")
			}
		}
		if item.RangeRef != nil {
			state.text(path+"/range_ref", *item.RangeRef, 64, true)
			if canonical, err := canonicalNativeRangeReference(*item.RangeRef); err != nil {
				state.issue("INVALID_REFERENCE", path+"/range_ref", err.Error())
			} else if canonical != *item.RangeRef {
				state.issue("INVALID_REFERENCE", path+"/range_ref", "range_ref must use canonical bounded A1 notation")
			}
		}
		if classified {
			validateNativeUnsupportedClassLocation(state, path, item, class)
			if scopedSheet != nil {
				sheetID := strings.TrimPrefix(item.ScopeID, "sheet:")
				switch class.impact {
				case nativeUnsupportedCellImpact:
					if item.CellRef != nil {
						cell, found := cellsBySheet[sheetID][*item.CellRef]
						if !found {
							state.issue("INVALID_REFERENCE", path+"/cell_ref", "cell-scoped unsupported source is absent from modeled cells")
						} else {
							if cell.Editable {
								state.issue("INVALID_VALUE", path+"/cell_ref", "cell-scoped unsupported source requires that modeled cell to be mutation-refused")
							}
							validateNativeUnsupportedCellSemantics(state, path, item.Code, cell)
							cellReasons[sheetID][cell.Ref] = true
							if observedCellCodes[sheetID][cell.Ref] == nil {
								observedCellCodes[sheetID][cell.Ref] = map[string]bool{}
							}
							observedCellCodes[sheetID][cell.Ref][item.Code] = true
						}
					}
				case nativeUnsupportedRangeImpact:
					if item.RangeRef != nil {
						minimumRow, minimumColumn, maximumRow, maximumColumn, err := parseDimensionReference(*item.RangeRef)
						if err == nil {
							rangeReasons[sheetID] = append(rangeReasons[sheetID], nativeFormulaGroupRange{minimumRow: minimumRow, minimumColumn: minimumColumn, maximumRow: maximumRow, maximumColumn: maximumColumn, ref: *item.RangeRef})
						}
						sheetReasons[sheetID]["FORMULA_GROUPS"] = true
						observedFormulaRanges[sheetID][*item.RangeRef] = true
					}
				case nativeUnsupportedSheetImpact:
					sheetReasons[sheetID][class.refusalCode] = true
				}
			}
		}
		if item.Code == "MERGED_CELLS" && scopedSheet != nil {
			mergedSourceSheets[scopedSheet.ID] = true
		}
		location := strings.Join([]string{item.Code, item.Capability, item.ScopeID, nativeOptionalString(item.PartName), nativeOptionalString(item.CellRef), nativeOptionalString(item.RangeRef)}, "\x00")
		if seenLocation[location] {
			state.issue("DUPLICATE_ID", path, "duplicate unsupported location")
		}
		seenLocation[location] = true
		seedKey := strings.Join([]string{item.Code, item.Capability, item.ScopeID, nativeOptionalString(item.PartName), nativeOptionalString(item.CellRef), nativeOptionalString(item.RangeRef)}, "\x00")
		seed := sha256.Sum256([]byte(seedKey))
		if item.ID != "unsupported:"+hex.EncodeToString(seed[:]) {
			state.issue("INVALID_ID", path+"/id", "unsupported id does not match its canonical source location")
		}
	}
	styleUnsupportedCodes := map[string]string{
		"alignment-extended":   "STYLE_ALIGNMENT_EXTENDED",
		"border":               "STYLE_BORDER",
		"fill":                 "STYLE_FILL",
		"font-color":           "STYLE_FONT_COLOR",
		"horizontal-alignment": "STYLE_HORIZONTAL_ALIGNMENT",
		"number-format":        "STYLE_NUMBER_FORMAT",
		"vertical-alignment":   "STYLE_VERTICAL_ALIGNMENT",
	}
	for styleIndex := range workbook.Styles {
		expected := map[string]bool{}
		for _, effectiveCode := range workbook.Styles[styleIndex].Effective.Unsupported {
			if diagnosticCode := styleUnsupportedCodes[effectiveCode]; diagnosticCode != "" {
				expected[diagnosticCode] = true
			}
		}
		for diagnosticCode := range expected {
			if !observedStyleCodes[styleIndex][diagnosticCode] {
				state.issue("REQUIRED", "/styles/"+strconv.Itoa(styleIndex)+"/effective/unsupported", "style projection lacks its canonical "+diagnosticCode+" source diagnostic")
			}
		}
		for diagnosticCode := range observedStyleCodes[styleIndex] {
			if !expected[diagnosticCode] {
				state.issue("INVALID_REFERENCE", "/unsupported", diagnosticCode+" contradicts the effective style projection for style:"+strconv.Itoa(styleIndex))
			}
		}
	}
	for _, sheet := range workbook.Sheets {
		if len(rangeReasons[sheet.ID]) > 1 {
			if err := rejectOverlappingNativeFormulaRanges(rangeReasons[sheet.ID]); err != nil {
				state.issue("INVALID_UNION", "/sheets/"+strconv.Itoa(sheet.Order), err.Error())
			}
		}
		covered := nativeCellsCoveredByUnsupportedRanges(sheet.Cells, rangeReasons[sheet.ID])
		mergeCovered := nativeCellsCoveredByUnsupportedRanges(sheet.Cells, mergedRangesBySheet[sheet.ID])
		if (len(sheet.MergedRanges) > 0) != mergedSourceSheets[sheet.ID] {
			state.issue("INVALID_REFERENCE", "/sheets/"+strconv.Itoa(sheet.Order)+"/merged_ranges", "modeled merged ranges and their exact MERGED_CELLS source diagnostic must be present together")
		}
		for _, cell := range sheet.Cells {
			if covered[cell.Ref] && cell.Editable {
				state.issue("INVALID_VALUE", "/sheets/"+strconv.Itoa(sheet.Order)+"/cells/"+cell.Ref+"/editable", "range-scoped unsupported source requires covered modeled cells to be mutation-refused")
			}
			if mergeCovered[cell.Ref] && cell.Editable {
				state.issue("INVALID_VALUE", "/sheets/"+strconv.Itoa(sheet.Order)+"/cells/"+cell.Ref+"/editable", "modeled cells covered by a merged range must be mutation-refused")
			}
			if mergeCovered[cell.Ref] && !mergedTopLeftBySheet[sheet.ID][cell.Ref] && (cell.Value != nil || cell.Formula != nil) {
				state.issue("INVALID_UNION", "/sheets/"+strconv.Itoa(sheet.Order)+"/cells/"+cell.Ref, "non-anchor merged cells cannot carry value or formula authority")
			}
			if !cell.Editable && !cellReasons[sheet.ID][cell.Ref] && !covered[cell.Ref] && !mergeCovered[cell.Ref] {
				state.issue("REQUIRED", "/sheets/"+strconv.Itoa(sheet.Order)+"/cells/"+cell.Ref+"/editable", "non-editable modeled cell requires an authoritative unsupported or merged-range reason")
			}
			for _, code := range sortedNativeStringKeys(expectedCellCodes[sheet.ID][cell.Ref]) {
				if !observedCellCodes[sheet.ID][cell.Ref][code] {
					state.issue("REQUIRED", "/sheets/"+strconv.Itoa(sheet.Order)+"/cells/"+cell.Ref, "modeled rich/formula refusal lacks its canonical unsupported source")
				}
			}
		}
		expectedRangeRefs := sortedNativeStringKeys(expectedFormulaRanges[sheet.ID])
		for _, ref := range expectedRangeRefs {
			if !observedFormulaRanges[sheet.ID][ref] {
				state.issue("REQUIRED", "/sheets/"+strconv.Itoa(sheet.Order), "modeled formula group lacks its exact range-scoped unsupported source")
			}
		}
		observedRangeRefs := sortedNativeStringKeys(observedFormulaRanges[sheet.ID])
		for _, ref := range observedRangeRefs {
			if !expectedFormulaRanges[sheet.ID][ref] {
				state.issue("INVALID_REFERENCE", "/sheets/"+strconv.Itoa(sheet.Order), "range-scoped formula refusal does not match a modeled formula-group range")
			}
		}
		expectedRefusal := nativeExpectedSheetRefusal(sheetReasons[sheet.ID])
		if expectedRefusal == "" {
			if !sheet.Editable || sheet.RefusalCode != nil {
				state.issue("INVALID_VALUE", "/sheets/"+strconv.Itoa(sheet.Order), "sheet refusal has no authoritative unsupported source")
			}
		} else if sheet.Editable || sheet.RefusalCode == nil || *sheet.RefusalCode != expectedRefusal {
			state.issue("INVALID_VALUE", "/sheets/"+strconv.Itoa(sheet.Order)+"/refusal_code", "sheet refusal does not match its authoritative unsupported source")
		}
	}
}

func canonicalNativeRangeReference(ref string) (string, error) {
	minimumRow, minimumColumn, maximumRow, maximumColumn, err := parseDimensionReference(ref)
	if err != nil {
		return "", err
	}
	canonical := cellReference(minimumRow, minimumColumn)
	if minimumRow != maximumRow || minimumColumn != maximumColumn {
		canonical += ":" + cellReference(maximumRow, maximumColumn)
	}
	return canonical, nil
}

func sortedNativeStringKeys(values map[string]bool) []string {
	keys := make([]string, 0, len(values))
	for key := range values {
		keys = append(keys, key)
	}
	sort.Strings(keys)
	return keys
}

type nativeUnsupportedImpact int

const (
	nativeUnsupportedNoImpact nativeUnsupportedImpact = iota
	nativeUnsupportedCellImpact
	nativeUnsupportedRangeImpact
	nativeUnsupportedSheetImpact
)

type nativeUnsupportedClassification struct {
	capability, scope, refusalCode string
	impact                         nativeUnsupportedImpact
}

func nativeUnsupportedClassificationForCode(code string) (nativeUnsupportedClassification, bool) {
	if class, found := nativeXLSXSchemaUnsupportedClassifications[code]; found {
		return class, true
	}
	return nativeUnsupportedClassification{}, false
}

func validateNativeUnsupportedClassLocation(state *nativeWorkbookValidationState, path string, item *NativeWorkbookUnsupportedV1, class nativeUnsupportedClassification) {
	actualScope := "workbook"
	if strings.HasPrefix(item.ScopeID, "sheet:") {
		actualScope = "sheet"
	} else if strings.HasPrefix(item.ScopeID, "style:") {
		actualScope = "style"
	}
	if actualScope != class.scope {
		state.issue("INVALID_REFERENCE", path+"/scope_id", "unsupported code is in the wrong scope")
	}
	if class.impact == nativeUnsupportedCellImpact && (item.CellRef == nil || item.RangeRef != nil) {
		state.issue("REQUIRED", path+"/cell_ref", "cell-impact unsupported code requires exactly cell_ref")
	}
	if class.impact == nativeUnsupportedRangeImpact && (item.RangeRef == nil || item.CellRef != nil) {
		state.issue("REQUIRED", path+"/range_ref", "range-impact unsupported code requires exactly range_ref")
	}
	if (class.impact == nativeUnsupportedNoImpact || class.impact == nativeUnsupportedSheetImpact) && (item.CellRef != nil || item.RangeRef != nil) {
		state.issue("INVALID_UNION", path, "this unsupported code cannot carry a cell/range location")
	}
}

func validateNativeUnsupportedCellSemantics(state *nativeWorkbookValidationState, path, code string, cell NativeWorkbookCellV1) {
	switch code {
	case "RICH_CELL_STRING":
		if cell.Value == nil || !cell.Value.Rich {
			state.issue("INVALID_REFERENCE", path+"/cell_ref", "rich-cell refusal does not reference a rich modeled value")
		}
	case "FORMULA_ATTRIBUTES":
		if cell.Formula == nil {
			state.issue("INVALID_REFERENCE", path+"/cell_ref", "formula-attribute refusal does not reference a modeled formula")
		}
	case "FORMULA_SHARED", "FORMULA_ARRAY", "FORMULA_DATATABLE":
		expected := strings.ToLower(strings.TrimPrefix(code, "FORMULA_"))
		if expected == "datatable" {
			expected = "dataTable"
		}
		if cell.Formula == nil || cell.Formula.Type != expected {
			state.issue("INVALID_REFERENCE", path+"/cell_ref", "formula refusal type does not match the modeled formula")
		}
	}
}

func nativeCellsCoveredByUnsupportedRanges(cells []NativeWorkbookCellV1, ranges []nativeFormulaGroupRange) map[string]bool {
	covered := map[string]bool{}
	if len(ranges) == 0 {
		return covered
	}
	type event struct{ row, minimumColumn, maximumColumn, delta int }
	events := make([]event, 0, len(ranges)*2)
	for _, group := range ranges {
		events = append(events, event{group.minimumRow, group.minimumColumn, group.maximumColumn, 1}, event{group.maximumRow + 1, group.minimumColumn, group.maximumColumn, -1})
	}
	sort.Slice(events, func(i, j int) bool {
		if events[i].row != events[j].row {
			return events[i].row < events[j].row
		}
		return events[i].delta < events[j].delta
	})
	tree, eventIndex := newNativeColumnRangeTree(excelMaxColumns), 0
	for _, cell := range cells {
		for eventIndex < len(events) && events[eventIndex].row <= cell.Row {
			item := events[eventIndex]
			tree.add(item.minimumColumn, item.maximumColumn, item.delta)
			eventIndex++
		}
		if tree.maximum(cell.Column, cell.Column) > 0 {
			covered[cell.Ref] = true
		}
	}
	return covered
}

func nativeExpectedSheetRefusal(reasons map[string]bool) string {
	for _, code := range []string{"FORMULA_GROUPS", "SHEET_PROTECTION", "SHEET_DATA_ATTRIBUTES", "UNSAFE_WORKSHEET_ATTRIBUTES", "SHEET_DECLARATION_ATTRIBUTES"} {
		if reasons[code] {
			return code
		}
	}
	return ""
}

func nativeOptionalString(value *string) string {
	if value == nil {
		return ""
	}
	return *value
}

func rejectDuplicateNativeWorkbookJSONKeys(data []byte) error {
	if !utf8.Valid(data) {
		return fmt.Errorf("JSON must be valid UTF-8")
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.UseNumber()
	tokens := 0
	if err := walkNativeWorkbookJSON(decoder, 0, "", &tokens); err != nil {
		return err
	}
	if token, err := decoder.Token(); err != io.EOF {
		if err == nil {
			return fmt.Errorf("trailing JSON token %v", token)
		}
		return err
	}
	return nil
}

func walkNativeWorkbookJSON(decoder *json.Decoder, depth int, path string, tokens *int) error {
	if depth > nativeWorkbookMaxJSONDepth {
		return fmt.Errorf("JSON nesting exceeds %d", nativeWorkbookMaxJSONDepth)
	}
	token, err := decoder.Token()
	if err != nil {
		return err
	}
	*tokens++
	if *tokens > nativeWorkbookMaxJSONTokens {
		return fmt.Errorf("JSON token count exceeds %d", nativeWorkbookMaxJSONTokens)
	}
	delimiter, isDelimiter := token.(json.Delim)
	if !isDelimiter {
		return nil
	}
	switch delimiter {
	case '{':
		seen := map[string]bool{}
		for decoder.More() {
			keyToken, keyErr := decoder.Token()
			if keyErr != nil {
				return keyErr
			}
			*tokens++
			key, ok := keyToken.(string)
			if !ok {
				return fmt.Errorf("object key is not a string")
			}
			if seen[key] {
				return fmt.Errorf("duplicate object key %q at %s", key, path)
			}
			seen[key] = true
			if err := walkNativeWorkbookJSON(decoder, depth+1, path+"/"+key, tokens); err != nil {
				return err
			}
		}
		closing, err := decoder.Token()
		if err != nil || closing != json.Delim('}') {
			return fmt.Errorf("unterminated JSON object")
		}
		*tokens++
	case '[':
		index := 0
		for decoder.More() {
			if err := walkNativeWorkbookJSON(decoder, depth+1, path+"/"+strconv.Itoa(index), tokens); err != nil {
				return err
			}
			index++
		}
		closing, err := decoder.Token()
		if err != nil || closing != json.Delim(']') {
			return fmt.Errorf("unterminated JSON array")
		}
		*tokens++
	default:
		return fmt.Errorf("unexpected JSON delimiter %q", delimiter)
	}
	return nil
}
