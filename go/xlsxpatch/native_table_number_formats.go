package xlsxpatch

import (
	"fmt"
	"strconv"
	"strings"
)

// Project only explicit table/column DXF number formats, with source-qualified
// General-style cells. This is display evidence, never mutation authority.
func inspectNativeTableNumberFormats(pkg *nativeWorkbookPackage, root *previewXML, table *NativeTablePreviewV1) error {
	if table.SheetPart == "" {
		return nil
	}
	top, left, bottom, right, err := parseDimensionReference(table.Ref)
	if err != nil {
		return err
	}
	if bottom-top+1 < table.HeaderRows+table.TotalRows {
		return fmt.Errorf("table row regions overlap")
	}
	columns := root.child("tableColumns")
	if columns == nil {
		return nil
	}
	if right-left+1 > 128 {
		table.Warnings = append(table.Warnings, "Table number-format preview unavailable: more than 128 columns.")
		return nil
	}
	if len(columns.children) != right-left+1 {
		return fmt.Errorf("table columns do not match table reference")
	}
	worksheet, err := parsePreviewXML(pkg.files[table.SheetPart])
	if err != nil {
		return err
	}
	for _, child := range worksheet.children {
		if child.name.Local == "conditionalFormatting" {
			table.Warnings = append(table.Warnings, "Table number-format preview unavailable while conditional formatting is unevaluated.")
			return nil
		}
	}
	location, err := locateWorkbookPartBytes(pkg.index, func(name string) ([]byte, bool) { v, ok := pkg.files[name]; return v, ok })
	if err != nil {
		return err
	}
	ns, relNS, expected, opposing := spreadsheetMLTransitional, officeRelNamespaceTransitional, relTypeStylesTransitional, relTypeStylesStrict
	if location.strict {
		ns, relNS, expected, opposing = spreadsheetMLStrict, officeRelNamespaceStrict, relTypeStylesStrict, relTypeStylesTransitional
	}
	extractor := &nativeWorkbookExtractor{pkg: pkg, workbook: location, namespace: ns, relNamespace: relNS, modeled: map[string]bool{}, unsupportedKeys: map[string]bool{}, claimedXML: map[string]bool{}}
	stylesPart, err := extractor.relatedCorePart(expected, opposing, stylesPartContentType, "styles", false)
	if err != nil {
		return err
	}
	if stylesPart == "" {
		return nil
	}
	styles, err := parsePreviewXML(pkg.files[stylesPart])
	if err != nil {
		return err
	}
	dxfs := styles.child("dxfs")
	registry, err := newStyleRegistry(pkg.files[stylesPart])
	if err != nil {
		return err
	}
	eligible := []int{}
	for id, xf := range registry.cellXfs {
		base := effectiveCellStyleXF(registry.styleXfs[xf.xfID])
		if len(eligible) < 4096 && effectiveStyleComponent(xf.numFmtID, base.numFmtID, xf.applyNumberFormat) == 0 && (xf.applyNumberFormat == nil || !*xf.applyNumberFormat) {
			eligible = append(eligible, id)
		}
	}
	namedStyleWarning := false
	for offset, column := range columns.children {
		if column.name.Space != root.name.Space || column.name.Local != "tableColumn" {
			return fmt.Errorf("unqualified table column namespace")
		}
		for _, region := range []struct {
			attribute, cellStyle string
			first, last          int
			enabled              bool
		}{{"headerRowDxfId", "headerRowCellStyle", top, top, table.HeaderRows == 1}, {"dataDxfId", "dataCellStyle", top + table.HeaderRows, bottom - table.TotalRows, true}, {"totalsRowDxfId", "totalsRowCellStyle", bottom, bottom, table.TotalRows == 1}} {
			index := column.attr(region.attribute)
			if index == "" {
				index = root.attr(region.attribute)
			}
			if index == "" || !region.enabled || region.first > region.last {
				continue
			}
			if root.attr(region.cellStyle) != "" || column.attr(region.cellStyle) != "" {
				if !namedStyleWarning {
					table.Warnings = append(table.Warnings, "Named table cell styles are not resolved for differential number formatting.")
					namedStyleWarning = true
				}
				continue
			}
			id, e := strconv.Atoi(index)
			if e != nil || id < 0 || id > 65535 || strconv.Itoa(id) != index || dxfs == nil || id >= len(dxfs.children) {
				return fmt.Errorf("table differential format index is invalid")
			}
			dxf := dxfs.children[id]
			if dxf.name.Space != root.name.Space || dxf.name.Local != "dxf" {
				return fmt.Errorf("unqualified differential format namespace")
			}
			var number *previewXML
			for _, child := range dxf.children {
				if child.name.Local == "numFmt" {
					if child.name.Space != root.name.Space || number != nil {
						return fmt.Errorf("ambiguous differential number format")
					}
					number = child
				}
			}
			if number == nil {
				continue
			}
			if len(number.children) != 0 || strings.TrimSpace(number.text) != "" {
				return fmt.Errorf("unsupported differential number-format structure")
			}
			for _, a := range number.attrs {
				if a.Name.Space != "" || (a.Name.Local != "numFmtId" && a.Name.Local != "formatCode") {
					return fmt.Errorf("unsupported differential number-format attribute")
				}
			}
			format := number.attr("formatCode")
			formatID, formatErr := strconv.Atoi(number.attr("numFmtId"))
			if formatErr != nil || formatID < 0 || formatID > 65535 || strconv.Itoa(formatID) != number.attr("numFmtId") {
				return fmt.Errorf("invalid differential number-format identifier")
			}
			if format == "" || len(format) > 4096 {
				return fmt.Errorf("table differential format code unavailable or unbounded")
			}
			table.NumberFormats = append(table.NumberFormats, NativeTableNumberFormatV1{Ref: cellReference(region.first, left+offset) + ":" + cellReference(region.last, left+offset), DxfID: id, NumberFormat: format, StyleIDs: eligible})
		}
	}
	if len(table.NumberFormats) > 0 {
		table.Warnings = append(table.Warnings, "Explicit table differential number formats are previewed for default-format cells; formula results remain saved caches, not recalculated values.")
	}
	return nil
}
