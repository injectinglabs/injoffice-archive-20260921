package xlsxpatch

import (
	"strconv"
	"strings"
)

func qualifyNativeTableBorders(root *previewXML, table *NativeTablePreviewV1, registry *styleRegistry, styles *previewXML, accent string) {
	if registry == nil || styles == nil || len(registry.borders) == 0 {
		return
	}
	border := registry.borders[0]
	if !border.supported || border.left != nil || border.right != nil || border.top != nil || border.bottom != nil {
		return
	}
	var overrides func(*previewXML) bool
	overrides = func(node *previewXML) bool {
		for _, a := range node.attrs {
			if a.Name.Space != "" {
				continue
			}
			if strings.HasSuffix(a.Name.Local, "CellStyle") || strings.HasSuffix(a.Name.Local, "BorderDxfId") {
				return true
			}
			if strings.HasSuffix(a.Name.Local, "DxfId") {
				id, e := strconv.Atoi(a.Value)
				dxfs := styles.child("dxfs")
				if e != nil || id < 0 || dxfs == nil || id >= len(dxfs.children) {
					return true
				}
				dxf := dxfs.children[id]
				for _, child := range dxf.children {
					if child.name.Space != root.name.Space || child.name.Local != "numFmt" {
						return true
					}
				}
			}
		}
		for _, child := range node.children {
			if overrides(child) {
				return true
			}
		}
		return false
	}
	if overrides(root) {
		return
	}
	ids := []int{}
	for id, xf := range registry.cellXfs {
		base := effectiveCellStyleXF(registry.styleXfs[xf.xfID])
		if len(ids) < 4096 && effectiveStyleComponent(xf.borderID, base.borderID, xf.applyBorder) == 0 && (xf.applyBorder == nil || !*xf.applyBorder) {
			ids = append(ids, id)
		}
	}
	// Original Excel PDF operators establish 1pt thin edges and a 3pt totals
	// divider (two 1pt strokes and a 1pt gap), not a generic CSS grid border.
	table.BorderPreview = &NativeTableBorderPreviewV1{Color: tableLightenHLS(accent, 0.4), TotalsColor: accent, WidthPoints: 1, TotalsWidthPoints: 3, StyleIDs: ids}
	table.Warnings = append(table.Warnings, "Medium2 borders are previewed only between source-qualified default-border cells. Explicit, unknown or conflicting neighboring borders are not replaced.")
}
