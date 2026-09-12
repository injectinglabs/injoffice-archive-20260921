package xlsxpatch

import (
	"strconv"
	"strings"
)

// Explicit source settings only. This is not a printer-default resolver.
type NativeSheetPageSettingsV1 struct {
	SheetID   string                   `json:"sheet_id"`
	SheetPart string                   `json:"sheet_part"`
	Status    string                   `json:"status"`
	Settings  *NativeSheetPageConfigV1 `json:"settings,omitempty"`
	Warnings  []string                 `json:"warnings"`
}
type NativeSheetPageConfigV1 struct {
	Paper       string  `json:"paper"`
	Orientation string  `json:"orientation"`
	Scale       int     `json:"scale"`
	Left        float64 `json:"left_inches"`
	Right       float64 `json:"right_inches"`
	Top         float64 `json:"top_inches"`
	Bottom      float64 `json:"bottom_inches"`
	PageOrder   string  `json:"page_order,omitempty"`
}

func previewNativePageSettings(raw []byte, part, id string) NativeSheetPageSettingsV1 {
	result := NativeSheetPageSettingsV1{SheetID: id, SheetPart: part, Status: "unavailable", Warnings: []string{"Worksheet page preview requires explicit bounded paper, orientation, scale and margins; printer defaults and unsupported print options are not inferred."}}
	root, err := parsePreviewXML(raw)
	if err != nil || root.name.Local != "worksheet" || !isSpreadsheetMLNamespace(root.name.Space) {
		return result
	}
	setup, margins := root.child("pageSetup"), root.child("pageMargins")
	if setup == nil || margins == nil {
		return result
	}
	// Do not silently ignore alternate/foreign settings or page-affecting data.
	for _, child := range root.children {
		switch child.name.Local {
		case "pageSetup", "pageMargins":
			if child.name.Space != root.name.Space {
				return result
			}
		case "rowBreaks", "colBreaks", "headerFooter", "printOptions":
			return result
		case "sheetPr":
			for _, p := range child.children {
				if p.name.Local == "pageSetUpPr" {
					return result
				}
			}
		}
	}
	exactLeaf := func(n *previewXML, names ...string) bool {
		if len(n.children) != 0 || strings.TrimSpace(n.text) != "" {
			return false
		}
		allowed := map[string]bool{}
		for _, name := range names {
			allowed[name] = true
		}
		count := 0
		for _, a := range n.attrs {
			if isPreviewNamespaceDeclaration(a) {
				continue
			}
			if a.Name.Space != "" || !allowed[a.Name.Local] {
				return false
			}
			count++
		}
		return count == len(names)
	}
	setupFields := []string{"paperSize", "orientation", "scale"}
	order := setup.attr("pageOrder")
	if order != "" {
		if order != "downThenOver" && order != "overThenDown" {
			return result
		}
		setupFields = append(setupFields, "pageOrder")
	}
	if !exactLeaf(setup, setupFields...) || !exactLeaf(margins, "left", "right", "top", "bottom", "header", "footer") {
		return result
	}
	paper := map[string]string{"1": "Letter", "9": "A4"}[setup.attr("paperSize")]
	orientation := setup.attr("orientation")
	scale, err := strconv.Atoi(setup.attr("scale"))
	if paper == "" || (orientation != "portrait" && orientation != "landscape") || err != nil || scale < 10 || scale > 400 || strconv.Itoa(scale) != setup.attr("scale") {
		return result
	}
	values := map[string]float64{}
	for _, name := range []string{"left", "right", "top", "bottom", "header", "footer"} {
		value, ok := boundedPreviewRowNumber(margins.attr(name), 20)
		if !ok {
			return result
		}
		values[name] = value
	}
	result.Status = "available"
	result.Settings = &NativeSheetPageConfigV1{paper, orientation, scale, values["left"], values["right"], values["top"], values["bottom"], order}
	result.Warnings = []string{"Read-only selected-range page geometry approximation. Native print areas, titles, charts, headers and printer behavior are not reproduced; no Excel fidelity claim."}
	return result
}
