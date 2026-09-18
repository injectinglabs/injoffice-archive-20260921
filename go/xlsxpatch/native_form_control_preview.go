package xlsxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
	"strings"
)

// Namespaces the form-control preview reads. x14 carries the worksheet
// <controls> block's markup-compatibility wrapper; the control elements
// themselves are written in the SpreadsheetML default namespace.
const (
	nativeMarkupCompatibilityNamespace = "http://schemas.openxmlformats.org/markup-compatibility/2006"
	nativeFormControlNamespace         = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
	nativeVMLNamespace                 = "urn:schemas-microsoft-com:vml"
	nativeVMLExcelNamespace            = "urn:schemas-microsoft-com:office:excel"
)

// The only x:ClientData ObjectType and formControlPr objectType this preview
// paints. Every other control is a different bounded shape with its own
// states, and none of them is a rectangle with a border: painting any of them
// as one would be inventing a control. They stay unsupported and disclosed.
const (
	nativeFormControlCheckboxClientType = "Checkbox"
	nativeFormControlCheckboxPropType   = "CheckBox"
)

// NativeFormControlV1 is one worksheet form control, read from the three parts
// that together state it: the worksheet's <controls> anchor (placement), the
// ctrlProps part (control type and state) and the legacy VML drawing
// (ObjectType, caption text and caption alignment).
//
// Kind is "checkbox" only when all three agree that this is an unchecked or
// checked checkbox with a positive two-cell anchor and a caption. Anything
// else is "unsupported": the anchor may still be reported, but no consumer may
// paint a box for it.
type NativeFormControlV1 struct {
	SheetID   string `json:"sheet_id"`
	SheetPart string `json:"sheet_part"`
	Ordinal   int    `json:"ordinal"`
	ShapeID   string `json:"shape_id"`
	Name      string `json:"name"`
	// "checkbox" or "unsupported".
	Kind string `json:"kind"`
	// The source ObjectType, verbatim, so a refusal names what it refused.
	ObjectType    string                 `json:"object_type,omitempty"`
	Checked       bool                   `json:"checked"`
	Anchor        *NativeDrawingAnchorV1 `json:"anchor,omitempty"`
	Caption       string                 `json:"caption,omitempty"`
	CaptionSize   float64                `json:"caption_size_points,omitempty"`
	CaptionAlign  string                 `json:"caption_align,omitempty"`
	CaptionVAlign string                 `json:"caption_valign,omitempty"`
	ControlPart   string                 `json:"control_part,omitempty"`
	LegacyPart    string                 `json:"legacy_part,omitempty"`
	Warnings      []string               `json:"warnings"`
}

const nativeFormControlCheckboxWarning = "Source anchor, control type and legacy caption only. The box is Excel's 12 pt control glyph placed inside the source anchor rectangle; three-dimensional shading, focus state, the linked cell and macro behaviour are not reproduced."

func nativeFormControlUnsupported(objectType string) string {
	if objectType == "" {
		return "Form control placement unavailable: the control does not state an ObjectType this tier reads. No box is painted and no position is guessed."
	}
	return fmt.Sprintf("Form control %q is not a checkbox. Only ObjectType Checkbox is painted; every other control is a different bounded shape with its own states and is not drawn as a rectangle. No box is painted.", objectType)
}

// previewNativeFormControls reads every worksheet's <controls> block within the
// shared bounded preview limits. An unreadable, ambiguous or unexpected part
// yields no control at all rather than a guessed one.
func previewNativeFormControls(pkg *nativeWorkbookPackage, sheets []NativeWorkbookSheetV2) ([]NativeFormControlV1, error) {
	result := []NativeFormControlV1{}
	for _, sheet := range sheets {
		root, err := parsePreviewXML(pkg.files[sheet.PartName])
		if err != nil {
			return nil, err
		}
		if root.name.Local != "worksheet" || !isSpreadsheetMLNamespace(root.name.Space) {
			continue
		}
		controls := nativeFormControlElements(root)
		if len(controls) == 0 {
			continue
		}
		legacy := ""
		if node := root.child("legacyDrawing"); node != nil {
			r := officeRelNamespaceTransitional
			if root.name.Space == spreadsheetMLStrict {
				r = officeRelNamespaceStrict
			}
			if previewAttrs(node, xml.Name{Space: r, Local: "id"}) && len(node.children) == 0 {
				legacy, err = previewInternalTarget(pkg, sheet.PartName, previewQualifiedAttr(node, r, "id"), "vmlDrawing")
				if err != nil {
					return nil, err
				}
			}
		}
		shapes := map[string]*previewXML{}
		if legacy != "" {
			if shapes, err = nativeFormControlVMLShapes(pkg.files[legacy]); err != nil {
				return nil, err
			}
		}
		for i, control := range controls {
			item, err := nativeFormControlFrom(pkg, sheet, root, control, i+1, legacy, shapes)
			if err != nil {
				return nil, err
			}
			result = append(result, item)
			if len(result) > 256 {
				return nil, fmt.Errorf("form control preview exceeds 256 source controls")
			}
		}
	}
	return result, nil
}

// nativeFormControlElements collects the worksheet's control elements in source
// order. Excel writes them as
// mc:AlternateContent/mc:Choice[@Requires=x14]/controls, with each control
// wrapped in its own mc:AlternateContent/mc:Choice. An mc:Fallback is never
// read: its whole purpose is to state what a reader without x14 should do
// instead, and this reader has the x14 content.
func nativeFormControlElements(root *previewXML) []*previewXML {
	controls := []*previewXML{}
	for _, child := range root.children {
		if child.name != (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "AlternateContent"}) {
			continue
		}
		for _, choice := range child.children {
			if choice.name != (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "Choice"}) || choice.attr("Requires") != "x14" {
				continue
			}
			for _, block := range choice.children {
				if block.name != (xml.Name{Space: root.name.Space, Local: "controls"}) {
					continue
				}
				for _, entry := range block.children {
					if entry.name == (xml.Name{Space: root.name.Space, Local: "control"}) {
						controls = append(controls, entry)
						continue
					}
					if entry.name != (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "AlternateContent"}) {
						continue
					}
					for _, inner := range entry.children {
						if inner.name != (xml.Name{Space: nativeMarkupCompatibilityNamespace, Local: "Choice"}) || inner.attr("Requires") != "x14" {
							continue
						}
						for _, candidate := range inner.children {
							if candidate.name == (xml.Name{Space: root.name.Space, Local: "control"}) {
								controls = append(controls, candidate)
							}
						}
					}
				}
			}
		}
	}
	return controls
}

func nativeFormControlFrom(pkg *nativeWorkbookPackage, sheet NativeWorkbookSheetV2, root, control *previewXML, ordinal int, legacy string, shapes map[string]*previewXML) (NativeFormControlV1, error) {
	r := officeRelNamespaceTransitional
	xdr := "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
	if root.name.Space == spreadsheetMLStrict {
		r = officeRelNamespaceStrict
		xdr = "http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing"
	}
	item := NativeFormControlV1{
		SheetID: sheet.ID, SheetPart: sheet.PartName, Ordinal: ordinal,
		ShapeID: control.attr("shapeId"), Name: control.attr("name"),
		Kind: "unsupported", LegacyPart: legacy,
	}
	part, err := previewInternalTarget(pkg, sheet.PartName, previewQualifiedAttr(control, r, "id"), "ctrlProp")
	if err != nil {
		return item, err
	}
	item.ControlPart = part
	objectType := ""
	if part != "" {
		props, err := parsePreviewXML(pkg.files[part])
		if err != nil {
			return item, err
		}
		if props.name == (xml.Name{Space: nativeFormControlNamespace, Local: "formControlPr"}) && len(props.children) == 0 {
			objectType = props.attr("objectType")
			// CT_FormControlPr states the box's state in `checked`
			// (ST_CheckedState: Unchecked, Checked, Mixed). An absent
			// attribute is Unchecked, and a Mixed box is neither state
			// this tier can paint.
			switch props.attr("checked") {
			case "", "Unchecked":
			case "Checked":
				item.Checked = true
			default:
				objectType = ""
			}
		}
	}
	item.ObjectType = objectType
	item.Anchor = nativeFormControlAnchor(control, root.name.Space, xdr)

	shape := shapes["_x0000_s"+item.ShapeID]
	clientType := ""
	if shape != nil {
		if client := previewChildNS(shape, nativeVMLExcelNamespace, "ClientData"); client != nil {
			clientType = client.attr("ObjectType")
			item.CaptionAlign, item.CaptionVAlign = nativeFormControlAlignment(client)
		}
		item.Caption, item.CaptionSize = nativeFormControlCaption(shape)
	}
	// The three parts must agree, and the caption and the box must both be
	// present before either is painted: a caption with no box is the failure
	// this preview exists to avoid.
	if objectType != nativeFormControlCheckboxPropType || clientType != nativeFormControlCheckboxClientType {
		reported := objectType
		if reported == "" {
			reported = clientType
		}
		item.Warnings = []string{nativeFormControlUnsupported(reported)}
		return item, nil
	}
	if item.Anchor == nil || item.Anchor.To == nil || item.Caption == "" || item.CaptionSize <= 0 {
		item.Warnings = []string{"Form control placement unavailable: the checkbox does not state a complete two-cell anchor and a legacy caption. No box is painted and no position is guessed."}
		return item, nil
	}
	item.Kind = "checkbox"
	item.Warnings = []string{nativeFormControlCheckboxWarning}
	return item, nil
}

// nativeFormControlAnchor reads controlPr/anchor, which is CT_ObjectAnchor: the
// same xdr from/to markers the drawing preview reads, so the same bounded
// marker parser qualifies it.
func nativeFormControlAnchor(control *previewXML, ns, xdr string) *NativeDrawingAnchorV1 {
	props := previewChildNS(control, ns, "controlPr")
	if props == nil {
		return nil
	}
	anchor := previewChildNS(props, ns, "anchor")
	if anchor == nil || len(anchor.children) != 2 {
		return nil
	}
	// CT_ObjectAnchor writes from/to in the worksheet's own namespace and their
	// col/row/colOff/rowOff markers in the spreadsheet-drawing namespace.
	from, ok := previewDrawingMarker(previewChildNS(anchor, ns, "from"), xdr)
	if !ok {
		return nil
	}
	to, ok := previewDrawingMarker(previewChildNS(anchor, ns, "to"), xdr)
	if !ok {
		return nil
	}
	return &NativeDrawingAnchorV1{Kind: "twoCellAnchor", From: from, To: &to}
}

// nativeFormControlAlignment reads x:TextHAlign and x:TextVAlign. Excel omits
// each when the control uses its default, which for a checkbox caption is left
// and top.
func nativeFormControlAlignment(client *previewXML) (string, string) {
	horizontal, vertical := "left", "top"
	for _, child := range client.children {
		if child.name.Space != nativeVMLExcelNamespace {
			continue
		}
		value := strings.ToLower(strings.TrimSpace(child.text))
		switch child.name.Local {
		case "TextHAlign":
			if value == "left" || value == "center" || value == "right" {
				horizontal = value
			}
		case "TextVAlign":
			if value == "top" || value == "center" || value == "bottom" {
				vertical = value
			}
		}
	}
	return horizontal, vertical
}

// nativeFormControlCaption reads the caption out of the shape's v:textbox. The
// VML <font size> attribute is in twentieths of a point, as Excel writes it
// (size="160" is 8 pt).
func nativeFormControlCaption(shape *previewXML) (string, float64) {
	box := previewChildNS(shape, nativeVMLNamespace, "textbox")
	if box == nil {
		return "", 0
	}
	text, size := "", 0.0
	var walk func(node *previewXML)
	walk = func(node *previewXML) {
		if node.name.Space == "" && node.name.Local == "font" {
			if v := node.attr("size"); v != "" {
				if n, err := strconv.Atoi(v); err == nil && n > 0 && n <= 8000 {
					size = float64(n) / 20
				}
			}
		}
		text += node.text
		for _, child := range node.children {
			walk(child)
		}
	}
	walk(box)
	text = strings.TrimSpace(text)
	if len(text) > 512 {
		return "", 0
	}
	return text, size
}

// nativeFormControlVMLShapes indexes the legacy drawing's v:shape elements by
// their VML id. A duplicate id makes the whole drawing ambiguous and yields no
// shapes at all.
func nativeFormControlVMLShapes(raw []byte) (map[string]*previewXML, error) {
	if len(raw) == 0 {
		return map[string]*previewXML{}, nil
	}
	root, err := parsePreviewXML(raw)
	if err != nil {
		// A legacy drawing this tier cannot read is not a package defect: it
		// leaves every control on the sheet without a caption, which the
		// control itself then reports as unsupported.
		return map[string]*previewXML{}, nil
	}
	shapes := map[string]*previewXML{}
	for _, child := range root.children {
		if child.name != (xml.Name{Space: nativeVMLNamespace, Local: "shape"}) {
			continue
		}
		id := child.attr("id")
		if id == "" {
			continue
		}
		if _, seen := shapes[id]; seen {
			return map[string]*previewXML{}, nil
		}
		shapes[id] = child
	}
	return shapes, nil
}

// nativeFormControlPrintArea is the smallest rectangle covering every painted
// checkbox on the sheet, or nil when none is painted. Excel's used range
// includes its anchored objects, so a worksheet whose only content is form
// controls still prints them; the dimension-derived fallback covers cells
// alone and would print an empty page here.
func nativeFormControlPrintArea(controls []NativeFormControlV1, sheetPart string) *NativePrintAreaRectV1 {
	var result *NativePrintAreaRectV1
	for _, control := range controls {
		if control.SheetPart != sheetPart || control.Kind != "checkbox" || control.Anchor == nil || control.Anchor.To == nil {
			continue
		}
		endRow, endColumn := control.Anchor.To.Row, control.Anchor.To.Column
		if control.Anchor.To.RowOffset == 0 && endRow > control.Anchor.From.Row {
			endRow--
		}
		if control.Anchor.To.ColumnOffset == 0 && endColumn > control.Anchor.From.Column {
			endColumn--
		}
		if result == nil {
			result = &NativePrintAreaRectV1{Row: 0, Column: 0, EndRow: endRow, EndColumn: endColumn}
			continue
		}
		if endRow > result.EndRow {
			result.EndRow = endRow
		}
		if endColumn > result.EndColumn {
			result.EndColumn = endColumn
		}
	}
	return result
}
