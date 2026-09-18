package xlsxpatch

import (
	"bytes"
	"strings"
	"testing"
)

// formControlFixture mirrors the bytes Excel 16.112.4 writes for a worksheet
// form control: an mc:AlternateContent/<controls> block anchored in EMU, a
// ctrlProps part stating the control type, and a legacy VML drawing carrying
// the ObjectType, the caption and the caption alignment. The anchors are the
// ones in the hard-v2 corpus file checkbox-form-control-align.xlsx.
func formControlFixture(t *testing.T) map[string]string {
	t.Helper()
	parts := nativeWorkbookFixture(false)
	parts["Charts/chart1.xml"] = previewChartFixture()
	ss, r := spreadsheetMLTransitional, officeRelNamespaceTransitional
	xdr := "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
	parts["[Content_Types].xml"] = strings.Replace(parts["[Content_Types].xml"], `</Types>`,
		`<Override PartName="/Sheets/ctrl1.xml" ContentType="application/vnd.ms-excel.controlproperties+xml"/>`+
			`<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/></Types>`, 1)
	controls := `<mc:AlternateContent xmlns:mc="` + nativeMarkupCompatibilityNamespace + `"><mc:Choice Requires="x14">` +
		`<controls xmlns="` + ss + `"><mc:AlternateContent xmlns:mc="` + nativeMarkupCompatibilityNamespace + `"><mc:Choice Requires="x14">` +
		`<control shapeId="1025" r:id="rCtrl1" name="Check Box 1" xmlns:r="` + r + `">` +
		`<controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0"><anchor moveWithCells="1">` +
		`<from xmlns:xdr="` + xdr + `"><xdr:col>0</xdr:col><xdr:colOff>123825</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>57150</xdr:rowOff></from>` +
		`<to xmlns:xdr="` + xdr + `"><xdr:col>2</xdr:col><xdr:colOff>428625</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>85725</xdr:rowOff></to>` +
		`</anchor></controlPr></control></mc:Choice></mc:AlternateContent></controls></mc:Choice></mc:AlternateContent>`
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `</worksheet>`,
		`<legacyDrawing xmlns="`+ss+`" xmlns:r="`+r+`" r:id="rVml"/>`+controls+`</worksheet>`, 1)
	parts["Sheets/_rels/s1.xml.rels"] = strings.Replace(parts["Sheets/_rels/s1.xml.rels"], `</Relationships>`,
		`<Relationship Id="rCtrl1" Type="`+r+`/ctrlProp" Target="ctrl1.xml"/>`+
			`<Relationship Id="rVml" Type="`+r+`/vmlDrawing" Target="legacy.vml"/></Relationships>`, 1)
	parts["Sheets/ctrl1.xml"] = `<formControlPr xmlns="` + nativeFormControlNamespace + `" objectType="CheckBox" lockText="1" noThreeD="1"/>`
	parts["Sheets/legacy.vml"] = `<xml xmlns:v="` + nativeVMLNamespace + `" xmlns:x="` + nativeVMLExcelNamespace + `">` +
		`<v:shape id="_x0000_s1025" style='position:absolute;margin-left:9.75pt;margin-top:4.5pt;width:120pt;height:17.25pt'>` +
		`<v:textbox><div style='text-align:right'><font face="DejaVu Sans" size="160" color="#000000">All effects</font></div></v:textbox>` +
		`<x:ClientData ObjectType="Checkbox"><x:SizeWithCells/><x:Anchor>0, 13, 0, 6, 2, 45, 1, 9</x:Anchor>` +
		`<x:AutoFill>False</x:AutoFill><x:TextHAlign>Right</x:TextHAlign><x:TextVAlign>Bottom</x:TextVAlign><x:NoThreeD/>` +
		`</x:ClientData></v:shape></xml>`
	return parts
}

func formControlsOf(t *testing.T, parts map[string]string) []NativeFormControlV1 {
	t.Helper()
	source := buildZip(t, parts)
	before := bytes.Clone(source)
	objects, err := InspectNativeWorkbookObjectsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	if !bytes.Equal(source, before) {
		t.Fatal("source modified")
	}
	return objects.FormControls
}

func TestNativeFormControlPreviewReadsCheckboxAnchorTypeAndCaption(t *testing.T) {
	controls := formControlsOf(t, formControlFixture(t))
	if len(controls) != 1 {
		t.Fatalf("form controls = %d, want 1: %+v", len(controls), controls)
	}
	control := controls[0]
	if control.Kind != "checkbox" || control.ObjectType != "CheckBox" || control.Checked {
		t.Fatalf("control type/state not read: %+v", control)
	}
	if control.SheetPart != "Sheets/s1.xml" || control.ShapeID != "1025" || control.Name != "Check Box 1" {
		t.Fatalf("control identity not read: %+v", control)
	}
	if control.ControlPart != "Sheets/ctrl1.xml" || control.LegacyPart != "Sheets/legacy.vml" {
		t.Fatalf("control parts not resolved: %+v", control)
	}
	if control.Anchor == nil || control.Anchor.To == nil {
		t.Fatalf("anchor not read: %+v", control)
	}
	if control.Anchor.From.Column != 0 || control.Anchor.From.ColumnOffset != 123825 || control.Anchor.From.Row != 0 || control.Anchor.From.RowOffset != 57150 {
		t.Fatalf("anchor from wrong: %+v", control.Anchor.From)
	}
	if control.Anchor.To.Column != 2 || control.Anchor.To.ColumnOffset != 428625 || control.Anchor.To.Row != 1 || control.Anchor.To.RowOffset != 85725 {
		t.Fatalf("anchor to wrong: %+v", *control.Anchor.To)
	}
	if control.Caption != "All effects" || control.CaptionSize != 8 {
		t.Fatalf("caption not read: %q %v", control.Caption, control.CaptionSize)
	}
	if control.CaptionAlign != "right" || control.CaptionVAlign != "bottom" {
		t.Fatalf("caption alignment not read: %+v", control)
	}
}

// ST_CheckedState (Unchecked, Checked, Mixed). Only an explicit Checked draws a
// mark; Mixed is neither state this tier paints.
func TestNativeFormControlPreviewReadsCheckedState(t *testing.T) {
	for _, testCase := range []struct {
		attribute string
		kind      string
		checked   bool
	}{
		{"", "checkbox", false},
		{` checked="Unchecked"`, "checkbox", false},
		{` checked="Checked"`, "checkbox", true},
		{` checked="Mixed"`, "unsupported", false},
	} {
		t.Run("checked"+testCase.attribute, func(t *testing.T) {
			parts := formControlFixture(t)
			parts["Sheets/ctrl1.xml"] = strings.Replace(parts["Sheets/ctrl1.xml"], ` lockText="1"`, testCase.attribute+` lockText="1"`, 1)
			controls := formControlsOf(t, parts)
			if len(controls) != 1 || controls[0].Kind != testCase.kind || controls[0].Checked != testCase.checked {
				t.Fatalf("state %q projected as %+v", testCase.attribute, controls)
			}
		})
	}
}

// Every other form control is a different bounded shape with its own states.
// None of them is a rectangle with a border, so none is painted as one.
func TestNativeFormControlPreviewRefusesEveryNonCheckboxObjectType(t *testing.T) {
	for _, objectType := range []string{"Button", "Drop", "Radio", "Spin", "Scroll", "List", "GBox", "Label", "EditBox"} {
		t.Run(objectType, func(t *testing.T) {
			parts := formControlFixture(t)
			parts["Sheets/ctrl1.xml"] = strings.Replace(parts["Sheets/ctrl1.xml"], `objectType="CheckBox"`, `objectType="`+objectType+`"`, 1)
			parts["Sheets/legacy.vml"] = strings.Replace(parts["Sheets/legacy.vml"], `ObjectType="Checkbox"`, `ObjectType="`+objectType+`"`, 1)
			controls := formControlsOf(t, parts)
			if len(controls) != 1 {
				t.Fatalf("lost unsupported control owner: %+v", controls)
			}
			control := controls[0]
			if control.Kind != "unsupported" {
				t.Fatalf("%s painted as a checkbox: %+v", objectType, control)
			}
			if len(control.Warnings) == 0 || !strings.Contains(control.Warnings[0], objectType) {
				t.Fatalf("%s refusal does not name the control: %+v", objectType, control.Warnings)
			}
		})
	}
}

// The three parts must agree. A checkbox ctrlProps under a non-checkbox VML
// shape, or the reverse, is not a checkbox at a known position.
func TestNativeFormControlPreviewRefusesDisagreeingOrIncompleteSources(t *testing.T) {
	for _, testCase := range []struct{ name, part, from, to string }{
		{"vml-says-button", "Sheets/legacy.vml", `ObjectType="Checkbox"`, `ObjectType="Button"`},
		{"props-say-button", "Sheets/ctrl1.xml", `objectType="CheckBox"`, `objectType="Button"`},
		{"no-caption", "Sheets/legacy.vml", `<v:textbox><div style='text-align:right'><font face="DejaVu Sans" size="160" color="#000000">All effects</font></div></v:textbox>`, ``},
		{"no-caption-size", "Sheets/legacy.vml", ` size="160"`, ``},
		{"no-anchor", "Sheets/s1.xml", `<anchor moveWithCells="1">`, `<anchor moveWithCells="1"><xtra/>`},
		{"unparsable-anchor-offset", "Sheets/s1.xml", `<xdr:colOff>123825</xdr:colOff>`, `<xdr:colOff>0x10</xdr:colOff>`},
		{"foreign-anchor-marker", "Sheets/s1.xml", `<xdr:col>0</xdr:col>`, `<xdr:col xmlns:xdr="urn:foreign">0</xdr:col>`},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			parts := formControlFixture(t)
			replaced := strings.Replace(parts[testCase.part], testCase.from, testCase.to, 1)
			if replaced == parts[testCase.part] {
				t.Fatalf("fixture no longer contains %q", testCase.from)
			}
			parts[testCase.part] = replaced
			controls := formControlsOf(t, parts)
			if len(controls) != 1 {
				t.Fatalf("lost unsupported control owner: %+v", controls)
			}
			if controls[0].Kind != "unsupported" || len(controls[0].Warnings) == 0 {
				t.Fatalf("incomplete source painted: %+v", controls[0])
			}
		})
	}
}

// Excel's printed used range includes its anchored objects, so a worksheet
// whose dimension states only A1 still prints the controls anchored past it.
func TestNativeFormControlPrintAreaCoversAnchoredCheckboxes(t *testing.T) {
	parts := formControlFixture(t)
	// The corpus file's dimension states A1 and its sheetData is empty: every
	// printable mark on the page is a form control.
	parts["Sheets/s1.xml"] = strings.Replace(parts["Sheets/s1.xml"], `<dimension ref="A1:K2"/>`, `<dimension ref="A1"/>`, 1)
	source := buildZip(t, parts)
	objects, err := InspectNativeWorkbookObjectsV1(source)
	if err != nil {
		t.Fatal(err)
	}
	var set *NativeSheetPrintAreaSetV1
	for i := range objects.PrintAreaSets {
		if objects.PrintAreaSets[i].SheetPart == "Sheets/s1.xml" {
			set = &objects.PrintAreaSets[i]
		}
	}
	if set == nil || set.Status != "available" || len(set.Areas) != 1 {
		t.Fatalf("no dimension-derived print area: %+v", objects.PrintAreaSets)
	}
	// The anchor ends inside row 1 and column 2, so the printed range must
	// reach them: a range that stopped at the dimension would print a page
	// with no control on it.
	if set.Areas[0].Row != 0 || set.Areas[0].Column != 0 || set.Areas[0].EndRow != 1 || set.Areas[0].EndColumn != 2 {
		t.Fatalf("print area does not cover the anchored control: %+v", set.Areas[0])
	}
	if !strings.Contains(set.Warnings[0], "form controls") {
		t.Fatalf("extended print area is not disclosed: %+v", set.Warnings)
	}
}

// A non-checkbox control never extends the printed range: nothing is painted
// for it, so printing extra empty rows and columns would be inventing content.
func TestNativeFormControlPrintAreaIgnoresUnpaintedControls(t *testing.T) {
	parts := formControlFixture(t)
	parts["Sheets/ctrl1.xml"] = strings.Replace(parts["Sheets/ctrl1.xml"], `objectType="CheckBox"`, `objectType="Button"`, 1)
	objects, err := InspectNativeWorkbookObjectsV1(buildZip(t, parts))
	if err != nil {
		t.Fatal(err)
	}
	if area := nativeFormControlPrintArea(objects.FormControls, "Sheets/s1.xml"); area != nil {
		t.Fatalf("unpainted control extended the printed range: %+v", area)
	}
}

// The controls block is the only foreign worksheet markup this tier clears as
// non-dimensional, and only in the exact shape Excel writes it.
func TestNativeFormControlMarkupIsNonDimensionalOnlyInItsOwnShape(t *testing.T) {
	parts := formControlFixture(t)
	if codes := previewNativeDimensionNeutrality([]byte(parts["Sheets/s1.xml"]), "Sheets/s1.xml").Codes; !containsCode(codes, "FOREIGN_WORKSHEET_MARKUP") {
		t.Fatalf("controls markup kept the dimension refusal: %+v", codes)
	}
	for _, testCase := range []struct{ name, from, to string }{
		{"other-foreign-child", `<legacyDrawing`, `<alien:thing xmlns:alien="urn:alien"/><legacyDrawing`},
		{"alternate-content-holds-something-else", `<controls xmlns="` + spreadsheetMLTransitional + `">`, `<sheetData xmlns="` + spreadsheetMLTransitional + `"/><controls xmlns="` + spreadsheetMLTransitional + `">`},
		{"fallback-with-content", `</mc:Choice></mc:AlternateContent></worksheet>`, `</mc:Choice><mc:Fallback xmlns:mc="` + nativeMarkupCompatibilityNamespace + `"><controls xmlns="` + spreadsheetMLTransitional + `"/></mc:Fallback></mc:AlternateContent></worksheet>`},
	} {
		t.Run(testCase.name, func(t *testing.T) {
			sheet := strings.Replace(parts["Sheets/s1.xml"], testCase.from, testCase.to, 1)
			if sheet == parts["Sheets/s1.xml"] {
				t.Fatalf("fixture no longer contains %q", testCase.from)
			}
			if codes := previewNativeDimensionNeutrality([]byte(sheet), "Sheets/s1.xml").Codes; containsCode(codes, "FOREIGN_WORKSHEET_MARKUP") {
				t.Fatalf("unrelated foreign markup cleared the dimension refusal: %+v", codes)
			}
		})
	}
}

func containsCode(codes []string, want string) bool {
	for _, code := range codes {
		if code == want {
			return true
		}
	}
	return false
}
