package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

const (
	nativeAuthoredAutoFitCode     = "pptx.autofit-authored-scale-approximate"
	nativeTextColumnsCode         = "pptx.text-columns-approximate"
	nativeTextWarpFlattenedCode   = "pptx.text-warp-flattened-approximate"
	nativeAuthoredAutoFitFullSize = int64(100000)
	nativeMaxTextColumns          = int64(16)
)

// nativeAuthoredAutoFit records the read-only frame-layout policy values that
// PowerPoint authored into a:bodyPr. fontScale/lnSpcReduction are the values
// saved by the last autofit pass; both default to a 100% / 0% no-op when the
// attributes are absent. Column metadata is disclosed, not laid out. The
// values are only honored behind AllowSourceFrameAutoFitPreview; strict
// extraction keeps refusing normAutofit and multiple columns.
type nativeAuthoredAutoFit struct {
	normAutofit      bool
	fontScale        int64
	lnSpcReduction   int64
	columns          int64
	columnSpacingEMU int64
	// warpFlattened records an a:prstTxWarp the approximate tier paints as
	// unwarped text in the saved frame instead of refusing the text body.
	warpFlattened bool
}

func (fit *nativeAuthoredAutoFit) approximate() bool {
	return fit != nil && (fit.normAutofit || fit.columns > 1 || fit.columnSpacingEMU > 0 || fit.warpFlattened)
}

// parseNativeAuthoredNormAutofit validates a:normAutofit as a bounded exact
// element. Only canonical integer percentages are accepted; the Strict
// percent-string form and every other attribute or child remain refusals.
func parseNativeAuthoredNormAutofit(node *nativeXMLNode, fit *nativeAuthoredAutoFit) error {
	if node == nil || fit == nil {
		return fmt.Errorf("pptxpatch: native extract: missing normAutofit")
	}
	if requireOnlyNativeAttrs(node, xml.Name{Local: "fontScale"}, xml.Name{Local: "lnSpcReduction"}) != nil || requireOnlyNativeChildren(node) != nil || !onlyNativeXMLSpace(node.Text) {
		return unsupportedNativeTextLayout("a:normAutofit contains unsupported markup")
	}
	if duplicateNativeAttrs(node.Attrs) {
		return unsupportedNativeTextLayout("a:normAutofit repeats an attribute")
	}
	fit.normAutofit = true
	fit.fontScale = nativeAuthoredAutoFitFullSize
	fit.lnSpcReduction = 0
	if value, ok := exactNativeAttr(node, "", "fontScale"); ok {
		scale, err := parseCanonicalNativeInt(value, 1000, nativeAuthoredAutoFitFullSize)
		if err != nil {
			return unsupportedNativeTextLayout("a:normAutofit fontScale is not a canonical 1%%-100%% integer percentage")
		}
		fit.fontScale = scale
	}
	if value, ok := exactNativeAttr(node, "", "lnSpcReduction"); ok {
		reduction, err := parseCanonicalNativeInt(value, 0, nativeAuthoredAutoFitFullSize-1)
		if err != nil {
			return unsupportedNativeTextLayout("a:normAutofit lnSpcReduction is not a canonical 0%%-99%% integer percentage")
		}
		fit.lnSpcReduction = reduction
	}
	return nil
}

// parseNativeAuthoredTextColumns validates numCol/spcCol without laying
// columns out. The renderer paints one column; the omission is disclosed.
func parseNativeAuthoredTextColumns(bodyPr *nativeXMLNode, fit *nativeAuthoredAutoFit) error {
	if bodyPr == nil || fit == nil {
		return fmt.Errorf("pptxpatch: native extract: missing text body properties")
	}
	fit.columns = 1
	if value, ok := exactNativeAttr(bodyPr, "", "numCol"); ok {
		columns, err := parseCanonicalNativeInt(value, 1, nativeMaxTextColumns)
		if err != nil {
			return unsupportedNativeTextLayout("text column count is outside the bounded 1-16 subset")
		}
		fit.columns = columns
	}
	if value, ok := exactNativeAttr(bodyPr, "", "spcCol"); ok {
		spacing, err := parseCanonicalNativeInt(value, 0, 51206400)
		if err != nil {
			return unsupportedNativeTextLayout("text column spacing is outside the bounded nonnegative subset")
		}
		fit.columnSpacingEMU = spacing
	}
	return nil
}

// nativeApplyAuthoredFontScale scales every resolved run size by the authored
// fontScale. Sizes are rounded half up to whole hundredths of a point and never
// fall below one hundredth. Nothing else about the paragraphs changes.
func nativeApplyAuthoredFontScale(paragraphs []NativeParagraph, fit *nativeAuthoredAutoFit) {
	if fit == nil || !fit.normAutofit || fit.fontScale == nativeAuthoredAutoFitFullSize {
		return
	}
	for pi := range paragraphs {
		for ri := range paragraphs[pi].Runs {
			size := paragraphs[pi].Runs[ri].FontSizeHundredthPt
			if size == nil || *size <= 0 {
				continue
			}
			scaled := (*size*fit.fontScale + nativeAuthoredAutoFitFullSize/2) / nativeAuthoredAutoFitFullSize
			if scaled < 1 {
				scaled = 1
			}
			paragraphs[pi].Runs[ri].FontSizeHundredthPt = int64Pointer(scaled)
		}
	}
}

func nativeFormatPercent(value int64) string {
	whole := value / 1000
	fraction := value % 1000
	if fraction == 0 {
		return strconv.FormatInt(whole, 10) + "%"
	}
	text := fmt.Sprintf("%d.%03d", whole, fraction)
	for len(text) > 0 && text[len(text)-1] == '0' {
		text = text[:len(text)-1]
	}
	return text + "%"
}

// nativeAuthoredColumnsFit reports whether the saved frame still leaves a
// positive width for every authored column after the insets and the authored
// gaps. When it does not, the preview keeps painting one disclosed column
// instead of emitting a column projection the renderer could not honor.
func nativeAuthoredColumnsFit(element *NativeElement, fit *nativeAuthoredAutoFit) bool {
	if element == nil || element.TextBody == nil || fit == nil || fit.columns < 2 {
		return false
	}
	if element.Transform.Cx == nil || element.TextBody.LeftInsetEMU == nil || element.TextBody.RightInsetEMU == nil {
		return false
	}
	content := *element.Transform.Cx - *element.TextBody.LeftInsetEMU - *element.TextBody.RightInsetEMU - (fit.columns-1)*fit.columnSpacingEMU
	return content > 0 && content/fit.columns > 0
}

// nativeMarkAuthoredAutoFit labels the projection. The element becomes
// preserve-only because PowerPoint recomputes these values on edit.
func nativeMarkAuthoredAutoFit(element *NativeElement, fit *nativeAuthoredAutoFit) {
	if element == nil || !fit.approximate() {
		return
	}
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	if fit.warpFlattened {
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning,
			Code:     nativeTextWarpFlattenedCode,
			Message:  "Read-only approximate preview paints the authored a:prstTxWarp text unwarped in its saved frame; the warp geometry, and any wrapping and overflow it causes, differ from PowerPoint.",
		})
	}
	if fit.normAutofit {
		reduction := ", and the authored lnSpcReduction=" + nativeFormatPercent(fit.lnSpcReduction) + " leaves line pitch at the measured natural line height"
		if fit.lnSpcReduction > 0 {
			reduction = ", and reduces the line pitch by the authored lnSpcReduction=" + nativeFormatPercent(fit.lnSpcReduction)
			if element.TextBody != nil {
				element.TextBody.LineSpacingReductionPercent1000 = int64Pointer(fit.lnSpcReduction)
			}
		}
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning,
			Code:     nativeAuthoredAutoFitCode,
			Message: "Read-only approximate autofit preview applies the authored a:normAutofit fontScale=" + nativeFormatPercent(fit.fontScale) +
				" to resolved run sizes" + reduction + ", without resizing the frame; wrapping and overflow may still differ from PowerPoint.",
		})
	}
	if fit.columns > 1 || fit.columnSpacingEMU > 0 {
		authored := "the authored numCol=" + strconv.FormatInt(fit.columns, 10) +
			" spcCol=" + strconv.FormatInt(fit.columnSpacingEMU, 10) + " EMU text body"
		message := "Read-only approximate preview paints " + authored +
			" as a single column because the saved frame leaves no positive column width; column flow and line breaks differ from PowerPoint."
		if fit.columns > 1 && nativeAuthoredColumnsFit(element, fit) {
			element.TextBody.ColumnCount = int64Pointer(fit.columns)
			element.TextBody.ColumnSpacingEMU = int64Pointer(fit.columnSpacingEMU)
			message = "Read-only approximate preview flows " + authored +
				" left to right through equal-width columns, wrapping at the column width and continuing in the next column once the frame height is reached; rtlCol is not modeled and column balancing, line breaks and overflow may differ from PowerPoint."
		}
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning,
			Code:     nativeTextColumnsCode,
			Message:  message,
		})
	}
}
