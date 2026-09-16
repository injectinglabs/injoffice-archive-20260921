package pptxpatch

import (
	"encoding/xml"
	"fmt"
	"strconv"
)

const (
	nativeAuthoredAutoFitCode     = "pptx.autofit-authored-scale-approximate"
	nativeTextColumnsOmittedCode  = "pptx.text-columns-single-column-approximate"
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
}

func (fit *nativeAuthoredAutoFit) approximate() bool {
	return fit != nil && (fit.normAutofit || fit.columns > 1 || fit.columnSpacingEMU > 0)
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

// nativeMarkAuthoredAutoFit labels the projection. The element becomes
// preserve-only because PowerPoint recomputes these values on edit.
func nativeMarkAuthoredAutoFit(element *NativeElement, fit *nativeAuthoredAutoFit) {
	if element == nil || !fit.approximate() {
		return
	}
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
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
		element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{
			Severity: NativeDiagnosticSeverityWarning,
			Code:     nativeTextColumnsOmittedCode,
			Message: "Read-only approximate preview paints the authored numCol=" + strconv.FormatInt(fit.columns, 10) +
				" spcCol=" + strconv.FormatInt(fit.columnSpacingEMU, 10) + " EMU text body as a single column; column flow and line breaks differ from PowerPoint.",
		})
	}
}
