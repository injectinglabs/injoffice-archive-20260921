package pptxpatch

import (
	"encoding/xml"
	"fmt"
)

const (
	nativeDefaultTextInsetHorizontalEMU int64 = 91440
	nativeDefaultTextInsetVerticalEMU   int64 = 45720
	nativeMaxTextInsetEMU               int64 = 2147483647
)

type nativeTextLayoutUnsupportedError struct {
	message string
}

type nativeTextContentUnsupportedError struct {
	message string
}

func (err nativeTextLayoutUnsupportedError) Error() string {
	return "pptxpatch: native extract: unsupported text-body layout: " + err.message
}

func unsupportedNativeTextLayout(format string, values ...any) error {
	return nativeTextLayoutUnsupportedError{message: fmt.Sprintf(format, values...)}
}

func isNativeTextLayoutUnsupported(err error) bool {
	_, ok := err.(nativeTextLayoutUnsupportedError)
	return ok
}

func (err nativeTextContentUnsupportedError) Error() string {
	return "pptxpatch: native extract: unsupported text content: " + err.message
}

func unsupportedNativeTextContent(message string) error {
	return nativeTextContentUnsupportedError{message: message}
}

func isNativeTextContentUnsupported(err error) bool {
	_, ok := err.(nativeTextContentUnsupportedError)
	return ok
}

func extractNativeTextBodyLayout(txBody *nativeXMLNode, dialect nativeExtractDialect) (*NativeTextBodyLayout, error) {
	return extractNativeTextBodyLayoutPolicy(txBody, dialect, false)
}

func extractNativeTextBodyLayoutPolicy(txBody *nativeXMLNode, dialect nativeExtractDialect, allowSourceFrame bool) (*NativeTextBodyLayout, error) {
	if txBody == nil {
		return nil, fmt.Errorf("pptxpatch: native extract: missing text body")
	}
	bodyPr, err := nativeSingleton(txBody, dialect.drawing, "bodyPr", true)
	if err != nil {
		return nil, err
	}
	if !onlyNativeXMLSpace(bodyPr.Text) {
		return nil, unsupportedNativeTextLayout("a:bodyPr contains text content")
	}
	if err := requireOnlyNativeAttrs(bodyPr,
		xml.Name{Local: "rot"}, xml.Name{Local: "spcFirstLastPara"},
		xml.Name{Local: "vertOverflow"}, xml.Name{Local: "horzOverflow"},
		xml.Name{Local: "vert"}, xml.Name{Local: "wrap"},
		xml.Name{Local: "lIns"}, xml.Name{Local: "tIns"}, xml.Name{Local: "rIns"}, xml.Name{Local: "bIns"},
		xml.Name{Local: "numCol"}, xml.Name{Local: "spcCol"}, xml.Name{Local: "rtlCol"},
		xml.Name{Local: "fromWordArt"}, xml.Name{Local: "anchor"}, xml.Name{Local: "anchorCtr"},
		xml.Name{Local: "forceAA"}, xml.Name{Local: "upright"}, xml.Name{Local: "compatLnSpc"}); err != nil {
		return nil, unsupportedNativeTextLayout("a:bodyPr contains an unsupported attribute")
	}
	allowedChildren := []xml.Name{
		{Space: dialect.drawing, Local: "prstTxWarp"},
		{Space: dialect.drawing, Local: "noAutofit"},
		{Space: dialect.drawing, Local: "normAutofit"},
		{Space: dialect.drawing, Local: "spAutoFit"},
		{Space: dialect.drawing, Local: "scene3d"},
		{Space: dialect.drawing, Local: "sp3d"},
		{Space: dialect.drawing, Local: "flatTx"},
		{Space: dialect.drawing, Local: "extLst"},
	}
	if err := requireOnlyNativeChildren(bodyPr, allowedChildren...); err != nil {
		return nil, unsupportedNativeTextLayout("a:bodyPr contains an unsupported child")
	}
	for _, name := range allowedChildren {
		if _, err := nativeSingleton(bodyPr, name.Space, name.Local, false); err != nil {
			return nil, err
		}
	}
	noAutofit, err := nativeSingleton(bodyPr, dialect.drawing, "noAutofit", false)
	if err != nil {
		return nil, err
	}
	normalAutofit, err := nativeSingleton(bodyPr, dialect.drawing, "normAutofit", false)
	if err != nil {
		return nil, err
	}
	shapeAutofit, err := nativeSingleton(bodyPr, dialect.drawing, "spAutoFit", false)
	if err != nil {
		return nil, err
	}
	autofitCount := 0
	for _, child := range []*nativeXMLNode{noAutofit, normalAutofit, shapeAutofit} {
		if child != nil {
			autofitCount++
		}
	}
	if autofitCount > 1 {
		return nil, fmt.Errorf("pptxpatch: native extract: conflicting text autofit children")
	}
	if normalAutofit != nil {
		return nil, unsupportedNativeTextLayout("a:normAutofit requires font scaling and line-spacing reduction")
	}
	autoFit := "none"
	if shapeAutofit != nil {
		if !allowSourceFrame {
			return nil, unsupportedNativeTextLayout("a:spAutoFit requires content-dependent shape sizing")
		}
		if err := requireEmptyNativeElement(shapeAutofit); err != nil {
			return nil, unsupportedNativeTextLayout("a:spAutoFit contains unsupported markup")
		}
		autoFit = "shape-source-frame"
	}
	if noAutofit != nil {
		if err := requireEmptyNativeElement(noAutofit); err != nil {
			return nil, unsupportedNativeTextLayout("a:noAutofit contains unsupported markup")
		}
	}
	for _, name := range []string{"prstTxWarp", "scene3d", "sp3d", "flatTx", "extLst"} {
		if child, _ := nativeSingleton(bodyPr, dialect.drawing, name, false); child != nil {
			return nil, unsupportedNativeTextLayout("%s is not representable", name)
		}
	}

	left, err := optionalNativeTextInset(bodyPr, "lIns", nativeDefaultTextInsetHorizontalEMU)
	if err != nil {
		return nil, err
	}
	right, err := optionalNativeTextInset(bodyPr, "rIns", nativeDefaultTextInsetHorizontalEMU)
	if err != nil {
		return nil, err
	}
	top, err := optionalNativeTextInset(bodyPr, "tIns", nativeDefaultTextInsetVerticalEMU)
	if err != nil {
		return nil, err
	}
	bottom, err := optionalNativeTextInset(bodyPr, "bIns", nativeDefaultTextInsetVerticalEMU)
	if err != nil {
		return nil, err
	}

	wrap := NativeTextWrapSquare
	if value, ok := exactNativeAttr(bodyPr, "", "wrap"); ok {
		switch value {
		case "square":
		case "none":
			wrap = NativeTextWrapNone
		default:
			return nil, fmt.Errorf("pptxpatch: native extract: invalid text wrap")
		}
	}
	anchor := NativeTextVerticalAnchorTop
	if value, ok := exactNativeAttr(bodyPr, "", "anchor"); ok {
		switch value {
		case "t":
		case "ctr":
			anchor = NativeTextVerticalAnchorCenter
		case "b":
			anchor = NativeTextVerticalAnchorBottom
		case "just", "dist":
			return nil, unsupportedNativeTextLayout("distributed or justified vertical anchoring is not representable")
		default:
			return nil, fmt.Errorf("pptxpatch: native extract: invalid text anchor")
		}
	}
	if value, ok := exactNativeAttr(bodyPr, "", "horzOverflow"); ok && value != "overflow" {
		if value != "clip" {
			return nil, fmt.Errorf("pptxpatch: native extract: invalid horizontal text overflow")
		}
		return nil, unsupportedNativeTextLayout("clipped horizontal overflow is not representable")
	}
	if value, ok := exactNativeAttr(bodyPr, "", "vertOverflow"); ok && value != "overflow" {
		if value != "clip" && value != "ellipsis" {
			return nil, fmt.Errorf("pptxpatch: native extract: invalid vertical text overflow")
		}
		return nil, unsupportedNativeTextLayout("clipped or ellipsis vertical overflow is not representable")
	}
	var writingMode *string
	if value, ok := exactNativeAttr(bodyPr, "", "vert"); ok && value != "horz" {
		switch value {
		case "vert":
			writingMode = stringPointer("vertical-clockwise")
		case "vert270", "wordArtVert", "eaVert", "mongolianVert", "wordArtVertRtl":
			return nil, unsupportedNativeTextLayout("non-horizontal text flow is not representable")
		default:
			return nil, fmt.Errorf("pptxpatch: native extract: invalid text flow")
		}
	}
	if value, ok := exactNativeAttr(bodyPr, "", "rot"); ok {
		rotation, parseErr := parseCanonicalNativeInt(value, -2147483648, 2147483647)
		if parseErr != nil {
			return nil, fmt.Errorf("pptxpatch: native extract: invalid text rotation")
		}
		if rotation != 0 {
			return nil, unsupportedNativeTextLayout("rotated text is not representable")
		}
	}
	if value, ok := exactNativeAttr(bodyPr, "", "numCol"); ok {
		columns, parseErr := parseCanonicalNativeInt(value, 1, 2147483647)
		if parseErr != nil {
			return nil, fmt.Errorf("pptxpatch: native extract: invalid text column count")
		}
		if columns != 1 {
			return nil, unsupportedNativeTextLayout("multiple text columns are not representable")
		}
	}
	if _, ok := exactNativeAttr(bodyPr, "", "spcCol"); ok {
		return nil, unsupportedNativeTextLayout("column spacing is not representable")
	}
	for _, name := range []string{"rtlCol", "fromWordArt", "anchorCtr", "forceAA", "upright", "compatLnSpc", "spcFirstLastPara"} {
		if value, ok := exactNativeAttr(bodyPr, "", name); ok {
			enabled, boolErr := nativeBool(value)
			if boolErr != nil {
				return nil, fmt.Errorf("pptxpatch: native extract: invalid boolean for text-body %s", name)
			}
			if enabled {
				return nil, unsupportedNativeTextLayout("%s is not representable", name)
			}
		}
	}

	return &NativeTextBodyLayout{
		LeftInsetEMU: &left, RightInsetEMU: &right, TopInsetEMU: &top, BottomInsetEMU: &bottom,
		Wrap: wrap, VerticalAnchor: anchor, AutoFit: autoFit,
		HorizontalOverflow: "overflow", VerticalOverflow: "overflow",
		WritingMode: writingMode,
	}, nil
}

func nativeMarkSourceFrameAutoFit(element *NativeElement) {
	if element.TextBody == nil || element.TextBody.AutoFit != "shape-source-frame" {
		return
	}
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: "pptx.autofit-source-frame-approximate", Message: "Read-only approximate autofit preview uses the saved source frame without resizing; frame size, layout, and overflow or clipping may differ from PowerPoint."})
}

func nativeMarkVerticalTextPreview(element *NativeElement) {
	if element.TextBody == nil || element.TextBody.WritingMode == nil {
		return
	}
	if element.Compatibility.Status == NativeCompatibilityStatusEditable {
		element.Compatibility.Status = NativeCompatibilityStatusPreserveOnly
	}
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: "pptx.vertical-text-preview", Message: "clockwise vertical text is read-only; native rendering qualifies only non-bulleted ASCII Latin text"})
}

func validateNativeTextBodyBounds(layout *NativeTextBodyLayout, transform NativeTransform) error {
	if layout == nil || layout.LeftInsetEMU == nil || layout.RightInsetEMU == nil || layout.TopInsetEMU == nil || layout.BottomInsetEMU == nil || transform.Cx == nil || transform.Cy == nil {
		return fmt.Errorf("pptxpatch: native extract: incomplete text-body bounds")
	}
	if *transform.Cx-*layout.LeftInsetEMU-*layout.RightInsetEMU <= 0 {
		return unsupportedNativeTextLayout("horizontal insets leave no positive text-body width")
	}
	if *transform.Cy-*layout.TopInsetEMU-*layout.BottomInsetEMU <= 0 {
		return unsupportedNativeTextLayout("vertical insets leave no positive text-body height")
	}
	return nil
}

func optionalNativeTextInset(bodyPr *nativeXMLNode, name string, fallback int64) (int64, error) {
	value, ok := exactNativeAttr(bodyPr, "", name)
	if !ok {
		return fallback, nil
	}
	inset, err := parseCanonicalNativeInt(value, -nativeMaxSafeInteger, nativeMaxSafeInteger)
	if err != nil {
		return 0, fmt.Errorf("pptxpatch: native extract: invalid canonical text inset %s", name)
	}
	if inset < 0 || inset > nativeMaxTextInsetEMU {
		return 0, unsupportedNativeTextLayout("%s is outside the nonnegative signed 32-bit inset subset", name)
	}
	return inset, nil
}
