package pptxpatch

const nativeSourceAnchoredGraphicFrame = "source-anchored-v1"

func nativeTableFrameMatches(element NativeElement) bool {
	if element.Table == nil || element.Transform.Cx == nil || element.Transform.Cy == nil {
		return false
	}
	width, widthOK := nativeExactTableTrackTotal(element.Table.ColumnWidths)
	height, heightOK := nativeExactTableTrackTotal(element.Table.RowHeights)
	return widthOK && heightOK && width == *element.Transform.Cx && height == *element.Transform.Cy
}
func nativeMarkGraphicFrameLayout(element *NativeElement) error {
	if element.GraphicFrameLayout != nil {
		return nil
	}
	if len(element.Compatibility.Diagnostics) >= nativeMaxDiagnosticsPerScope {
		return refuseNativeGraphicFrame("pptx.graphic-frame-diagnostic-budget-unavailable", "source-anchored graphic-frame policy exceeds the bounded diagnostic scope")
	}
	element.GraphicFrameLayout = stringPointer(nativeSourceAnchoredGraphicFrame)
	element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
	element.Compatibility.Diagnostics = append(element.Compatibility.Diagnostics, NativeDiagnostic{Severity: NativeDiagnosticSeverityWarning, Code: "pptx.source-graphic-frame-preview", Message: "source-anchored graphic-frame preview preserves intrinsic table tracks and raw source; PowerPoint importer rewrites are not emulated"})
	return nil
}
func nativeMarkGraphicFrameDescendants(elements []NativeElement) (bool, error) {
	found := false
	for i := range elements {
		element := &elements[i]
		if element.Kind == NativeElementKindTable || element.Kind == NativeElementKindChart {
			if err := nativeMarkGraphicFrameLayout(element); err != nil {
				return false, err
			}
			found = true
		}
		nested, err := nativeMarkGraphicFrameDescendants(element.Children)
		if err != nil {
			return false, err
		}
		if nested {
			element.Compatibility.Status = worseNativeStatus(element.Compatibility.Status, NativeCompatibilityStatusPreserveOnly)
			found = true
		}
	}
	return found, nil
}
