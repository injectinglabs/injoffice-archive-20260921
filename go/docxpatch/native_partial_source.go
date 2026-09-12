package docxpatch

import (
	"encoding/json"
	"fmt"
)

// InspectNativePartialSourceV1 joins extraction and style resolution from the
// same immutable caller bytes. It supplies no font assets or mutation authority.
func InspectNativePartialSourceV1(data []byte) ([]byte, error) {
	if len(data) == 0 || len(data) > NativeDOCXMaxPackageBytes {
		return nil, fmt.Errorf("partial source package size must be 1..%d bytes", NativeDOCXMaxPackageBytes)
	}
	data = append([]byte(nil), data...)
	document, err := ExtractNativeDocumentV1(data)
	if err != nil {
		return nil, err
	}
	layout, err := ResolveNativeDocumentLayoutV1(data)
	if err != nil {
		return nil, err
	}
	if document.DocumentID != layout.DocumentID || document.Revision != layout.Revision || document.Source.MainPart != layout.SourceParts.MainPart {
		return nil, fmt.Errorf("partial source identity mismatch")
	}
	docJSON, err := EncodeNativeDocumentV1(document)
	if err != nil {
		return nil, err
	}
	layoutJSON, err := EncodeNativeResolvedLayoutInputV1(layout)
	if err != nil {
		return nil, err
	}
	equations, err := inspectNativePartialEquations(data, document)
	if err != nil {
		return nil, err
	}
	notices, err := inspectNativeEquationContext(data, document, layout, equations)
	if err != nil {
		return nil, err
	}
	result, err := json.Marshal(struct {
		Protocol               string                          `json:"protocol"`
		Version                int                             `json:"version"`
		PackageSHA256          string                          `json:"package_sha256"`
		Document               json.RawMessage                 `json:"document"`
		ResolvedLayout         json.RawMessage                 `json:"resolved_layout"`
		Equations              []NativePartialEquationV1       `json:"equations,omitempty"`
		EquationContextNotices []NativeEquationContextNoticeV1 `json:"equation_context_notices,omitempty"`
	}{"injoffice.docx.partial-source", 1, document.Source.PackageSHA256, docJSON, layoutJSON, equations, notices})
	if err != nil {
		return nil, err
	}
	if len(result) > 2*NativeDOCXMaxJSONBytes {
		return nil, fmt.Errorf("partial source JSON exceeds bounded response size")
	}
	return result, nil
}
