package xlsxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
)

// NativeWorkbookV2ExtractionOptions retains application/document identity across
// v2 revisions. Previous must itself be a valid v2 contract for the same workbook.
type NativeWorkbookV2ExtractionOptions struct {
	Previous   *NativeWorkbookV2
	DocumentID string
}

func canonicalNativeWorkbookV2Capabilities() []NativeWorkbookCapabilityV2 {
	return []NativeWorkbookCapabilityV2{
		{Name: "native-ooxml-parse", Level: "read-only", Detail: nativeWorkbookString("HTML-free SpreadsheetML extraction with lexical values and formula caches")},
		{Name: "native-geometry", Level: "exact", Detail: nativeWorkbookString("Integer-EMU row/column bands and merged rectangles")},
		{Name: "native-decorations", Level: "exact", Detail: nativeWorkbookString("Fills plus producer-issued orthogonal stroke/border segments")},
		{Name: "native-grid-commands", Level: "exact", Detail: nativeWorkbookString("Renderer-neutral geometry and decoration command replay")},
		{Name: "unsupported-content", Level: "preserve-exact", Detail: nativeWorkbookString("Unmodeled OOXML remains authoritative in the original package")},
	}
}

// ExtractNativeWorkbookV2 is the renderer-free XLSX v2 import boundary. It reuses
// the lockstep v1 extractor, then emits the v2 protocol identity and capability
// inventory. It does not claim cell glyph/display paint, number-format display,
// or host fonts.
func ExtractNativeWorkbookV2(data []byte) (*NativeWorkbookV2, error) {
	return ExtractNativeWorkbookV2WithOptions(data, NativeWorkbookV2ExtractionOptions{})
}

// ExtractNativeWorkbookV2WithOptions retains application/document identity
// across revisions. Previous is accepted only when it is a valid v2 contract
// for the same canonical workbook.
func ExtractNativeWorkbookV2WithOptions(data []byte, options NativeWorkbookV2ExtractionOptions) (*NativeWorkbookV2, error) {
	v1Options := NativeWorkbookExtractionOptions{DocumentID: options.DocumentID}
	if options.Previous != nil {
		if issues := ValidateNativeWorkbookV2(options.Previous); len(issues) != 0 {
			return nil, fmt.Errorf("xlsxpatch: native extract: previous identity contract is invalid: %w", &NativeWorkbookV2ValidationError{Issues: issues})
		}
		previous, err := nativeWorkbookV2AsV1Identity(options.Previous)
		if err != nil {
			return nil, fmt.Errorf("xlsxpatch: native extract: previous identity contract is invalid: %w", err)
		}
		v1Options.Previous = previous
	}
	v1, err := ExtractNativeWorkbookV1WithOptions(data, v1Options)
	if err != nil {
		return nil, err
	}
	v2, err := nativeWorkbookV1ToV2(v1)
	if err != nil {
		return nil, err
	}
	if err := enrichNativeWorkbookV2Display(data, v2); err != nil {
		return nil, err
	}
	if issues := ValidateNativeWorkbookV2(v2); len(issues) != 0 {
		return nil, fmt.Errorf("xlsxpatch: native extract produced invalid contract: %w", &NativeWorkbookV2ValidationError{Issues: issues})
	}
	if _, err := EncodeNativeWorkbookV2(v2); err != nil {
		return nil, err
	}
	return v2, nil
}

func nativeWorkbookV1ToV2(v1 *NativeWorkbookV1) (*NativeWorkbookV2, error) {
	raw, err := json.Marshal(v1)
	if err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: project v2: %w", err)
	}
	var v2 NativeWorkbookV2
	if err := json.Unmarshal(raw, &v2); err != nil {
		return nil, fmt.Errorf("xlsxpatch: native extract: project v2: %w", err)
	}
	v2.Protocol = NativeXLSXV2Protocol
	v2.Version = NativeXLSXV2Version
	v2.SchemaSHA256 = NativeXLSXV2SchemaSHA256
	v2.Capabilities = canonicalNativeWorkbookV2Capabilities()
	copyNativeWorkbookRichRunsV1ToV2(v1, &v2)
	return &v2, nil
}

func copyNativeWorkbookRichRunsV1ToV2(v1 *NativeWorkbookV1, v2 *NativeWorkbookV2) {
	if len(v1.Sheets) != len(v2.Sheets) {
		return
	}
	for sheetIndex := range v1.Sheets {
		v1Sheet := v1.Sheets[sheetIndex]
		v2Sheet := &v2.Sheets[sheetIndex]
		if len(v1Sheet.Cells) != len(v2Sheet.Cells) {
			continue
		}
		for cellIndex := range v1Sheet.Cells {
			v1Cell := v1Sheet.Cells[cellIndex]
			v2Cell := &v2Sheet.Cells[cellIndex]
			if v1Cell.Value != nil && v2Cell.Value != nil && len(v1Cell.Value.runs) != 0 {
				v2Cell.Value.Runs = cloneRichRuns(v1Cell.Value.runs)
			}
			if v1Cell.Formula != nil && v1Cell.Formula.Cached != nil && v2Cell.Formula != nil && v2Cell.Formula.Cached != nil && len(v1Cell.Formula.Cached.runs) != 0 {
				v2Cell.Formula.Cached.Runs = cloneRichRuns(v1Cell.Formula.Cached.runs)
			}
		}
	}
}

func cloneRichRuns(runs []NativeWorkbookRichRunV2) []NativeWorkbookRichRunV2 {
	if len(runs) == 0 {
		return nil
	}
	out := make([]NativeWorkbookRichRunV2, len(runs))
	copy(out, runs)
	for index := range out {
		out[index].FontName = cloneString(out[index].FontName)
		out[index].Bold = cloneBool(out[index].Bold)
		out[index].Italic = cloneBool(out[index].Italic)
		out[index].FontSizePoints = cloneFloat(out[index].FontSizePoints)
		out[index].FontColor = cloneString(out[index].FontColor)
		out[index].scheme = cloneString(out[index].scheme)
		out[index].themeIndex = cloneInt(out[index].themeIndex)
		out[index].themeTint = cloneFloat(out[index].themeTint)
	}
	return out
}

func nativeWorkbookV2AsV1Identity(v2 *NativeWorkbookV2) (*NativeWorkbookV1, error) {
	raw, err := json.Marshal(v2)
	if err != nil {
		return nil, err
	}
	var v1 NativeWorkbookV1
	if err := json.Unmarshal(raw, &v1); err != nil {
		return nil, err
	}
	v1.Protocol = NativeXLSXProtocol
	v1.Version = NativeXLSXVersion
	v1.Capabilities = []NativeWorkbookCapabilityV1{
		{Name: "native-ooxml-parse", Level: "read-only", Detail: nativeWorkbookString("HTML-free SpreadsheetML extraction with lexical values and formula caches")},
		{Name: "native-v1-mutations", Level: "partial", Detail: nativeWorkbookString("Cells, formulas, row/column dimensions, and the bounded style.patch projection")},
		{Name: "unsupported-content", Level: "preserve-exact", Detail: nativeWorkbookString("Unmodeled OOXML remains authoritative in the original package")},
	}
	if issues := ValidateNativeWorkbookV1(&v1); len(issues) != 0 {
		return nil, &NativeWorkbookValidationError{Issues: issues}
	}
	return &v1, nil
}

func nativeRawStyleProjectionDigestV2(projection NativeWorkbookEffectiveStyleV2) (string, error) {
	var buffer bytes.Buffer
	encoder := json.NewEncoder(&buffer)
	encoder.SetEscapeHTML(false)
	if err := encoder.Encode(projection); err != nil {
		return "", err
	}
	return nativeWorkbookDigest(bytes.TrimSuffix(buffer.Bytes(), []byte{'\n'})), nil
}
