package officehttp

import (
	"context"
	"encoding/base64"
	"encoding/json"
	"fmt"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

const pptxWorkbookPayloadLimit = 8 * 1024 * 1024

type pptxWorkbookContract struct {
	Part         string `json:"part"`
	SHA256       string `json:"sha256"`
	ContractJSON string `json:"contract_json,omitempty"`
	Refusal      string `json:"refusal,omitempty"`
}
type pptxWorkbookPayload struct {
	InspectionJSON string                 `json:"inspection_json"`
	Workbooks      []pptxWorkbookContract `json:"workbooks"`
}

// Both engines read the same source bytes. TS revalidates each engine contract
// and rebinds the source identities; no pre-resolved chart JSON is trusted.
func pptxChartWorkbookPayload(ctx context.Context, data []byte) (pptxWorkbookPayload, error) {
	var out pptxWorkbookPayload
	if err := ctx.Err(); err != nil {
		return out, err
	}
	inspection, err := pptxpatch.InspectNativePPTXChartWorkbooks(data)
	if err != nil {
		return out, err
	}
	if err := ctx.Err(); err != nil {
		return out, err
	}
	encoded, err := json.Marshal(inspection)
	if err != nil {
		return out, err
	}
	out.InspectionJSON = string(encoded)
	out.Workbooks = []pptxWorkbookContract{}

	if err := checkPPTXWorkbookPayload(out); err != nil {
		return out, err
	}
	for _, resource := range inspection.Workbooks {
		if err := ctx.Err(); err != nil {
			return out, err
		}
		record := pptxWorkbookContract{Part: resource.Part, SHA256: resource.SHA256}
		bytes, decodeErr := base64.StdEncoding.DecodeString(resource.BytesBase64)
		if decodeErr != nil {
			return out, decodeErr
		}
		workbook, extractErr := xlsxpatch.ExtractNativeWorkbookV2(bytes)
		if extractErr != nil {
			record.Refusal = "Embedded workbook is outside the native XLSX extraction profile"
		} else {
			contract, encodeErr := xlsxpatch.EncodeNativeWorkbookV2(workbook)
			if encodeErr != nil {
				return out, encodeErr
			}
			record.ContractJSON = string(contract)
		}
		out.Workbooks = append(out.Workbooks, record)
		if err := checkPPTXWorkbookPayload(out); err != nil {
			return out, err
		}
	}
	return out, nil
}

func validPPTXWorkbookPreviewMode(result []byte, workbook, literal bool) bool {
	var fields map[string]json.RawMessage
	if json.Unmarshal(result, &fields) != nil {
		return false
	}
	flag, has := fields["workbook_chart_preview"]
	if !workbook {
		return !has && validPPTXChartPreviewMode(result, literal)
	}
	if literal || !has || string(flag) != "true" || fields["source_chart_preview"] != nil {
		return false
	}
	var policy string
	return json.Unmarshal(fields["chart_axis_layout_policy"], &policy) == nil && policy == "supplied-outline-margins-v1"
}

func checkPPTXWorkbookPayload(payload pptxWorkbookPayload) error {
	encoded, err := json.Marshal(payload)
	if err != nil {
		return err
	}
	if len(encoded) > pptxWorkbookPayloadLimit {
		return fmt.Errorf("embedded workbook preview payload exceeds 8 MiB aggregate UTF-8 budget")
	}
	return nil
}
