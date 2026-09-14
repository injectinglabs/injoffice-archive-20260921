package main

import (
	"encoding/json"
	"fmt"
	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

func inspectChartWorkbooksJSON(data []byte) ([]byte, error) {
	result, err := pptxpatch.InspectNativePPTXChartWorkbooks(data)
	if err != nil {
		return nil, err
	}
	encoded, err := json.Marshal(result)
	if err != nil {
		return nil, err
	}
	if len(encoded) > 32*1024*1024 {
		return nil, fmt.Errorf("chart workbook inspection response budget exceeded")
	}
	return encoded, nil
}
