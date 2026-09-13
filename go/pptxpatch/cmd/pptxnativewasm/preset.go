package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"regexp"

	"github.com/injectinglabs/injoffice/go/pptxpatch"
)

type presetRequest struct {
	Name        string           `json:"name"`
	Width       int64            `json:"widthEmu"`
	Height      int64            `json:"heightEmu"`
	Adjustments map[string]int64 `json:"adjustments"`
}

// The bridge accepts only bounded JSON objects with unique names. This avoids
// different callers interpreting duplicate request fields differently.
func presetObject(dec *json.Decoder, depth int) error {
	if depth > 2 {
		return fmt.Errorf("preset request nesting exceeded")
	}
	token, err := dec.Token()
	if err != nil {
		return err
	}
	if token != json.Delim('{') {
		return fmt.Errorf("preset request requires objects")
	}
	seen := map[string]bool{}
	for dec.More() {
		key, err := dec.Token()
		if err != nil {
			return err
		}
		name, ok := key.(string)
		if !ok || seen[name] {
			return fmt.Errorf("duplicate preset request field")
		}
		// encoding/json struct matching folds case. The wire contract does not.
		if depth == 0 {
			switch name {
			case "name", "widthEmu", "heightEmu", "adjustments":
			default:
				return fmt.Errorf("unknown preset request field %q", name)
			}
		}
		seen[name] = true
		if len(seen) > 1024 {
			return fmt.Errorf("preset request field budget exceeded")
		}
		var raw json.RawMessage
		if err := dec.Decode(&raw); err != nil {
			return err
		}
		if len(raw) > 0 && raw[0] == '{' {
			if err := presetObject(json.NewDecoder(bytes.NewReader(raw)), depth+1); err != nil {
				return err
			}
		}
	}
	_, err = dec.Token()
	return err
}

func evaluatePresetJSON(payload []byte) ([]byte, error) {
	if len(payload) == 0 || len(payload) > 65536 {
		return nil, fmt.Errorf("preset evaluation request budget exceeded")
	}
	if err := presetObject(json.NewDecoder(bytes.NewReader(payload)), 0); err != nil {
		return nil, err
	}
	decoder := json.NewDecoder(bytes.NewReader(payload))
	decoder.DisallowUnknownFields()
	var request presetRequest
	if err := decoder.Decode(&request); err != nil {
		return nil, err
	}
	if err := decoder.Decode(new(any)); err != io.EOF {
		return nil, fmt.Errorf("preset request trailing data")
	}
	if !regexp.MustCompile(`^[A-Za-z][A-Za-z0-9]{0,63}$`).MatchString(request.Name) || request.Width <= 0 || request.Height <= 0 {
		return nil, fmt.Errorf("invalid preset name or frame")
	}
	// Null is not an integer and explicit null adjustment objects are ambiguous.
	var fields map[string]json.RawMessage
	_ = json.Unmarshal(payload, &fields)
	if raw, ok := fields["adjustments"]; ok && bytes.Equal(bytes.TrimSpace(raw), []byte("null")) {
		return nil, fmt.Errorf("adjustments must be an object")
	}
	for _, raw := range fields {
		if bytes.Equal(bytes.TrimSpace(raw), []byte("null")) {
			return nil, fmt.Errorf("preset request fields cannot be null")
		}
	}
	if request.Adjustments == nil {
		request.Adjustments = map[string]int64{}
	}
	var values map[string]json.RawMessage
	_ = json.Unmarshal(fields["adjustments"], &values)
	for name := range request.Adjustments {
		if bytes.Equal(bytes.TrimSpace(values[name]), []byte("null")) {
			return nil, fmt.Errorf("adjustment values cannot be null")
		}
		value := request.Adjustments[name]
		if !regexp.MustCompile(`^[A-Za-z_][A-Za-z0-9_]{0,127}$`).MatchString(name) || value < -9007199254740991 || value > 9007199254740991 || bytes.Equal(bytes.TrimSpace(values[name]), []byte("-0")) {
			return nil, fmt.Errorf("invalid preset adjustment")
		}
	}
	geometry, err := pptxpatch.EvaluateNativePPTXPresetGeometry(request.Name, request.Width, request.Height, request.Adjustments)
	if err != nil {
		return nil, err
	}
	return json.Marshal(struct {
		Protocol string                             `json:"protocol"`
		Request  presetRequest                      `json:"request"`
		Geometry *pptxpatch.NativeEvaluatedGeometry `json:"geometry"`
	}{"pptx-preset-evaluation-v1", request, geometry})
}
