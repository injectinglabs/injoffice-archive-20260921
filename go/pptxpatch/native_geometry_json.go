package pptxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
)

// Numeric zeros and stroke=false are meaningful required fields. Preserve their
// JSON presence without adding optionality to the evaluated Go engine model.
func decodeNativeGeometryObject(data []byte, target any, required ...string) error {
	var fields map[string]json.RawMessage
	if err := json.Unmarshal(data, &fields); err != nil {
		return err
	}
	for _, key := range required {
		value, ok := fields[key]
		if !ok || bytes.Equal(bytes.TrimSpace(value), []byte("null")) {
			return fmt.Errorf("missing geometry field %s", key)
		}
	}
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	return decoder.Decode(target)
}
func (r *NativeGeometryTextRect) UnmarshalJSON(data []byte) error {
	type plain NativeGeometryTextRect
	var result plain
	if err := decodeNativeGeometryObject(data, &result, "x", "y", "cx", "cy"); err != nil {
		return err
	}
	*r = NativeGeometryTextRect(result)
	return nil
}
func (p *NativeGeometryPath) UnmarshalJSON(data []byte) error {
	type plain NativeGeometryPath
	var result plain
	if err := decodeNativeGeometryObject(data, &result, "fillMode", "stroke", "commands"); err != nil {
		return err
	}
	*p = NativeGeometryPath(result)
	return nil
}
