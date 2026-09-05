package collab

import (
	"bytes"
	"encoding/json"
	"io"
)

func cloneRaw(raw json.RawMessage) json.RawMessage {
	if raw == nil {
		return nil
	}
	out := make(json.RawMessage, len(raw))
	copy(out, raw)
	return out
}

func walkJSONDepth(raw json.RawMessage, maxDepth int) error {
	if !json.Valid(raw) {
		return errMalformedJSON
	}
	dec := json.NewDecoder(bytes.NewReader(raw))
	depth := 0
	for {
		tok, err := dec.Token()
		if err == io.EOF {
			return nil
		}
		if err != nil {
			return errMalformedJSON
		}
		d, ok := tok.(json.Delim)
		if !ok {
			continue
		}
		switch d {
		case '{', '[':
			depth++
			if depth > maxDepth {
				return ErrJSONTooDeep
			}
		case '}', ']':
			depth--
		}
	}
}

// sentinel distinguished from exported invalid-ops/selection errors.
var errMalformedJSON = errMalformed{}

type errMalformed struct{}

func (errMalformed) Error() string { return "malformed json" }

func validateOps(ops json.RawMessage) error {
	if len(ops) > MaxOpBytes {
		return ErrOpTooLarge
	}
	if len(ops) == 0 {
		return ErrInvalidOps
	}
	if err := walkJSONDepth(ops, MaxJSONDepth); err != nil {
		if err == ErrJSONTooDeep {
			return err
		}
		return ErrInvalidOps
	}
	var arr []json.RawMessage
	if err := json.Unmarshal(ops, &arr); err != nil {
		return ErrInvalidOps
	}
	if len(arr) == 0 {
		return ErrInvalidOps
	}
	if len(arr) > MaxOpsPerBatch {
		return ErrBatchTooLarge
	}
	return nil
}

func validateSelection(sel json.RawMessage) error {
	if len(sel) > MaxSelectionBytes {
		return ErrSelectionTooLarge
	}
	if len(sel) == 0 {
		return ErrInvalidSelection
	}
	if err := walkJSONDepth(sel, MaxJSONDepth); err != nil {
		if err == ErrJSONTooDeep {
			return err
		}
		return ErrInvalidSelection
	}
	return nil
}
