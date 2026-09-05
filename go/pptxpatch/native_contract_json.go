package pptxpatch

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io"
	"sort"
	"strconv"
	"unicode/utf16"
)

// DecodeNativePPTXJSON decodes exactly one v1 contract document, rejects
// unknown properties at every struct level, and applies semantic validation.
func DecodeNativePPTXJSON(data []byte) (NativePPTXDeck, error) {
	if len(data) > nativeMaxJsonBytes {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: input exceeds %d bytes", nativeMaxJsonBytes)
	}
	lexical := json.NewDecoder(bytes.NewReader(data))
	lexical.UseNumber()
	budget := nativeJSONBudget{}
	if err := validateNativeJSONValue(lexical, "$", 0, &budget); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: %w", err)
	}
	if token, err := lexical.Token(); err != io.EOF {
		if err == nil {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: trailing token %v", token)
		}
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: %w", err)
	}

	var deck NativePPTXDeck
	decoder := json.NewDecoder(bytes.NewReader(data))
	decoder.DisallowUnknownFields()
	if err := decoder.Decode(&deck); err != nil {
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: %w", err)
	}
	var trailing any
	if err := decoder.Decode(&trailing); err != io.EOF {
		if err == nil {
			return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: trailing JSON value")
		}
		return NativePPTXDeck{}, fmt.Errorf("pptxpatch: decode native PPTX contract: %w", err)
	}
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		return NativePPTXDeck{}, NativeContractValidationError{Issues: issues}
	}
	return deck, nil
}

// The v1 schema admits neither null nor duplicate object keys. encoding/json
// normally accepts both (null becomes a zero pointer and the final duplicate
// wins), which would make two distinct inputs canonicalize to the same deck.
type nativeJSONBudget struct{ nodes int }

func validateNativeJSONValue(decoder *json.Decoder, p string, depth int, budget *nativeJSONBudget) error {
	if depth > nativeMaxDepth {
		return fmt.Errorf("%s: nesting exceeds %d", p, nativeMaxDepth)
	}
	budget.nodes++
	if budget.nodes > nativeMaxNodes {
		return fmt.Errorf("%s: contract exceeds %d JSON nodes", p, nativeMaxNodes)
	}
	token, err := decoder.Token()
	if err != nil {
		return err
	}
	if token == nil {
		return fmt.Errorf("%s: null is not allowed", p)
	}
	delim, ok := token.(json.Delim)
	if !ok {
		if number, numeric := token.(json.Number); numeric && string(number) == "-0" {
			return fmt.Errorf("%s: negative zero is not allowed", p)
		}
		return nil
	}
	switch delim {
	case '{':
		seen := map[string]bool{}
		for decoder.More() {
			keyToken, err := decoder.Token()
			if err != nil {
				return err
			}
			key, ok := keyToken.(string)
			if !ok {
				return fmt.Errorf("%s: object key is not a string", p)
			}
			if seen[key] {
				return fmt.Errorf("%s.%s: duplicate object key", p, key)
			}
			seen[key] = true
			if err := validateNativeJSONValue(decoder, p+"."+key, depth+1, budget); err != nil {
				return err
			}
		}
		end, err := decoder.Token()
		if err != nil || end != json.Delim('}') {
			return fmt.Errorf("%s: unterminated object", p)
		}
	case '[':
		index := 0
		for decoder.More() {
			if err := validateNativeJSONValue(decoder, fmt.Sprintf("%s[%d]", p, index), depth+1, budget); err != nil {
				return err
			}
			index++
		}
		end, err := decoder.Token()
		if err != nil || end != json.Delim(']') {
			return fmt.Errorf("%s: unterminated array", p)
		}
	default:
		return fmt.Errorf("%s: unexpected delimiter %q", p, delim)
	}
	return nil
}

// NativeContractValidationError keeps all bounded, machine-readable issues
// available to a gateway while still implementing error.
type NativeContractValidationError struct {
	Issues []NativeContractIssue
}

func (e NativeContractValidationError) Error() string {
	if len(e.Issues) == 0 {
		return "pptxpatch: invalid native PPTX contract"
	}
	return fmt.Sprintf("pptxpatch: invalid native PPTX contract: %s: %s", e.Issues[0].Path, e.Issues[0].Message)
}

// MarshalNativePPTXJSON validates and emits deterministic UTF-8 JSON. Object
// keys are sorted by UTF-16 code units. Assets are sorted by durable id;
// every order-bearing array (slides, elements/z-order, paragraphs, rows, and
// runs) remains untouched.
func MarshalNativePPTXJSON(deck NativePPTXDeck) ([]byte, error) {
	if issues := ValidateNativePPTX(deck); len(issues) != 0 {
		return nil, NativeContractValidationError{Issues: issues}
	}
	normalized := deck
	normalized.Assets = make([]NativeAsset, len(deck.Assets))
	copy(normalized.Assets, deck.Assets)
	sort.SliceStable(normalized.Assets, func(i, j int) bool {
		return lessUTF16(normalized.Assets[i].ID, normalized.Assets[j].ID)
	})
	raw, err := json.Marshal(normalized)
	if err != nil {
		return nil, fmt.Errorf("pptxpatch: marshal native PPTX contract: %w", err)
	}
	decoder := json.NewDecoder(bytes.NewReader(raw))
	decoder.UseNumber()
	var value any
	if err := decoder.Decode(&value); err != nil {
		return nil, fmt.Errorf("pptxpatch: canonicalize native PPTX contract: %w", err)
	}
	var out bytes.Buffer
	if err := writeCanonicalJSON(&out, value); err != nil {
		return nil, err
	}
	if out.Len() > nativeMaxJsonBytes {
		return nil, fmt.Errorf("pptxpatch: marshal native PPTX contract: output exceeds %d bytes", nativeMaxJsonBytes)
	}
	return out.Bytes(), nil
}

func writeCanonicalJSON(out *bytes.Buffer, value any) error {
	switch typed := value.(type) {
	case nil:
		out.WriteString("null")
	case bool:
		if typed {
			out.WriteString("true")
		} else {
			out.WriteString("false")
		}
	case string:
		quoted, err := quoteJSONString(typed)
		if err != nil {
			return err
		}
		out.Write(quoted)
	case json.Number:
		integer, err := strconv.ParseInt(string(typed), 10, 64)
		if err != nil || integer < -nativeMaxSafeInteger || integer > nativeMaxSafeInteger {
			return fmt.Errorf("pptxpatch: canonical JSON contains a non-safe integer %q", typed)
		}
		out.WriteString(strconv.FormatInt(integer, 10))
	case []any:
		out.WriteByte('[')
		for index, item := range typed {
			if index != 0 {
				out.WriteByte(',')
			}
			if err := writeCanonicalJSON(out, item); err != nil {
				return err
			}
		}
		out.WriteByte(']')
	case map[string]any:
		keys := make([]string, 0, len(typed))
		for key := range typed {
			keys = append(keys, key)
		}
		sort.Slice(keys, func(i, j int) bool { return lessUTF16(keys[i], keys[j]) })
		out.WriteByte('{')
		for index, key := range keys {
			if index != 0 {
				out.WriteByte(',')
			}
			quoted, err := quoteJSONString(key)
			if err != nil {
				return err
			}
			out.Write(quoted)
			out.WriteByte(':')
			if err := writeCanonicalJSON(out, typed[key]); err != nil {
				return err
			}
		}
		out.WriteByte('}')
	default:
		return fmt.Errorf("pptxpatch: unsupported canonical JSON value %T", typed)
	}
	return nil
}

func quoteJSONString(value string) ([]byte, error) {
	var out bytes.Buffer
	encoder := json.NewEncoder(&out)
	encoder.SetEscapeHTML(false)
	if err := encoder.Encode(value); err != nil {
		return nil, fmt.Errorf("pptxpatch: quote canonical JSON string: %w", err)
	}
	return bytes.TrimSuffix(out.Bytes(), []byte("\n")), nil
}

// ECMAScript's ordinary string comparison and RFC 8785 both compare UTF-16
// code units. Go's native string order compares UTF-8 bytes, so use the same
// explicit order as the TypeScript binding for cross-runtime determinism.
func lessUTF16(left, right string) bool {
	a, b := utf16.Encode([]rune(left)), utf16.Encode([]rune(right))
	for index := 0; index < len(a) && index < len(b); index++ {
		if a[index] != b[index] {
			return a[index] < b[index]
		}
	}
	return len(a) < len(b)
}
