//go:build js && wasm

// Command pptxnativewasm exposes pptxpatch's native v1 extraction and surgical
// mutation boundary to a browser worker. It never mints server capabilities.
package main

import (
	"fmt"
	"syscall/js"
)

func main() {
	api := js.Global().Get("Object").New()
	api.Set("extract", guarded(jsExtract))
	api.Set("inspect", guarded(jsInspect))
	api.Set("apply", guarded(jsApply))
	js.Global().Set("pptxnative", api)
	if ready := js.Global().Get("pptxnativeOnReady"); ready.Type() == js.TypeFunction {
		ready.Invoke()
	}
	select {}
}

// guarded recovers inside the callback so a refused operation cannot escape
// syscall/js.handleEvent and terminate the reusable WASM instance.
func guarded(fn func(this js.Value, args []js.Value) any) js.Func {
	return js.FuncOf(func(this js.Value, args []js.Value) (result any) {
		defer func() {
			if recovered := recover(); recovered != nil {
				result = failFatal(fmt.Sprint(recovered))
			}
		}()
		return fn(this, args)
	})
}

func ok(value any) js.Value {
	result := js.Global().Get("Object").New()
	result.Set("ok", true)
	result.Set("value", value)
	return result
}

func fail(message string) js.Value {
	return failure(message, false)
}

// failFatal marks recovered panics separately from expected validation and CAS
// refusals. The worker must be discarded because native state may be unknown.
func failFatal(message string) js.Value {
	return failure(message, true)
}

func failure(message string, fatal bool) js.Value {
	if message == "" {
		message = "pptxnative failed"
	}
	result := js.Global().Get("Object").New()
	result.Set("ok", false)
	result.Set("error", message)
	result.Set("fatal", fatal)
	return result
}

// extract(bytes) -> {ok:true, value:string} | {ok:false, error:string}.
func jsExtract(_ js.Value, args []js.Value) any {
	if len(args) != 1 {
		return fail("extract(bytes) requires 1 argument")
	}
	data, err := bytesFromJS(args[0])
	if err != nil {
		return fail(err.Error())
	}
	encoded, err := extractNativeJSON(data)
	if err != nil {
		return fail(err.Error())
	}
	return ok(string(encoded))
}

func jsInspect(_ js.Value, args []js.Value) any {
	if len(args) != 1 {
		return fail("inspect(bytes) requires 1 argument")
	}
	data, err := bytesFromJS(args[0])
	if err != nil {
		return fail(err.Error())
	}
	encoded, err := inspectNativeJSON(data)
	if err != nil {
		return fail(err.Error())
	}
	return ok(string(encoded))
}

// apply(original, payload, expectedRevision) returns mutated PPTX bytes.
func jsApply(_ js.Value, args []js.Value) any {
	if len(args) != 3 {
		return fail("apply(original, payload, expectedRevision) requires 3 arguments")
	}
	original, err := bytesFromJS(args[0])
	if err != nil {
		return fail(err.Error())
	}
	payload, err := payloadFromJS(args[1])
	if err != nil {
		return fail(err.Error())
	}
	expectedRevision := jsString(args[2])
	if expectedRevision == "" {
		return fail("expectedRevision is required (sha256:<digest>)")
	}
	produced, err := applyNativePayload(original, payload, expectedRevision)
	if err != nil {
		return fail(err.Error())
	}
	packed, err := bytesToJS(produced)
	if err != nil {
		return fail(err.Error())
	}
	return ok(packed)
}

func jsString(value js.Value) string {
	switch value.Type() {
	case js.TypeString, js.TypeNumber:
		return value.String()
	default:
		return ""
	}
}

func payloadFromJS(value js.Value) ([]byte, error) {
	switch value.Type() {
	case js.TypeString:
		return []byte(value.String()), nil
	case js.TypeObject:
		return bytesFromJS(value)
	default:
		return nil, fmt.Errorf("payload must be a JSON string or Uint8Array")
	}
}

func bytesFromJS(value js.Value) ([]byte, error) {
	if value.IsNull() || value.IsUndefined() {
		return nil, fmt.Errorf("expected Uint8Array or ArrayBuffer")
	}
	uint8Array := js.Global().Get("Uint8Array")
	switch {
	case value.InstanceOf(uint8Array):
	case value.InstanceOf(js.Global().Get("ArrayBuffer")):
		value = uint8Array.New(value)
	default:
		if constructor := value.Get("constructor"); constructor.Truthy() && constructor.Get("name").String() == "Uint8Array" {
			break
		}
		if buffer := value.Get("buffer"); buffer.Truthy() && value.Get("byteLength").Type() == js.TypeNumber {
			value = uint8Array.New(buffer, value.Get("byteOffset").Int(), value.Get("byteLength").Int())
			break
		}
		return nil, fmt.Errorf("expected Uint8Array or ArrayBuffer")
	}
	size := value.Get("byteLength").Int()
	data := make([]byte, size)
	if copied := js.CopyBytesToGo(data, value); copied != size {
		return nil, fmt.Errorf("copied %d bytes, want %d", copied, size)
	}
	return data, nil
}

func bytesToJS(data []byte) (js.Value, error) {
	result := js.Global().Get("Uint8Array").New(len(data))
	if copied := js.CopyBytesToJS(result, data); copied != len(data) {
		return js.Undefined(), fmt.Errorf("copied %d bytes, want %d", copied, len(data))
	}
	return result, nil
}
