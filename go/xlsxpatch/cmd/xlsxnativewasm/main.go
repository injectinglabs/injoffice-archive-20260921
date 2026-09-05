//go:build js && wasm

// Command xlsxnativewasm wraps the same xlsxpatch extract/apply functions as
// cmd/xlsxnative for the optional @injoffice/xlsx-wasm browser worker.
package main

import (
	"fmt"
	"syscall/js"

	"github.com/injectinglabs/injoffice/go/xlsxpatch"
)

func main() {
	obj := js.Global().Get("Object").New()
	obj.Set("extract", guarded(jsExtract))
	obj.Set("apply", guarded(jsApply))
	js.Global().Set("xlsxnative", obj)
	if ready := js.Global().Get("xlsxnativeOnReady"); ready.Type() == js.TypeFunction {
		ready.Invoke()
	}
	select {}
}

// guarded recovers inside the JS callback so a refused extract/apply cannot
// reach syscall/js.handleEvent and exit the Go WASM instance.
func guarded(fn func(this js.Value, args []js.Value) any) js.Func {
	return js.FuncOf(func(this js.Value, args []js.Value) (result any) {
		defer func() {
			if r := recover(); r != nil {
				result = failFatal(fmt.Sprint(r))
			}
		}()
		return fn(this, args)
	})
}

func ok(value any) js.Value {
	obj := js.Global().Get("Object").New()
	obj.Set("ok", true)
	obj.Set("value", value)
	return obj
}

func fail(message string) js.Value {
	return failure(message, false)
}

// failFatal marks recovered panics separately from expected validation and CAS
// refusals. A caller should discard the worker after a fatal response because
// the operation may have left package-processing state in an unknown state.
func failFatal(message string) js.Value {
	return failure(message, true)
}

func failure(message string, fatal bool) js.Value {
	if message == "" {
		message = "xlsxnative failed"
	}
	obj := js.Global().Get("Object").New()
	obj.Set("ok", false)
	obj.Set("error", message)
	obj.Set("fatal", fatal)
	return obj
}

// extract(bytes) -> {ok:true, value: json} | {ok:false, error: string, fatal: boolean}.
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

// apply(original, payload, expectedRevision) -> {ok:true, value: Uint8Array} | {ok:false, error: string, fatal: boolean}.
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
	outer := jsString(args[2])
	if outer == "" {
		return fail("expectedRevision is required (sha256:<digest>)")
	}
	produced, err := xlsxpatch.ApplyNativeWorkbookMutationPayloadV1(original, payload, outer)
	if err != nil {
		return fail(err.Error())
	}
	packed, err := bytesToJS(produced.Package)
	if err != nil {
		return fail(err.Error())
	}
	return ok(packed)
}

func extractNativeJSON(data []byte) ([]byte, error) {
	workbook, err := xlsxpatch.ExtractNativeWorkbookV2(data)
	if err != nil {
		return nil, err
	}
	return xlsxpatch.EncodeNativeWorkbookV2(workbook)
}

func jsString(v js.Value) string {
	switch v.Type() {
	case js.TypeString, js.TypeNumber:
		return v.String()
	default:
		return ""
	}
}

func payloadFromJS(v js.Value) ([]byte, error) {
	switch v.Type() {
	case js.TypeString:
		return []byte(v.String()), nil
	case js.TypeObject:
		return bytesFromJS(v)
	default:
		return nil, fmt.Errorf("payload must be a JSON string or Uint8Array")
	}
}

func bytesFromJS(v js.Value) ([]byte, error) {
	if v.IsNull() || v.IsUndefined() {
		return nil, fmt.Errorf("expected Uint8Array or ArrayBuffer")
	}
	uint8Array := js.Global().Get("Uint8Array")
	switch {
	case v.InstanceOf(uint8Array):
	case v.InstanceOf(js.Global().Get("ArrayBuffer")):
		v = uint8Array.New(v)
	default:
		if ctor := v.Get("constructor"); ctor.Truthy() && ctor.Get("name").String() == "Uint8Array" {
			break
		}
		if buf := v.Get("buffer"); buf.Truthy() && v.Get("byteLength").Type() == js.TypeNumber {
			v = uint8Array.New(buf, v.Get("byteOffset").Int(), v.Get("byteLength").Int())
			break
		}
		return nil, fmt.Errorf("expected Uint8Array or ArrayBuffer")
	}
	n := v.Get("byteLength").Int()
	out := make([]byte, n)
	if copied := js.CopyBytesToGo(out, v); copied != n {
		return nil, fmt.Errorf("copied %d bytes, want %d", copied, n)
	}
	return out, nil
}

func bytesToJS(data []byte) (js.Value, error) {
	dst := js.Global().Get("Uint8Array").New(len(data))
	if copied := js.CopyBytesToJS(dst, data); copied != len(data) {
		return js.Undefined(), fmt.Errorf("copied %d bytes, want %d", copied, len(data))
	}
	return dst, nil
}
