//go:build js && wasm

// Command docxnativewasm exposes docxpatch's native v1 extract/apply boundary
// to JavaScript without introducing a second DOCX reader or writer.
package main

import (
	"fmt"
	"syscall/js"

	"github.com/injectinglabs/injoffice/go/docxpatch"
)

func main() {
	obj := js.Global().Get("Object").New()
	obj.Set("extract", guarded(jsExtract))
	obj.Set("inspect", guarded(jsInspect))
	obj.Set("apply", guarded(jsApply))
	js.Global().Set("docxnative", obj)
	if ready := js.Global().Get("docxnativeOnReady"); ready.Type() == js.TypeFunction {
		ready.Invoke()
	}
	select {}
}

// guarded recovers inside the callback so a refused or malformed request does
// not escape syscall/js.handleEvent and terminate the reusable WASM instance.
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
	obj := js.Global().Get("Object").New()
	obj.Set("ok", true)
	obj.Set("value", value)
	return obj
}

func fail(message string) js.Value {
	return failure(message, false)
}

// failFatal distinguishes a recovered panic from an expected validation or
// CAS refusal. Consumers must discard the worker after a fatal result.
func failFatal(message string) js.Value {
	return failure(message, true)
}

func failure(message string, fatal bool) js.Value {
	if message == "" {
		message = "docxnative failed"
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
	document, err := docxpatch.ExtractNativeDocumentV1(data)
	if err != nil {
		return fail(err.Error())
	}
	encoded, err := docxpatch.EncodeNativeDocumentV1(document)
	if err != nil {
		return fail(err.Error())
	}
	return ok(string(encoded))
}

// inspect(bytes) -> same-source read-only document and resolved layout JSON.
func jsInspect(_ js.Value, args []js.Value) any {
	if len(args) != 1 {
		return fail("inspect(bytes) requires 1 argument")
	}
	data, err := bytesFromJS(args[0])
	if err != nil {
		return fail(err.Error())
	}
	encoded, err := docxpatch.InspectNativePartialSourceV1(data)
	if err != nil {
		return fail(err.Error())
	}
	return ok(string(encoded))
}

// apply(original, payload, expectedRevision) returns mutated DOCX bytes.
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
	result, err := docxpatch.ApplyNativeTextMutationPayloadV1(original, payload, expectedRevision)
	if err != nil {
		return fail(err.Error())
	}
	packed, err := bytesToJS(result.Package)
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

	length := value.Get("byteLength").Int()
	result := make([]byte, length)
	if copied := js.CopyBytesToGo(result, value); copied != length {
		return nil, fmt.Errorf("copied %d bytes, want %d", copied, length)
	}
	return result, nil
}

func bytesToJS(data []byte) (js.Value, error) {
	destination := js.Global().Get("Uint8Array").New(len(data))
	if copied := js.CopyBytesToJS(destination, data); copied != len(data) {
		return js.Undefined(), fmt.Errorf("copied %d bytes, want %d", copied, len(data))
	}
	return destination, nil
}
