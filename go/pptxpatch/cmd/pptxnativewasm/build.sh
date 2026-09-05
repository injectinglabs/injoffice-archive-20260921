#!/usr/bin/env bash
# Build the browser-local pptxnative syscall/js module with standard Go WASM.
set -euo pipefail

src=$(cd "$(dirname "$0")" && pwd)
module=$(cd "$src/../.." && pwd)
out=${PPTXNATIVE_WASM_OUT:-"$src/dist"}
package_worker="$module/../../packages/pptx-wasm/worker/pptxnative.worker.js"
go_root=$(go env GOROOT)
wasm_exec="$go_root/misc/wasm/wasm_exec.js"
if [[ ! -f $wasm_exec ]]; then
  wasm_exec="$go_root/lib/wasm/wasm_exec.js"
fi
if [[ ! -f $wasm_exec ]]; then
  echo "wasm_exec.js not found under $go_root" >&2
  exit 1
fi

mkdir -p "$out"
echo "building GOOS=js GOARCH=wasm -> $out/pptxnative.wasm" >&2
(
  cd "$module"
  GOOS=js GOARCH=wasm go build -trimpath -ldflags='-s -w' -o "$out/pptxnative.wasm" ./cmd/pptxnativewasm
)
cp "$wasm_exec" "$out/wasm_exec.js"
if [[ -f $package_worker ]]; then
  cp "$package_worker" "$out/pptxnative.worker.js"
fi

size=$(wc -c < "$out/pptxnative.wasm" | tr -d ' ')
echo "pptxnative.wasm ${size} bytes" >&2
echo "copied wasm_exec.js from $wasm_exec" >&2
if [[ -f $package_worker ]]; then
  echo "copied npm worker from $package_worker" >&2
fi
echo "contract: go test ./cmd/pptxnativewasm" >&2
