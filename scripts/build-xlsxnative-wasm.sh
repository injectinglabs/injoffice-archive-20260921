#!/usr/bin/env bash
# Build the xlsxnative browser engine (syscall/js, standard Go).
# The @injoffice/xlsx-wasm package supplies the playground's default local
# browser engine. The HTTP server remains an explicit fallback.
set -euo pipefail

root=$(cd "$(dirname "$0")/.." && pwd)
src="$root/go/xlsxpatch/cmd/xlsxnativewasm"
out="$src/dist"
package_out=${XLSXNATIVE_WASM_PACKAGE_OUT:-}
package_worker="$root/packages/xlsx-wasm/worker/xlsxnative.worker.js"
goroot=$(go env GOROOT)
wasm_exec="$goroot/misc/wasm/wasm_exec.js"
if [[ ! -f $wasm_exec ]]; then
  wasm_exec="$goroot/lib/wasm/wasm_exec.js"
fi
if [[ ! -f $wasm_exec ]]; then
  echo "wasm_exec.js not found under $goroot" >&2
  exit 1
fi

mkdir -p "$out"
echo "building GOOS=js GOARCH=wasm -> $out/xlsxnative.wasm" >&2
(
  cd "$root/go/xlsxpatch"
  GOOS=js GOARCH=wasm go build -trimpath -ldflags='-s -w' -o "$out/xlsxnative.wasm" ./cmd/xlsxnativewasm
)
cp "$wasm_exec" "$out/wasm_exec.js"
if [[ -n $package_out ]]; then
  if [[ ! -f $package_worker ]]; then
    echo "XLSX WASM package worker not found at $package_worker" >&2
    exit 1
  fi
  mkdir -p "$package_out"
  cp "$out/xlsxnative.wasm" "$out/wasm_exec.js" "$package_worker" "$package_out/"
fi

size=$(wc -c < "$out/xlsxnative.wasm" | tr -d ' ')
echo "xlsxnative.wasm ${size} bytes" >&2
echo "copied wasm_exec.js from $wasm_exec" >&2
if [[ -n $package_out ]]; then
  echo "npm package artifacts: $package_out" >&2
fi
echo "contract: node $src/node_contract.mjs extract --input go/xlsxpatch/testdata/excel-authored/happy-tree.xlsx" >&2
