#!/usr/bin/env bash
set -euo pipefail

command_dir=$(cd "$(dirname "$0")" && pwd)
module_dir=$(cd "$command_dir/../.." && pwd)
output_dir=${1:-"$command_dir/dist"}
go_root=$(go env GOROOT)
wasm_exec="$go_root/misc/wasm/wasm_exec.js"
if [[ ! -f $wasm_exec ]]; then
  wasm_exec="$go_root/lib/wasm/wasm_exec.js"
fi
if [[ ! -f $wasm_exec ]]; then
  echo "wasm_exec.js not found under $go_root" >&2
  exit 1
fi

mkdir -p "$output_dir"
(
  cd "$module_dir"
  GOOS=js GOARCH=wasm go build -trimpath -ldflags='-s -w' -o "$output_dir/docxnative.wasm" ./cmd/docxnativewasm
)
cp "$wasm_exec" "$output_dir/wasm_exec.js"

size=$(wc -c < "$output_dir/docxnative.wasm" | tr -d ' ')
echo "docxnative.wasm $size bytes" >&2
echo "copied wasm_exec.js from $wasm_exec" >&2
