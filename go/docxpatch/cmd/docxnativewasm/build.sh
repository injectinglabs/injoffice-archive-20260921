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
# Includes the same-byte style resolver used by partial text/equation inspection.
# Enforce the same bounded budget in local package builds and CI.
# An additional 32 KiB covers source-bound textbox axes, alignment, stacking,
# and inline textbox source evidence. A further 32 KiB covers the package
# tolerance and indent modelling added for coverage: reading a dangling
# relationship and an unstored content-type Override as an absent part, an
# absent or duplicate font-table entry, character-unit indents, and tracked row
# moves (#272-#279, measured +17,610 bytes into 13,721 bytes of headroom).
# A further 32 KiB covers the East-Asian theme font slot: the fontScheme <a:ea>
# and script-table read, settings' w:themeFontLang and its bounded BCP-47 to
# script map, and per-slot deferral of the script run properties (measured
# +22,775 bytes into 14,380 bytes of CI headroom; note the CI toolchain builds
# this module about 28 KiB larger than a local darwin/arm64 build, so the local
# number is not the one this gate sees).
max_size=$((13 * 1024 * 1024 / 2 + 160 * 1024))
if (( size > max_size )); then
  echo "docxnative.wasm $size bytes exceeds the 6.5 MiB + 160 KiB size ceiling ($max_size bytes)" >&2
  exit 1
fi
echo "docxnative.wasm $size bytes" >&2
echo "copied wasm_exec.js from $wasm_exec" >&2
