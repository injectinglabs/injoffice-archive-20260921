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
# +22,775 bytes into 14,380 bytes of CI headroom).
# A further 32 KiB covers table-style border projection, the tblLook inert
# waiver and multi-column band geometry (#353-#360), which together left only
# 584 bytes of CI headroom and forced one PR to be rewritten to fit.
# A further 32 KiB covers the run-typography and spacing wave: the Word 2010
# typography extensions this tier records and does not apply (#369, measured
# +17,485 bytes locally), the lowerRoman note numFmt, the continuous margin band
# and the raised font-manifest face cap (#367, #368, #370). Together those left
# main at 7,005,596 bytes locally and 7,010,988 in CI - 1,364 bytes of headroom,
# so the next docxpatch change of any size failed the gate no matter what it
# contained. Measured: a 2,693-byte local addition came back as 7,014,013 in CI,
# 1,661 bytes over.
#
# The CI toolchain builds this module larger than a local darwin/arm64 build,
# so the local number is not the one this gate sees. Measured CI-local on one
# commit: +5,401 bytes; measured twice more on 2026-09-17, +5,392 bytes (main,
# local 7,005,596 / CI 7,010,988) and +5,724 bytes (local 7,008,289 / CI
# 7,014,013). An earlier note in this file put that gap at ~28 KiB;
# that figure was never reproduced and is wrong. Measure, do not assume.
max_size=$((13 * 1024 * 1024 / 2 + 224 * 1024))
if (( size > max_size )); then
  echo "docxnative.wasm $size bytes exceeds the 6.5 MiB + 224 KiB size ceiling ($max_size bytes)" >&2
  exit 1
fi
echo "docxnative.wasm $size bytes" >&2
echo "copied wasm_exec.js from $wasm_exec" >&2
