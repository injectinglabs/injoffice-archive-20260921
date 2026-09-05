// Cell-level spreadsheet diff — the history feature nobody else in this
// market shows. Pure: two value grids in, a structured change set out. The
// rail renders it as "B4: 120 → 300" rows; agents consume the same shape for
// "what changed since v3".
//
// Semantics: positional comparison over the union extent. Deliberately NOT a
// row-alignment/LCS diff — for spreadsheets, "row 7 changed" (because a row
// was inserted above) is the truthful positional answer, and alignment
// heuristics guess wrong often enough to erode trust in the numbers. Summary
// counts make large shape changes legible (rowsAdded/rowsRemoved etc.).

export interface CellChange {
  row: number
  col: number
  /** A1-style address for display ("B4"). */
  address: string
  from: string
  to: string
}

export interface GridDiff {
  changes: CellChange[]
  /** Positive when `to` has more rows/cols than `from`; negative when fewer. */
  rowDelta: number
  colDelta: number
  /** True when the change list was truncated at maxChanges. */
  truncated: boolean
  /** Total changed-cell count BEFORE truncation. */
  changeCount: number
}

export function colName(col: number): string {
  let n = col + 1
  let out = ''
  while (n > 0) {
    const rem = (n - 1) % 26
    out = String.fromCharCode(65 + rem) + out
    n = Math.floor((n - 1) / 26)
  }
  return out
}

export function cellAddress(row: number, col: number): string {
  return `${colName(col)}${row + 1}`
}

function norm(v: unknown): string {
  if (v == null) return ''
  return String(v)
}

export interface GridDiffOptions {
  /** Cap on emitted CellChange entries (default 500). changeCount stays exact. */
  maxChanges?: number
}

export function diffGrids(from: unknown[][], to: unknown[][], opts: GridDiffOptions = {}): GridDiff {
  const maxChanges = opts.maxChanges ?? 500
  const rows = Math.max(from.length, to.length)
  const fromCols = Math.max(0, ...from.map((r) => r.length))
  const toCols = Math.max(0, ...to.map((r) => r.length))
  const cols = Math.max(fromCols, toCols)

  const changes: CellChange[] = []
  let changeCount = 0
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const a = norm(from[r]?.[c])
      const b = norm(to[r]?.[c])
      if (a === b) continue
      changeCount++
      if (changes.length < maxChanges) {
        changes.push({ row: r, col: c, address: cellAddress(r, c), from: a, to: b })
      }
    }
  }
  return {
    changes,
    rowDelta: to.length - from.length,
    colDelta: toCols - fromCols,
    truncated: changeCount > changes.length,
    changeCount,
  }
}
