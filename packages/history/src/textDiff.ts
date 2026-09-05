// Line-level text diff for document artifacts (markdown first). Classic LCS
// over lines — documents, unlike grids, DO want alignment: an inserted
// paragraph shouldn't mark everything after it as changed.
//
// Output is a flat op list the rail renders directly (context/add/remove
// rows). Pure and dependency-free.

export interface TextDiffOp {
  kind: 'same' | 'add' | 'remove'
  line: string
}

export interface TextDiff {
  ops: TextDiffOp[]
  added: number
  removed: number
}

export function diffText(from: string, to: string): TextDiff {
  const a = from.split('\n')
  const b = to.split('\n')

  // LCS table (a.length+1 × b.length+1). Guard against pathological sizes:
  // beyond ~2000×2000 fall back to whole-document replace semantics.
  if (a.length * b.length > 4_000_000) {
    return {
      ops: [
        ...a.map((line) => ({ kind: 'remove' as const, line })),
        ...b.map((line) => ({ kind: 'add' as const, line })),
      ],
      added: b.length,
      removed: a.length,
    }
  }

  const m = a.length
  const n = b.length
  const lcs: Uint32Array = new Uint32Array((m + 1) * (n + 1))
  const at = (i: number, j: number) => i * (n + 1) + j
  for (let i = m - 1; i >= 0; i--) {
    for (let j = n - 1; j >= 0; j--) {
      lcs[at(i, j)] = a[i] === b[j] ? lcs[at(i + 1, j + 1)] + 1 : Math.max(lcs[at(i + 1, j)], lcs[at(i, j + 1)])
    }
  }

  const ops: TextDiffOp[] = []
  let added = 0
  let removed = 0
  let i = 0
  let j = 0
  while (i < m && j < n) {
    if (a[i] === b[j]) {
      ops.push({ kind: 'same', line: a[i] })
      i++
      j++
    } else if (lcs[at(i + 1, j)] >= lcs[at(i, j + 1)]) {
      ops.push({ kind: 'remove', line: a[i] })
      removed++
      i++
    } else {
      ops.push({ kind: 'add', line: b[j] })
      added++
      j++
    }
  }
  for (; i < m; i++) {
    ops.push({ kind: 'remove', line: a[i] })
    removed++
  }
  for (; j < n; j++) {
    ops.push({ kind: 'add', line: b[j] })
    added++
  }
  return { ops, added, removed }
}
