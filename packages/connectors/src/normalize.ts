// Pure normalization: raw source payloads → values grids. Lives engine-side
// (and is mirrored by the gateway fetcher) so the tabulation rules are one
// tested truth, not per-host improvisation.

/** Resolve a dot-path ("data.items") into a JSON value. Empty/undefined path
 *  returns the root. Missing segments resolve to undefined. */
export function resolvePath(root: unknown, path?: string): unknown {
  if (!path) return root
  let cur: unknown = root
  for (const seg of path.split('.')) {
    if (seg === '') continue
    if (cur == null || typeof cur !== 'object') return undefined
    cur = (cur as Record<string, unknown>)[seg]
  }
  return cur
}

function cellValue(v: unknown): string | number | null {
  if (v == null) return null
  if (typeof v === 'number' && Number.isFinite(v)) return v
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE'
  if (typeof v === 'object') return JSON.stringify(v)
  return String(v)
}

/** Tabulate a JSON value:
 *  - array of objects → header = union of keys in first-seen order, one row
 *    per object
 *  - array of arrays  → used as-is
 *  - array of scalars → one column ("value")
 *  - plain object     → two columns (key, value)
 *  Anything else yields an empty grid. */
export function jsonToGrid(data: unknown, path?: string): (string | number | null)[][] {
  const v = resolvePath(data, path)
  if (Array.isArray(v)) {
    if (v.length === 0) return []
    if (Array.isArray(v[0])) {
      return (v as unknown[][]).map((row) => row.map(cellValue))
    }
    if (typeof v[0] === 'object' && v[0] !== null) {
      const keys: string[] = []
      const seen = new Set<string>()
      for (const item of v) {
        if (item && typeof item === 'object' && !Array.isArray(item)) {
          for (const k of Object.keys(item)) {
            if (!seen.has(k)) {
              seen.add(k)
              keys.push(k)
            }
          }
        }
      }
      const rows = v.map((item) =>
        keys.map((k) =>
          item && typeof item === 'object' && !Array.isArray(item)
            ? cellValue((item as Record<string, unknown>)[k])
            : null,
        ),
      )
      return [keys, ...rows]
    }
    return [['value'], ...v.map((x) => [cellValue(x)])]
  }
  if (v && typeof v === 'object') {
    return [['key', 'value'], ...Object.entries(v).map(([k, val]) => [k, cellValue(val)])]
  }
  return []
}

/** Minimal-but-correct CSV parser: quoted fields, doubled-quote escapes,
 *  commas and newlines inside quotes, CRLF/LF endings. Numeric-looking
 *  unquoted fields become numbers. */
export function csvToGrid(text: string): (string | number | null)[][] {
  const rows: (string | number | null)[][] = []
  let row: (string | number | null)[] = []
  let field = ''
  let quoted = false
  let wasQuoted = false

  const pushField = () => {
    if (!wasQuoted && field !== '' && /^-?\d+(\.\d+)?$/.test(field)) {
      row.push(Number(field))
    } else {
      row.push(field === '' && !wasQuoted ? null : field)
    }
    field = ''
    wasQuoted = false
  }
  const pushRow = () => {
    pushField()
    // Trailing blank line artifacts: a lone null row from "\n" at EOF is noise.
    if (!(row.length === 1 && row[0] === null)) rows.push(row)
    row = []
  }

  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          field += '"'
          i++
        } else {
          quoted = false
        }
      } else {
        field += ch
      }
      continue
    }
    switch (ch) {
      case '"':
        if (field === '') {
          quoted = true
          wasQuoted = true
        } else {
          field += ch
        }
        break
      case ',':
        pushField()
        break
      case '\r':
        break
      case '\n':
        pushRow()
        break
      default:
        field += ch
    }
  }
  if (field !== '' || wasQuoted || row.length > 0) pushRow()
  return rows
}
