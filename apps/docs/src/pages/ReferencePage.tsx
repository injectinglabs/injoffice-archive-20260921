import { useMemo, useState } from 'react'
import api from '../generated/api-reference.json'

type SymbolInfo = { name: string; kind: string; signature: string }
type ExportEntry = { subpath: string; file: string; symbols: SymbolInfo[] }
type PackageEntry = { name: string; description: string; exports: ExportEntry[] }

const PACKAGES = api.packages as PackageEntry[]
const GO = api.go as { path: string; module: string; summary: string }[]

export function ReferencePage() {
  const [query, setQuery] = useState('')
  const needle = query.trim().toLowerCase()
  const packages = useMemo(() => {
    if (!needle) return PACKAGES
    return PACKAGES.map((pkg) => {
      const nameHit = pkg.name.toLowerCase().includes(needle) || pkg.description.toLowerCase().includes(needle)
      if (nameHit) return pkg
      const exports = pkg.exports.map((entry) => ({
        ...entry,
        symbols: entry.symbols.filter((symbol) =>
          symbol.name.toLowerCase().includes(needle) || symbol.signature.toLowerCase().includes(needle)),
      })).filter((entry) => entry.symbols.length > 0)
      return { ...pkg, exports }
    }).filter((pkg) => pkg.exports.some((entry) => entry.symbols.length > 0))
  }, [needle])

  return (
    <article className="article">
      <h1>API reference</h1>
      <p className="lead">Generated from each package’s published <code>.d.ts</code> exports. Guides still carry the runnable examples.</p>
      <label className="field" style={{ maxWidth: 360, marginBottom: 20 }}>
        Filter
        <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="decodeWorkbookMutationBatch" />
      </label>
      <h2 id="ts">TypeScript packages</h2>
      {packages.map((pkg) => (
        <section key={pkg.name} className="api-package">
          <h3 id={pkg.name.replace('@', '').replace('/', '-')}><code>{pkg.name}</code></h3>
          <p>{pkg.description}</p>
          {pkg.exports.map((entry) => (
            <div key={entry.subpath}>
              <p><code>{pkg.name}{entry.subpath === '.' ? '' : entry.subpath.slice(1)}</code> · {entry.symbols.length} export{entry.symbols.length === 1 ? '' : 's'}</p>
              <table>
                <thead><tr><th>Name</th><th>Kind</th><th>Signature</th></tr></thead>
                <tbody>
                  {entry.symbols.map((symbol) => (
                    <tr key={`${entry.subpath}:${symbol.name}`}>
                      <td><code>{symbol.name}</code></td>
                      <td>{symbol.kind}</td>
                      <td><code>{symbol.signature}</code></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))}
        </section>
      ))}
      <h2 id="go">Go modules</h2>
      <table>
        <thead><tr><th>Path</th><th>Module</th><th>Summary</th></tr></thead>
        <tbody>
          {GO.map((item) => (
            <tr key={item.path}>
              <td><code>{item.path}</code></td>
              <td><code>{item.module}</code></td>
              <td>{item.summary}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </article>
  )
}
