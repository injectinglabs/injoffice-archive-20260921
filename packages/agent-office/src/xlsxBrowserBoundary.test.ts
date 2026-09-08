import { readFileSync } from 'node:fs'
import { dirname, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

/** Follow runtime imports, not just the adapter's first hop: Sheets' root includes Node-only font code. */
describe('public XLSX adapter browser dependency boundary', () => {
  it('reaches both native validators without a Node builtin, font engine, or broad Sheets entry', () => {
    const packageRoot = fileURLToPath(new URL('../', import.meta.url))
    const aliases: Record<string, string> = {
      '@injoffice/sheets/browser': resolve(packageRoot, '../sheets/src/browser.ts'),
      '@injoffice/agent-tools': resolve(packageRoot, '../agent-tools/src/index.ts'),
    }
    const visited = new Set<string>()
    const visit = (path: string) => {
      if (visited.has(path)) return
      visited.add(path)
      const source = readFileSync(path, 'utf8').replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '')
      const dependency = (specifier: string) => {
        expect(specifier, `${path} must keep its runtime graph browser-safe`).not.toMatch(/^(?:node:|@injoffice\/font-metrics(?:\/|$)|@injoffice\/sheets$)/)
        if (specifier.startsWith('.')) {
          const target = resolve(dirname(path), /\.(?:js|ts|json)$/.test(specifier) ? specifier.replace(/\.js$/, '.ts') : `${specifier}.ts`)
          if (!target.endsWith('.json')) visit(target)
        } else {
          expect(aliases[specifier], `Unexpected browser runtime dependency ${specifier} in ${path}`).toBeDefined()
          visit(aliases[specifier])
        }
      }
      // Follow literal dynamic/type imports too; fail closed on opaque or CommonJS loading.
      expect(/\brequire\s*\(|\bimport\s*\(\s*[^\s'"]/.test(source), `${path} must not introduce opaque runtime imports`).toBe(false)
      for (const match of source.matchAll(/\b(?:import|export)\s+(?!type\b)(\{[^}]*\}|\*(?:\s+as\s+\w+)?|\w+)\s+from\s+['"]([^'"]+)['"]/g)) {
        const clause = match[1]
        const onlyTypes = clause.startsWith('{') && clause.slice(1, -1).split(',').map((item) => item.trim()).filter(Boolean).every((item) => item.startsWith('type '))
        if (!onlyTypes) dependency(match[2])
      }
      for (const match of source.matchAll(/\bimport\s+['"]([^'"]+)['"]/g)) dependency(match[1])
      for (const match of source.matchAll(/\bimport\s*\(\s*['"]([^'"]+)['"]/g)) dependency(match[1])
    }
    visit(resolve(packageRoot, 'src/xlsx.ts'))
    expect([...visited]).toContain(resolve(packageRoot, '../sheets/src/nativeValidation.ts'))
    expect([...visited]).toContain(resolve(packageRoot, '../sheets/src/nativeValidationV2.ts'))
  })
})
