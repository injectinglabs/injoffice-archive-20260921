import assert from 'node:assert/strict'
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import test from 'node:test'
import { extensionlessDeclarationSpecifiers, rewriteDeclarationTree } from './fix-declaration-specifiers.mjs'

test('rewrites declaration-relative module specifiers for NodeNext consumers', () => {
  const root = mkdtempSync(join(tmpdir(), 'injoffice-declarations-'))
  try {
    mkdirSync(join(root, 'nested'))
    writeFileSync(join(root, 'index.d.ts'), [
      "export type { A } from './a'",
      "export { value } from './nested'",
      "export { generated } from './schema.generated'",
      "export type Dynamic = import('./a').A",
      "export type { External } from 'external-package'",
      "export type { Existing } from './existing.js'",
      '//# sourceMappingURL=index.d.ts.map',
      '',
    ].join('\n'))
    writeFileSync(join(root, 'index.d.ts.map'), '{}\n')
    writeFileSync(join(root, 'a.d.ts'), 'export interface A { value: string }\n')
    writeFileSync(join(root, 'nested', 'index.d.ts'), 'export declare const value = 1\n')
    writeFileSync(join(root, 'schema.generated.d.ts'), 'export declare const generated = 1\n')
    writeFileSync(join(root, 'existing.d.ts'), 'export interface Existing {}\n')

    assert.deepEqual(rewriteDeclarationTree(root), { files: 1, specifiers: 4 })
    assert.equal(readFileSync(join(root, 'index.d.ts'), 'utf8'), [
      "export type { A } from './a.js'",
      "export { value } from './nested/index.js'",
      "export { generated } from './schema.generated.js'",
      "export type Dynamic = import('./a.js').A",
      "export type { External } from 'external-package'",
      "export type { Existing } from './existing.js'",
      '',
    ].join('\n'))
    assert.equal(existsSync(join(root, 'index.d.ts.map')), false)
    assert.deepEqual(extensionlessDeclarationSpecifiers(root), [])
    assert.deepEqual(rewriteDeclarationTree(root), { files: 0, specifiers: 0 })
  } finally {
    rmSync(root, { recursive: true, force: true })
  }
})
