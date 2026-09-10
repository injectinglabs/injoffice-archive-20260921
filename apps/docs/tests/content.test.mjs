import assert from 'node:assert/strict'
import { test } from 'node:test'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const site = fileURLToPath(new URL('..', import.meta.url))
const root = resolve(site, '../..')
function files(directory) {
  return readdirSync(directory, { withFileTypes: true }).flatMap(entry => {
    if (entry.name.startsWith('.') || entry.name === 'node_modules') return []
    const path = join(directory, entry.name)
    return entry.isDirectory() ? files(path) : [path]
  })
}
test('every public package and Go module has a source-derived reference', () => {
  const inventory = JSON.parse(readFileSync(resolve(root, 'docs/api-reference.json'), 'utf8'))
  for (const pkg of inventory.packages) {
    const page = readFileSync(join(site, 'reference/generated/packages', pkg.name.replace('@injoffice/', '') + '.md'), 'utf8')
    assert.ok(page.includes('## Export inventory'))
    for (const entry of pkg.exports) for (const symbol of entry.symbols) assert.ok(page.includes('`' + symbol.name + '`'), symbol.name)
  }
  for (const module of inventory.go) assert.ok(existsSync(join(site, 'reference/generated', module.path + '.md')))
})
test('authored guide page links and imported TypeScript samples exist', () => {
  const pages = files(site).filter(file => file.endsWith('.md') && !file.includes('/generated/') && !file.endsWith('/README.md'))
  for (const page of pages) {
    const text = readFileSync(page, 'utf8').replace(/```[\s\S]*?```/g, '')
    for (const [, href] of text.matchAll(/\]\(([^\s)]+)\)/g)) {
      if (/^(https?:|mailto:|#)/.test(href)) continue
      const path = resolve(dirname(page), href.split('#')[0])
      assert.ok([path, path + '.md', join(path, 'index.md')].some(existsSync), `${page}: ${href}`)
    }
    for (const [, path] of text.matchAll(/<<< @\/(\S+)/g)) {
      assert.ok(existsSync(resolve(site, path)), `${page}: ${path}`)
      assert.ok(path.endsWith('.ts'), 'guide samples must have a checked TypeScript source')
    }
  }
  assert.ok(pages.length >= 20)
})
test('docs reference corrects obsolete mock-only file-execution claims', () => {
  const page = readFileSync(join(site, 'reference/generated/packages/agent-office.md'), 'utf8')
  assert.ok(!page.includes('other three playground agent formats are lifecycle simulations'))
  assert.ok(page.includes('downloaded bytes are real'))
})
test('the built docs are standalone and contain searchable article HTML', () => {
  const dist = join(site, '.vitepress/dist')
  const output = files(dist)
  assert.ok(output.filter(file => file.endsWith('.html')).length >= 60)
  assert.ok(!output.some(file => /\.(wasm|xlsx|docx|pptx|pdf)$/.test(file)))
  assert.match(readFileSync(join(dist, 'agents/quickstart.html'), 'utf8'), /preparePdfRotation/)
  assert.ok(output.some(file => /localSearchIndex/.test(file)), 'local search index exists')
  assert.ok(existsSync(join(dist, '404.html')))
})
