import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { dirname, relative, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'

const site = fileURLToPath(new URL('..', import.meta.url))
const root = resolve(site, '../..')
const inventory = JSON.parse(readFileSync(resolve(root, 'docs/api-reference.json'), 'utf8'))
const repo = 'https://github.com/injectinglabs/injoffice/blob/main/'
function write(path, value) {
  const target = resolve(site, path)
  mkdirSync(dirname(target), { recursive: true })
  writeFileSync(target, value)
}
// Source READMEs stay authoritative. Resolve their relative repository links
// before publishing them at a different URL; do not alter fenced code examples.
function links(markdown, source) {
  let fenced = false
  return markdown.split('\n').map(line => {
    if (/^\s*(`{3,}|~{3,})/.test(line)) { fenced = !fenced; return line }
    if (fenced) return line
    return line.replace(/\]\(([^\s)]+)(\s+"[^"]*")?\)/g, (whole, href, title = '') => {
      if (href.startsWith('#/')) return `](https://injoffice.com/${href}${title})`
      if (/^(https?:|mailto:|data:|#)/.test(href)) return whole
      const [file, hash] = href.split('#')
      const path = relative(root, resolve(dirname(resolve(root, source)), file)).split('/').map(encodeURIComponent).join('/')
      return `](${repo}${path}${hash ? '#' + hash : ''}${title})`
    })
  }).join('\n')
}
const pages = []
function sourcePage(destination, source, suffix = '') {
  const frontmatter = `---\neditLink: false\n---\n\n`
  const body = links(readFileSync(resolve(root, source), 'utf8'), source)
  write(destination, frontmatter + body + `\n\n---\n\nMaintained in [${source}](${repo}${source}). This reference follows the repository's main branch, not a frozen npm release.\n` + suffix)
  pages.push(destination.replace(/\.md$/, '.html'))
}
for (const pkg of inventory.packages) {
  const slug = pkg.name.replace('@injoffice/', '')
  const symbols = pkg.exports.map(entry => {
    const name = entry.subpath === '.' ? pkg.name : pkg.name + entry.subpath.slice(1)
    return `\n### \`${name}\`\n\n` + (entry.symbols.length
      ? entry.symbols.map(symbol => `- \`${symbol.name}\` (${symbol.kind})`).join('\n')
      : 'Asset or declaration entry; consult the source contract.')
  }).join('\n')
  sourcePage(`reference/generated/packages/${slug}.md`, `packages/${slug}/README.md`, '\n## Export inventory\n\nGenerated from the built declaration inventory. This is an index of exported names, not a complete parameter-level API specification. Use your editor’s declarations and the linked contracts for exact fields and overloads.\n' + symbols)
}
for (const module of inventory.go) {
  sourcePage(`reference/generated/go/${module.path.split('/').at(-1)}.md`, `${module.path}/README.md`)
}
const contracts = [
  'AGENT-CHANGESETS', 'AGENT-PROPOSAL-HOST', 'DOCX-NATIVE-CONTRACT',
  'PPTX-NATIVE-CONTRACT', 'PPTX-RENDER-TREE', 'COLLABORATION-PROTOCOL',
  'XLSX-CONNECTOR-EXTENSION', 'UNIVER-SHEETS-COMPATIBILITY', 'FUNCTIONS',
  'DEPENDENCY-TRANSPARENCY', 'PUBLIC-RELEASE',
]
for (const name of contracts) sourcePage(`reference/generated/contracts/${name.toLowerCase()}.md`, `docs/${name}.md`)
write('reference/generated/index.md', '# Source-derived reference\n\nThese pages are generated from maintained repository READMEs and contracts. Edit their linked source files, not the generated copies.\n\n## TypeScript packages\n\n'
  + inventory.packages.map(pkg => `- [${pkg.name}](./packages/${pkg.name.replace('@injoffice/', '')})`).join('\n')
  + '\n\n## Go modules\n\n' + inventory.go.map(module => `- [${module.path}](./go/${module.path.split('/').at(-1)})`).join('\n')
  + '\n\n## Protocols and release references\n\n' + contracts.map(name => `- [${name.toLowerCase().replaceAll('-', ' ')}](./contracts/${name.toLowerCase()})`).join('\n') + '\n')
write('public/logo.svg', readFileSync(resolve(root, 'logo.svg')))
write('public/llms.txt', '# InjOffice documentation\n\nGuides and source-derived references for capability-scoped document workflows. No model service is included.\n\n- [Quickstart](./getting-started/quickstart.html)\n- [Agent safety](./agents/safety.html)\n- [Support and limitations](./getting-started/support.html)\n- [Reference](./reference/index.html)\n')
console.log(`Generated ${pages.length} source-derived reference pages; no demo assets or document engines are bundled.`)
