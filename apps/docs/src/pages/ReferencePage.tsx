import { PAGES } from '../catalog'

const GO = [
  { name: 'xlsxpatch', role: 'Native XLSX extract/apply' },
  { name: 'docxpatch', role: 'Native DOCX extract/apply' },
  { name: 'pptxpatch', role: 'Native PPTX extract/apply' },
  { name: 'officecompat', role: 'OPC preservation and corpus' },
  { name: 'collab', role: 'In-process room hub' },
  { name: 'injoffice-server', role: 'Optional HTTP + SSE sidecar' },
]

export function ReferencePage() {
  return (
    <article className="article">
      <h1>Package reference</h1>
      <p className="lead">TypeScript packages under <code>packages/</code> and Go modules under <code>go/</code>. Guides carry the runnable examples.</p>
      <h2 id="ts">TypeScript</h2>
      <table>
        <thead><tr><th>Package</th><th>Guide</th></tr></thead>
        <tbody>
          {PAGES.filter((page) => page.packageName).map((page) => (
            <tr key={page.id}>
              <td><code>{page.packageName}</code></td>
              <td><a href={page.href}>{page.navTitle}</a> — {page.description}</td>
            </tr>
          ))}
        </tbody>
      </table>
      <h2 id="go">Go</h2>
      <table>
        <thead><tr><th>Module</th><th>Role</th></tr></thead>
        <tbody>
          {GO.map((item) => (
            <tr key={item.name}><td><code>go/{item.name}</code></td><td>{item.role}</td></tr>
          ))}
        </tbody>
      </table>
    </article>
  )
}
