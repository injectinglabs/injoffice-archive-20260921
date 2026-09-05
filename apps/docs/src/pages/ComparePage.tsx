import { Article } from '../components/Article'
import { Callout } from '../components/Callout'

export function ComparePage() {
  return (
    <Article id="compare">
      <h2 id="roles">Three different demos</h2>
      <table>
        <thead><tr><th></th><th>InjOffice</th><th>Univer OSS</th><th>Univer Pro</th></tr></thead>
        <tbody>
          <tr><td>Subject</td><td>Native file engines + optional grid</td><td>Apache editor snapshot</td><td>Licensed editor + Universer</td></tr>
          <tr><td>File truth</td><td>Original OOXML bytes</td><td>In-memory snapshot</td><td>Server unitId</td></tr>
          <tr><td>License</td><td>Apache-2.0</td><td>Apache-2.0</td><td>Commercial <code>@univerjs-pro/*</code></td></tr>
          <tr><td>Docs style</td><td>This site</td><td>Guides + showcase</td><td>Same docs, Pro-tagged pages</td></tr>
        </tbody>
      </table>
      <h2 id="reuse">What InjOffice reuses</h2>
      <p>The live Sheets editor is Univer 0.25.1 OSS via <code>createOssUniver</code>. Core grid, formulas, numfmt, CF, DV, filter, sort, notes, tables, comments, hyperlinks, find/replace, and images come from public presets.</p>
      <h2 id="pro">What Univer documents as Pro</h2>
      <p>Univer’s feature pages for charts, pivots, print, shapes, sparklines, outline, import/export, collaboration, and edit history install <code>@univerjs/preset-sheets-advanced</code> or <code>@univerjs-pro/*</code>, take a client license, and often require Universer.</p>
      <table>
        <thead><tr><th>Pro category</th><th>InjOffice status</th><th>Demo / API</th></tr></thead>
        <tbody>
          <tr><td>Charts</td><td>Independent, partial</td><td><a href="#/guides/charts">@injoffice/charts</a></td></tr>
          <tr><td>Pivots</td><td>Independent, partial</td><td><a href="#/guides/pivots">@injoffice/pivots</a></td></tr>
          <tr><td>Shapes</td><td>Independent, partial</td><td><a href="#/guides/shapes">@injoffice/shapes</a></td></tr>
          <tr><td>Import/export</td><td>Fail-closed native extract/apply</td><td><a href="#/guides/xlsx">Native XLSX</a></td></tr>
          <tr><td>Collaboration</td><td>Presence + ordered ops</td><td><a href="#/guides/collaboration">@injoffice/collab</a></td></tr>
          <tr><td>Edit history</td><td>Structured diffs</td><td><a href="#/guides/history">@injoffice/history</a></td></tr>
          <tr><td>Print, sparklines, outlines</td><td>Absent</td><td>—</td></tr>
          <tr><td>Pro formula engine / server calc</td><td>OSS engine audit only</td><td><a href="#/guides/formulas">@injoffice/formulas</a></td></tr>
        </tbody>
      </table>
      <Callout kind="caution">Do not depend on <code>@univerjs/presets</code> or <code>preset-sheets-advanced</code>. Those packages re-export commercial modules even though the preset manifest says Apache-2.0. That is why InjOffice bootstraps Univer with <code>createOssUniver</code>.</Callout>
      <h2 id="univer-docs">How Univer writes a feature page</h2>
      <ol>
        <li>Package badge, locale/CSS, Pro tag</li>
        <li>Live preview</li>
        <li>Preset vs plugin install tabs</li>
        <li>Facade API examples</li>
        <li>Limits (watermark, quotas) for unlicensed Pro</li>
      </ol>
      <p>InjOffice pages follow that order with different badges: package, runtime, and file authority. Limits are fail-closed refusals, not license watermarks.</p>
    </Article>
  )
}
