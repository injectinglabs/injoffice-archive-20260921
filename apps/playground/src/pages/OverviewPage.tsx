import { TOOL_WORKSPACES, workspaceHref } from '../toolWorkspaces'
import { preloadWorkspaceOnIntent } from '../workspaceRegistry'

const SUMMARIES = {
  sheets: 'Edit cells, visualize data, build pivots, explore formulas and connectors, and verify native XLSX edits.',
  docs: 'Edit a real DOCX, review precise text changes, compare versions, and explore collaborative writing.',
  slides: 'Create a deck, edit native PPTX text and shapes, and inspect authoring, rendering, and layout.',
  pdf: 'Read, annotate, fill forms, organize pages, and download edited PDFs. Explore advanced host operations when configured.',
} as const
const FILES = { sheets: 'XLSX', docs: 'DOCX', slides: 'PPTX', pdf: 'PDF' } as const

export default function OverviewPage({ sidecar }: { sidecar: 'checking' | 'connected' | 'offline' }) {
  return <div className="overview-page">
    <section className="task-launch four-tool-launch" aria-labelledby="showcase-title">
      <header className="task-launch__intro">
        <h1 id="showcase-title" tabIndex={-1}>Explore the four document tools</h1>
        <p>One workspace for each tool. Start with the editor, then explore its AI workflow, collaboration, and advanced capabilities without leaving the example.</p>
      </header>
      <ul className="task-launch__tasks" aria-label="Choose a document tool">
        {TOOL_WORKSPACES.map(workspace => <li key={workspace.tool}>
          <a className="tool-example task-launch__task" data-workspace-entry={workspace.tool}
            href={workspaceHref(workspace.tool)}
            onPointerEnter={() => preloadWorkspaceOnIntent(workspace.tool)}
            onFocus={() => preloadWorkspaceOnIntent(workspace.tool)}>
            <span className="task-launch__format">{FILES[workspace.tool]}</span>
            <span className="task-launch__task-copy"><strong>{workspace.title}</strong><span>{SUMMARIES[workspace.tool]}</span><small>Editor, mock AI, collaboration, and more</small></span>
            <span className="task-launch__open">Open {workspace.title}</span>
          </a>
        </li>)}
      </ul>
      <div className="four-tool-launch__notes">
        <p><strong>Simulated agent · real document operations</strong> AI proposals use a deterministic mock. No LLM, API key, or external AI service is used.</p>
        <p>Advanced views may use independent samples rather than the editor’s file. File operations, package-contract simulations, and host-only capabilities are labelled in each workspace.</p>
        <p>Scroll between the four tools; the navigation follows your position. Visited views retain your edits until you reset or close that workspace.</p>
        <p className="four-tool-launch__runtime" role="status" aria-live="polite">{sidecar === 'connected' ? 'Optional local server connected.' : sidecar === 'checking' ? 'Checking the optional local server…' : 'Browser tools are available. Optional server features need local setup.'}</p>
      </div>
    </section>
  </div>
}
