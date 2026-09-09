# InjOffice showcase

Run `npm run dev` from the repository root (Node 22 or newer). The playground normally starts on port 3100.

The showcase opens directly into Sheets on one continuous page with exactly four live examples: **Sheets, Docs, Slides, and PDF**. There is no introduction or duplicate tool catalogue. Empty and legacy `#/overview` URLs resolve to `#/sheets` without adding history entries. Scroll or use the four-item navigation, whose selection follows your position. Passive scrolling updates the URL without adding history entries or moving keyboard focus; explicit links support Back/Forward.

Each workspace starts with its editor. Grouped tabs collect the existing AI, analysis, collaboration, history, native-file, rendering, and layout views relevant to that tool. Every former capability remains reachable inside these workspaces. Views use independent sample/state models: an edit in one view does not silently update another view’s file. This covers implemented capabilities, not unrestricted Office parity.

Workspaces load near the viewport; internal views load only when opened. Visited views stay mounted but hidden when switching tabs, retaining edits and pending approvals. Inactive spreadsheet canvases retain a measurable, inert layout to avoid zero-width resizing. AI sessions and collaboration rooms are fixed to their workspace’s tool. Reset/close affects only its workspace; refreshing discards all in-memory state. Failed view loading is isolated with an explicit retry and confirmed reload fallback.

Guide & source follows the selected capability and contains a manual checklist, expected evidence, exact repository links, related capabilities, and lazily loaded source. Closing it retains progress for that view; the checklist does not perform or certify engine operations.

Untouched demos release their mounted editors after 30 seconds offscreen. Any interaction retains the editor while you explore; **Close demo** releases it explicitly. Reset and close ask before discarding interacted state and wait for supported document operations to finish. This is in-memory demo state, not persistence across page reloads.

The Documents workspace links selectable text passages to its edit panel, with undo and expandable verification details. PDF supports passage-level markup and search highlights, pointer or keyboard drawing, and undo/redo. PDF text selection covers whole extracted passages; DOCX remains a semantic preview rather than Word pagination.

Canonical links use `#/sheets`, `#/docs`, `#/slides`, or `#/pdf` with an optional `?feature=` selection, such as `#/sheets?feature=charts`, `#/docs?feature=agent`, or `#/slides?feature=pptx-native`. Old `#/charts`, `#/agent?format=docs`, and `#/sheets?view=native` bookmarks still open the matching view. Internal tab changes preserve keyboard focus and reading position; sidebar navigation remembers the last view.

## Capability coverage and boundaries

| Workspace | Included views |
| --- | --- |
| Sheets | Workbook editor; real XLSX round trip; sparklines/print/outlines/exchange package samples; charts, pivots, formulas, connectors, shapes; AI; collaboration; history |
| Docs | Real DOCX editing and verified download; AI; document collaboration; text history/diffs; typography contracts |
| Slides | DeckSpec editor; real native PPTX editing; authored native compilation; rendering commands; shapes; AI; deck collaboration; typography contracts |
| PDF | Viewer, search, markup/drawing, forms, stamps, page operations and undo/redo/download; advanced host tools; AI; annotation collaboration; typography contracts |

Advanced boundaries remain explicit: chart SVG export is a title-only adapter sample; standalone formula jobs use a limited demonstration evaluator; workbook print/exchange are contract samples; history uses a session-memory sample host. DeckSpec authoring and render commands are not saved PPTX bytes—use native PPTX editing for verified file output. Canonical font shaping and advanced PDF operations require their stated host runtime.

## Guided agent demo

Open `#/sheets?feature=agent` (or the legacy `#/agent?format=sheets`) for the real-file AI path. The browser loads the bundled `launch-readiness-plan.xlsx`, uses the public `createXlsxAgentAdapter` from `@injoffice/agent-office/xlsx`, and sends JSON calls through `createAgentToolDispatcher` from `@injoffice/agent-tools`. The browser host owns the native Worker, original and replacement bytes, approval, compare-and-swap checks, and exact-output readback.

The demo is labelled **Simulated agent · real document operations**. Guided controls let you choose a disclosed workstream and status, edit a document or presentation title, or select a PDF page and rotation. A deterministic mock turns those choices into one bounded proposal; no model, API key, or endpoint configuration is needed. Select **Run agent**, review the actual before/after change, explicitly approve the exact plan, then commit and download the verified result. **Technical details** and its nested **Try the safety boundaries** section are collapsed until requested. The proposal request/response remains separate from actual Office tool calls, and unsupported requests are refused rather than silently interpreted as the default edit.

In development, the browser calls the bundled `POST /api/agent/mock-propose` route on the same local server. Static builds simulate the response in the browser using the same generator, with no API call. Neither path contacts a provider, forwards an API token, or grants approval—even if live-provider environment variables are set. The demo has no live-provider or local-rule mode switch. The [standalone proposal host](../../docs/AGENT-PROPOSAL-HOST.md) is a separate integrator example, not a way to enable a model in this UI.

Task fields are the default input. In **Technical details**, the generated request is read-only until **Use an advanced request instead of task fields** is enabled. Advanced mode replaces the task fields with the exact active request; switching back restores the sample's guided defaults. Editing either input mode clears the previous preview and approval.

Only the proposal is mocked. All four agent demos edit and verify real files:

| Format | Demonstrated edit | Output verification |
| --- | --- | --- |
| XLSX | One discovered workstream status cell | Native browser Worker re-extracts saved bytes |
| DOCX | One exact, source-anchored text run | Native browser Worker re-extracts saved bytes |
| PPTX | One exact native text element | Native browser Worker re-extracts saved bytes |
| PDF | One page rotation | Fresh PDF parse checks rotation, page count, geometry, and SHA-256 |

The title tasks inspect the bundled Northstar launch brief and launch-review deck; the PDF task starts with a four-page operating review. Each uses a populated fixture, an isolated preview, exact-plan approval, verified download, and the same stale-source/retry/fault scenarios. Text and page-metadata previews are bounded projections, not full-fidelity Office renderers. These examples do not imply arbitrary edits or uploaded-file support.

The expandable tool trace shows actual dispatcher arguments and results, including `office.read`. Approval is host-owned and bound to the exact reviewed change set; a model/tool argument containing `confirmation: "approved"` is not approval. After approval, `office.commit` performs the native write and returns committed verification. `office.verify` is a planned-result check, not a separate post-commit reopen call.

Safety controls deliberately demonstrate three distinct outcomes:

- Change the source after planning, then approve: the stale change set is rejected without another native write or verified download.
- Retry a verified commit: the same idempotency key returns the same result without duplicating the write. The displayed native-write counter makes this visible.
- Inject a verification failure before committing: the UI reports **Write completed; verification failed** and does not offer the result as a verified download. This is a test injection, not a claim that the original bytes were restored.

Reload sample discards local session state and reloads the bundled source. For production, replace the local host state with authenticated authorization, atomic persistent revision checks, and durable idempotency. Keep model access to planning tools separate from the trusted approval action; browser UI state alone is not a server authorization boundary.

## Verification

- `npm run test -w apps-playground`
- `npm run typecheck -w apps-playground`
- `npm run test:office-browser-smoke:built` after building workspace packages
- `SHOWCASE_URL=http://127.0.0.1:3100/ npm run test:showcase-browser` with the playground running
- `npm run test:showcase-browser -- --built` after the Office browser smoke build above; starts and stops its own ephemeral server at `/injoffice-smoke/`, without touching an existing demo server
- `npm run test:showcase-browser -- --dev` after building workspace packages; starts and stops an isolated development server on an ephemeral IPv4 port, omitting the playground's fixed-port IPv6 listener so neither side of port 3100 is touched
- `npm run test:scroll-showcase-browser -- --built` (or `--dev`) checks continuous scrolling, lazy loading, sidebar tracking, history, focus, retained edits/approvals, and mobile navigation
- `node scripts/smoke-demo-ux-browser.mjs --built` checks PDF placement, markup, search, keyboard navigation, undo/redo, and invalid-upload recovery; build with the `/injoffice-smoke/` base first, or omit `--built` to use an isolated development server

Development smoke configures its own isolated loopback-only upstream stub and verifies that the guided demo sends it zero requests. It temporarily replaces any configured proposal URL and removes any configured token, then restores both environment values during cleanup without printing them. The test demonstrates that configuring a host cannot silently turn the demo into a live-model client. Built and existing-server smoke also never trigger live-provider requests. Separate host tests cover the standalone bridge's consent and response-envelope guards.

The showcase smoke runs in CI against built subpath and development deployments. It checks direct Sheets startup, legacy overview redirects, four workspaces/four navigation items, all internal capabilities, isolated lazy loading/failures, drawer focus and guides, retained view edits and approvals, native file round trips, all four mock-agent safety workflows and verified downloads, scrolling/history, responsive layouts, and both themes. It fails on unexpected console/uncaught errors and prints desktop/mobile screenshots in a temporary directory; CI retains failure artifacts. Set `CHROME_BIN` to override Chrome discovery or `SHOWCASE_OUTPUT` for diagnostics.
