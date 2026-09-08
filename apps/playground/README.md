# InjOffice showcase

Run `npm run dev` from the repository root (Node 22 or newer). The playground normally starts on port 3100.

The showcase is one continuous page: the overview and nineteen live examples follow the same order as the sidebar. Scroll to explore, or use the sidebar and searchable catalogue to jump to a section. The highlighted sidebar link follows your position. Passive scrolling updates the shareable URL without adding history entries or moving keyboard focus; explicit links support Back/Forward and direct deep links.

The overview starts with four guided document tasks: update a workstream status, revise a document title, update a presentation title, or rotate a PDF page. Each opens its existing agent section with a populated sample file. A task link never applies or approves a change. The searchable engine catalogue remains below the task starters.

Each live demo loads near the viewport and stays mounted after loading, so scrolling away preserves its edits and pending approvals. The four AI formats have independent sessions. AI workflows use the document's vertical scroll; canvas editors and long code traces retain their own necessary scrolling. Reset affects only its section; refreshing the page discards local demo state. A failed lazy download is isolated to its section, with retry and an explicitly confirmed reload fallback for browser-cached failures.

Every example has a Guide & source drawer containing a manual checklist, expected evidence, exact repository links, related examples, and lazily loaded demo source. Closing the drawer retains checklist progress on the current surface; the checklist does not perform or certify engine operations.

Untouched demos release their mounted editors after 30 seconds offscreen. Any interaction retains the editor while you explore; **Close demo** releases it explicitly. Reset and close ask before discarding interacted state and wait for supported document operations to finish. This is in-memory demo state, not persistence across page reloads.

The Documents workspace links selectable text passages to its edit panel, with undo and expandable verification details. PDF supports passage-level markup and search highlights, pointer or keyboard drawing, and undo/redo. PDF text selection covers whole extracted passages; DOCX remains a semantic preview rather than Word pagination.

The spreadsheet modes have shareable links: `#/sheets?view=editor`, `#/sheets?view=native`, and `#/sheets?view=tools`. Reset demo restores the current work surface's initial state and discards edits in that surface.

## Guided agent demo

Open `#/agent?format=sheets` for the real-file path. The browser loads the bundled `launch-readiness-plan.xlsx`, uses the public `createXlsxAgentAdapter` from `@injoffice/agent-office/xlsx`, and sends JSON calls through `createAgentToolDispatcher` from `@injoffice/agent-tools`. The browser host owns the native Worker, original and replacement bytes, approval, compare-and-swap checks, and exact-output readback.

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

The showcase smoke test runs in CI against both the built subpath deployment and development React. It opens an isolated Chrome profile and checks all 19 catalogue entries and 16 demo surfaces, search, empty-state recovery, cold-route continuity with a deliberately paused lazy chunk, the drawer's focus trap and inert background, checklist retention, source loading, spreadsheet reset and deep links, all four AI approval/refusal workflows, guided controls with different XLSX targets, actual tool traces, stale-source rejection, idempotent retry, write-versus-verification failure, real verified downloads, responsive widths and mobile navigation height, and both color schemes. It fails on uncaught or console errors (including development-only React duplicate-key diagnostics) and prints a temporary directory containing desktop/mobile screenshots. CI retains those screenshots on failure. Set `CHROME_BIN` to override Chrome discovery or `SHOWCASE_OUTPUT` to choose the diagnostic directory.
