# InjOffice showcase

Run `npm run dev` from the repository root (Node 22 or newer). The playground normally starts on port 3100.

The showcase is one continuous page: the overview and nineteen live examples follow the same order as the sidebar. Scroll to explore, or use the sidebar and searchable catalogue to jump to a section. The highlighted sidebar link follows your position. Passive scrolling updates the shareable URL without adding history entries or moving keyboard focus; explicit links support Back/Forward and direct deep links.

Each live demo loads near the viewport and stays mounted after loading, so scrolling away preserves its edits and pending approvals. The four AI formats have independent sessions. AI workflows use the document's vertical scroll; canvas editors and long code traces retain their own necessary scrolling. Reset affects only its section; refreshing the page discards local demo state. A failed lazy download is isolated to its section, with retry and an explicitly confirmed reload fallback for browser-cached failures.

Every example has a Guide & source drawer containing a manual checklist, expected evidence, exact repository links, related examples, and lazily loaded demo source. Closing the drawer retains checklist progress on the current surface; the checklist does not perform or certify engine operations.

The spreadsheet modes have shareable links: `#/sheets?view=editor`, `#/sheets?view=native`, and `#/sheets?view=tools`. Reset demo restores the current work surface's initial state and discards edits in that surface.

## Agent integration example

Open `#/agent?format=sheets` for the real-file path. The browser loads the bundled `launch-readiness-plan.xlsx`, uses the public `createXlsxAgentAdapter` from `@injoffice/agent-office/xlsx`, and sends JSON calls through `createAgentToolDispatcher` from `@injoffice/agent-tools`. The browser host owns the native Worker, original and replacement bytes, approval, compare-and-swap checks, and exact-output readback.

The default **Built-in mock agent (no LLM)** needs no endpoint configuration or credentials. Try `Mark Security as Ready`, `Mark Mobile as On track`, or `Set Analytics status to Review`: a deterministic mock resolves one status target from the workbook's bounded read. The UI labels the proposal as simulated and shows its request/response separately from actual Office tool calls. Unsupported requests are refused rather than silently interpreted as the default edit.

In development, the browser calls the bundled `POST /api/agent/mock-propose` route on the same local server. Static builds simulate the response in the browser using the same generator, with no API call. Neither path contacts a provider, forwards an API token, or grants approval—even if live-provider environment variables are set. No model SDK or external service is required. The original local rule-based mode and explicitly consented live integration remain optional choices; see [proposal modes](../../docs/AGENT-PROPOSAL-HOST.md).

Only the proposal is mocked. All four agent demos edit and verify real files:

| Format | Demonstrated edit | Output verification |
| --- | --- | --- |
| XLSX | One discovered workstream status cell | Native browser Worker re-extracts saved bytes |
| DOCX | One exact, source-anchored text run | Native browser Worker re-extracts saved bytes |
| PPTX | One exact native text element | Native browser Worker re-extracts saved bytes |
| PDF | One page rotation | Fresh PDF parse checks rotation, page count, geometry, and SHA-256 |

Docs starts with `Replace "Northstar Launch Brief" with "Northstar Beta Launch Brief"`; Slides uses `Replace "Northstar launch review" with "Northstar: ready for launch"`; PDF uses `Rotate page 2 by 90 degrees`. Each uses a populated fixture, an isolated preview, exact-plan approval, verified download, and the same stale-source/retry/fault scenarios. Text and page-metadata previews are bounded projections, not full-fidelity Office renderers. These examples do not imply arbitrary edits or uploaded-file support.

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

Development smoke also exercises the optional live-proposal bridge against its own loopback-only mock, never a real provider. It temporarily replaces any configured proposal URL and removes any configured token, then restores both environment values during cleanup without printing them. The test proves context discovery sends no upstream request, consent is required before sharing the exact disclosed context, provider-supplied approval flags cannot authorize a write, and separate human approval produces a verified XLSX. Built and existing-server smoke never trigger live-provider requests.

The showcase smoke test runs in CI against both the built subpath deployment and development React. It opens an isolated Chrome profile and checks all 19 catalogue entries and 16 demo surfaces, search, empty-state recovery, cold-route continuity with a deliberately paused lazy chunk, the drawer's focus trap and inert background, checklist retention, source loading, spreadsheet reset and deep links, all four AI approval/refusal workflows, editable requests with different XLSX targets, actual tool traces, stale-source rejection, idempotent retry, write-versus-verification failure, the real XLSX download, responsive widths and mobile navigation height, and both color schemes. It fails on uncaught or console errors (including development-only React duplicate-key diagnostics) and prints a temporary directory containing desktop/mobile screenshots. CI retains those screenshots on failure. Set `CHROME_BIN` to override Chrome discovery or `SHOWCASE_OUTPUT` to choose the diagnostic directory.
