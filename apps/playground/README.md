# InjOffice showcase

Run `npm run dev` from the repository root (Node 22 or newer). The playground normally starts on port 3100.

The landing page searches the registered examples by task, file type, package, and recipe language. Every example has a focused work surface and a Source / proof drawer containing a manual checklist, expected evidence, exact repository links, related examples, and lazily loaded demo source. Closing the drawer retains checklist progress on the current surface; the checklist does not perform or certify engine operations.

The spreadsheet modes have shareable links: `#/sheets?view=editor`, `#/sheets?view=native`, and `#/sheets?view=tools`. Reset demo restores the current work surface's initial state and discards edits in that surface.

## Agent integration example

Open `#/agent?format=sheets` for the real-file path. The browser loads the bundled `launch-readiness-plan.xlsx`, uses the public `createXlsxAgentAdapter` from `@injoffice/agent-office/xlsx`, and sends JSON calls through `createAgentToolDispatcher` from `@injoffice/agent-tools`. The browser host owns the native Worker, original and replacement bytes, approval, compare-and-swap checks, and exact-output readback.

The editable request uses a bounded deterministic parser, not a language model. Try `Mark Security as Ready` and `Mark Mobile as On track`: the proposal resolves a different target from the workbook's bounded read. Unsupported requests are refused rather than silently interpreted as the default edit. No model SDK, provider credentials, or document upload is required. Docs, Slides, and PDF remain explicitly labelled object-backed lifecycle simulations; they do not write Office file bytes.

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

Development smoke also exercises the optional live-proposal bridge against its own loopback-only mock, never a real provider. It temporarily replaces any configured proposal URL and removes any configured token, then restores both environment values during cleanup without printing them. The test proves context discovery sends no upstream request, consent is required before sharing the exact disclosed context, provider-supplied approval flags cannot authorize a write, and separate human approval produces a verified XLSX. Built and existing-server smoke never trigger live-provider requests.

The showcase smoke test runs in CI against both the built subpath deployment and development React. It opens an isolated Chrome profile and checks all 19 catalogue entries and 16 demo surfaces, search, empty-state recovery, cold-route continuity with a deliberately paused lazy chunk, the drawer's focus trap and inert background, checklist retention, source loading, spreadsheet reset and deep links, all four AI approval/refusal workflows, editable requests with different XLSX targets, actual tool traces, stale-source rejection, idempotent retry, write-versus-verification failure, the real XLSX download, responsive widths and mobile navigation height, and both color schemes. It fails on uncaught or console errors (including development-only React duplicate-key diagnostics) and prints a temporary directory containing desktop/mobile screenshots. CI retains those screenshots on failure. Set `CHROME_BIN` to override Chrome discovery or `SHOWCASE_OUTPUT` to choose the diagnostic directory.
