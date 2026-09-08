# InjOffice showcase

Run `npm run dev` from the repository root (Node 22 or newer). The playground normally starts on port 3100.

The landing page searches the registered examples by task, file type, package, and recipe language. Every example has a focused work surface and a Source / proof drawer containing a manual checklist, expected evidence, exact repository links, related examples, and lazily loaded demo source. Closing the drawer retains checklist progress on the current surface; the checklist does not perform or certify engine operations.

The spreadsheet modes have shareable links: `#/sheets?view=editor`, `#/sheets?view=native`, and `#/sheets?view=tools`. Reset demo restores the current work surface's initial state and discards edits in that surface.

## Verification

- `npm run test -w apps-playground`
- `npm run typecheck -w apps-playground`
- `npm run test:office-browser-smoke:built` after building workspace packages
- `SHOWCASE_URL=http://127.0.0.1:3100/ npm run test:showcase-browser` with the playground running
- `npm run test:showcase-browser -- --built` after the Office browser smoke build above; starts and stops its own ephemeral server at `/injoffice-smoke/`, without touching an existing demo server

The showcase smoke test runs in CI against the built subpath deployment. It opens an isolated Chrome profile and checks all 19 catalogue entries and 16 demo surfaces, search, empty-state recovery, cold-route continuity with a deliberately paused lazy chunk, the drawer's focus trap and inert background, checklist retention, source loading, spreadsheet reset and deep links, all four AI approval/refusal workflows, the real XLSX download, responsive widths and mobile navigation height, and both color schemes. It fails on uncaught or console errors and prints a temporary directory containing desktop/mobile screenshots. CI retains those screenshots on failure. Set `CHROME_BIN` to override Chrome discovery or `SHOWCASE_OUTPUT` to choose the diagnostic directory.
