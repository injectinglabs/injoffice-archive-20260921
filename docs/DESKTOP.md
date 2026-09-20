# InjOffice Desktop

InjOffice Desktop is a local Electron editor around the same native extract/apply
engines as the playground. It has no AI, no account, and no required hosted
backend. The renderer stays offline except the optional GitHub updater in
**signed public** builds.

The private workspace is `apps/desktop` (`@injoffice/desktop`, not published).
Host adapters, the mock-tested Electron main process, the start page, workspace
shell, OpenError, hidden-apply scheduler, and format helpers live there and
are tested in the `core` CI shard. **Electron is a desktop host dependency**;
renderer `/src/` stays offline (`check-office-architecture` allows the desktop
host to list it; the lockfile may contain it only when pulled by
`apps/desktop`. Renderer `/src/` still forbids `electron` imports).
Root `desktop:build`, `desktop:start`, and `desktop:dist` scripts forward to
`@injoffice/desktop`.

This is not Microsoft Office parity. Do not advertise it as such.

## Run from source

From the repository root:

```bash
npm ci
npm run desktop:build
npm run desktop:start
```

`npm run desktop:dist` packages unsigned installers into `apps/desktop/release/`.

Needs Node.js 22 or newer, the Go version declared in the native modules, and
Bash (Git Bash on Windows). The desktop build compiles the bundled WASM
engines (`@injoffice/docx-wasm`, `@injoffice/xlsx-wasm`, `@injoffice/pptx-wasm`
via the workspace `prebuild` script) and the Electron host. Vite fails the
renderer build when any engine `.wasm`, `wasm_exec.js`, or worker is missing
from those packages' `dist/`, or when the bundle has no hashed copy of them;
`INJOFFICE_ALLOW_STUBS=1` downgrades that to a warning for local UI work only
and must never be set for a packaged build. It does **not** start `injoffice-server` and it
does **not** use the playground on port 3100.

Create or open `.docx`, `.xlsx`, `.pptx`, or `.pdf`. Unsupported mutations
refuse and leave prior bytes. Binary Word 97–2003 (`.doc`), PowerPoint 97–2003
(`.ppt`), Excel `.xlsb`, OpenDocument `.ods`, and encrypted compound-file
`.docx` fail closed on extract. Macro-enabled `.docm` is zip OOXML and can
paint; VBA is not executed.

## Engine lockstep

Desktop WASM must track current `main`. After an engine merge, rebuild desktop
from that SHA. Do not freeze a lagging snapshot as the advertised product.
Unsigned developer previews are not public updates.

## Builds and updates

Installer output belongs in `apps/desktop/release/` and is gitignored. Do not
commit DMGs, NSIS installers, AppImages, or `latest-*.yml`.

Packaging icons live in `apps/desktop/build/icons/`: `icon.png` (1024 × 1024,
Linux), `icon.icns` (macOS), and `icon.ico` (Windows). They are nearest-neighbor
conversions of repository-root `logo.png`. Regenerate on macOS with
`python3 apps/desktop/build/icons/generate.py`.

### Unsigned developer previews

GitHub Actions → **Desktop builds** (`.github/workflows/desktop.yml`,
`workflow_dispatch` only). Matrix:

| Platform | Artifact |
| --- | --- |
| mac-arm64 | DMG + ZIP |
| mac-x64 | DMG + ZIP |
| win-x64 | NSIS |
| linux-x64 | AppImage, DEB, RPM |

These artifacts are unsigned. They must not set `injofficeRelease` and must not
drive the in-app public updater. `CSC_IDENTITY_AUTO_DISCOVERY` is false.

### Signed public drafts

Tag a commit `desktop-v*` (desktop versions are independent of npm package
versions) or dispatch **Desktop release draft**
(`.github/workflows/desktop-release.yml`) against that tag. The workflow uses
the `desktop-release` GitHub environment and creates a **draft** GitHub
Release. It never publishes. Ordinary CI artifacts and local `electron-builder`
output do not become public updates.

Required environment secrets (not supplied by this repository):

| Secret | Purpose |
| --- | --- |
| `MAC_CSC_LINK` / `MAC_CSC_KEY_PASSWORD` | Developer ID Application P12 |
| `APPLE_ID` / `APPLE_APP_SPECIFIC_PASSWORD` / `APPLE_TEAM_ID` | Notarization |
| `WINDOWS_CSC_LINK` / `WINDOWS_CSC_KEY_PASSWORD` | Authenticode PFX |

Publish a draft only after install, open/edit/save, and upgrade checks on each
shipped platform. If a release is broken, ship a higher patch version; do not
replace published bytes.

In-app updates (signed builds only) use `electron-updater` against GitHub
Releases whose tags start with `desktop-v`. The renderer remains offline; the
updater uses its own Electron session. Downloads and restart are
user-initiated. Restart must refuse unsaved documents and failed recovery
writes.

DEB/RPM users install a new package with their package manager. There is no
APT/YUM repository and no in-app update for those packages. Linux ARM, Windows
ARM, beta channels, staged rollouts, and store distribution are out of scope.

## Tests and CI

Desktop **host** tests (Node, no GUI) belong in the `core` shard of
`scripts/ci-test-shards.json` once the workspace has a `test` script. Native
WASM editor tests, packaging, and signed builds are **not** PR merge gates.

```bash
npm run test -w @injoffice/desktop
npm run typecheck -w @injoffice/desktop
# after desktop:build; no GUI:
npm run test:native -w @injoffice/desktop
```

Electron, TipTap, Univer, and HTML/SVG preview are not the OOXML file
authority. `scripts/check-office-architecture.mjs` must keep forbidding
Electron as native authority.

## Support limits

- No claim of Word/Excel/PowerPoint/Acrobat feature or layout parity.
- DOCX on-screen preview may use flowing HTML; page-paint is a separate export
  path and is the layout that should match a printed page.
- Public updates come only from **published** `desktop-v*` GitHub Releases.
- External Office oracles and private fidelity corpora stay local; do not add
  them to CI.
