# Troubleshooting

Start with the actual runtime and exact failing boundary. A demo label, successful import, or rendered preview is not an end-to-end file-processing check.

| Symptom | Check first | Safe next step |
| --- | --- | --- |
| npm cannot find a package | Release availability and package spelling | Use the release reference or a built source checkout |
| Browser bundle requests Node built-ins | Imported package entry | Switch to the documented browser-safe subpath |
| Worker fails on its first operation | URL, same-origin policy, response body, CSP | Verify all three matched assets; recreate the client after fatal failure |
| WASM compilation fails | MIME type, downloaded bytes, runtime match, CSP | Confirm the response is WASM, not an HTML fallback |
| Native mutation is refused | Issue codes, source fingerprint, target support | Keep original bytes; correct/re-plan the supported operation |
| Stale revision after review | Another write changed the artifact | Re-read and obtain approval for a new plan |
| Commit succeeds but verification fails | Exact committed output and adapter evidence | Withhold verified download; do not claim rollback |
| Browser-only demo has no server readiness | Whether the optional server is deployed | Treat browser readiness and server availability separately |
| Docs or demo on a subpath has broken assets | Build-time base and host routing | Rebuild with the intended base and test a direct deep link |

## A useful bug report

Include package versions, runtime/browser versions, a minimal reproducible snippet, exact issue codes, and whether the failure occurred during extract, plan, apply, or verification. If possible, include a synthetic file that reproduces the issue without private data.

Do not post credentials, confidential archives, tenant IDs, or unredacted network traces in public issues. Follow the repository's security reporting instructions for vulnerabilities.

## Reproduce against the source

```sh
npm ci
npm run build
npm run typecheck
npm run check:packages
npm test -w apps/docs
```

The native Worker integration check is `npm run test:wasm-tarball-browser`. Go modules have separate test commands; see the [server guide](server).
