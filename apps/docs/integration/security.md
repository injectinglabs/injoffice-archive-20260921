# Security and production checklist

The libraries enforce their format and lifecycle contracts. A production application must supply its own security and operational boundary.

## Untrusted files

- Enforce compressed input size and document-specific resource budgets before processing.
- Keep original files until output acceptance is complete.
- Isolate risky processing according to your deployment's threat model.
- Preserve or refuse unsupported content; do not strip it silently.
- Handle malformed, encrypted, oversized, and unsupported files as explicit outcomes.
- Avoid recording document contents, tokens, or signed URLs in logs.

## Identity and writes

- Authenticate the actual user; a client-supplied actor ID is not proof of identity.
- Authorize reads, previews, commits, downloads, and restores per artifact/tenant.
- Enforce compare-and-swap at the persistent write boundary.
- Bind approval to the exact plan and source revision.
- Keep durable idempotency records when retries can cross processes.
- Verify the committed replacement, not merely the planned model.

## Browser hosting

- Publish version-matched Worker, WASM, and runtime assets with correct MIME types.
- Use a reviewed CSP and appropriate CORS settings.
- Keep old hashed assets available for already-open tabs during a release.
- Test production URLs and non-root paths, not only development proxies.
- Do not introduce an implicit remote file-upload fallback.

## Server hosting

- Do not expose the optional localhost server as a tenant-safe service without additional controls.
- Put TLS, request limits, authorization, storage isolation, and monitoring at the appropriate boundary.
- Define retention, backups, restore testing, and resource exhaustion behavior.
- Make failures observable without leaking file content.

## Dependencies and releases

Pin tested versions and review dependency licenses and notices. Run the repository's package, provenance, and dependency-integrity checks for source-based releases. The [dependency reference](../reference/generated/contracts/dependency-transparency) describes what the repository records.

This checklist is not a security audit or a certification. Validate your actual deployment, files, threat model, and operational controls.
