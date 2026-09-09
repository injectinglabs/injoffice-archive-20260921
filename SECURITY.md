# Security policy

## Reporting a vulnerability

Do not open a public issue for a suspected vulnerability. When enabled, use
[GitHub private vulnerability reporting](https://github.com/injectinglabs/injoffice/security/advisories/new).
This feature is unavailable while the repository is private. Before making the
repository public, maintainers must enable it and verify that the reporting link
works. Collaborators with private repository access should contact a maintainer
through their existing private communication channel if reporting is unavailable.

Include the affected package and version or commit, impact, reproduction steps, and any suggested mitigation. Do not include real customer documents or secrets. A minimal synthetic artifact is preferred.

Maintainers will acknowledge a complete report as soon as practical, investigate it privately, and coordinate disclosure and a fixed release when the issue is confirmed. No response-time guarantee is made while the project is pre-1.0.

## Supported versions

The default branch and the latest 0.1.x release are supported. Users should expect
security fixes in the newest compatible pre-1.0 minor release unless a maintainer
states otherwise. See the [dependency advisory and consumer mitigation](docs/DEPENDENCY-TRANSPARENCY.md#security-advisory-snapshot)
for the Univer dependency used with 0.1.0.

## Security model

InjOffice parses complex, attacker-controlled ZIP/XML, PDF, font, JSON, and image data. Embedding applications should enforce input-size, time, memory, URL, and concurrency limits; isolate native or WASM processing where appropriate; and never treat successful parsing as proof that a document is benign.

The connector and collaboration packages do not provide authentication or authorization. Hosts must implement those controls and follow the requirements in the collaboration protocol.
