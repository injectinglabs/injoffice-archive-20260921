# Security policy

## Reporting a vulnerability

Do not open a public issue for a suspected vulnerability. Use GitHub's private vulnerability reporting feature for this repository. If that feature is unavailable, contact the repository owner privately through the verified contact listed on the GitHub organization profile.

Include the affected package and version or commit, impact, reproduction steps, and any suggested mitigation. Do not include real customer documents or secrets. A minimal synthetic artifact is preferred.

Maintainers will acknowledge a complete report as soon as practical, investigate it privately, and coordinate disclosure and a fixed release when the issue is confirmed. No response-time guarantee is made while the project is pre-1.0.

## Supported versions

Before the first public release, only the default branch is supported. After publishing begins, this file will list supported release lines. Users should expect security fixes to land in the newest compatible pre-1.0 minor release unless a maintainer states otherwise.

## Security model

InjOffice parses complex, attacker-controlled ZIP/XML, PDF, font, JSON, and image data. Embedding applications should enforce input-size, time, memory, URL, and concurrency limits; isolate native or WASM processing where appropriate; and never treat successful parsing as proof that a document is benign.

The connector and collaboration packages do not provide authentication or authorization. Hosts must implement those controls and follow the requirements in the collaboration protocol.
