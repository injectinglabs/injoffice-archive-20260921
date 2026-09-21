# ADR: Self-contained Office engine boundaries

- **Status:** Accepted
- **Date:** 2026-09-02
- **Decision owners:** InjOffice maintainers

## Context

InjOffice ships native XLSX, PPTX, and DOCX engines behind InjOffice-owned
contracts. Format authority and persistence must remain independent of optional
editor shells and rendering hosts.

## Decision

InjOffice does not take a runtime dependency on or vendor another Office suite's
source tree. Its public types, persistence protocols, and renderer inputs remain
maintained by InjOffice.

Do not add a production, test, or CI dependency on another Office suite's
packages. `scripts/check-office-architecture.mjs` already forbids Electron and
Mammoth as native authority, and must keep doing so.

Compatibility evidence is InjOffice corpus fixtures and Microsoft-authored
packages already in this repository. Licensed fixture and catalog data retain
their upstream provenance and notices.

## Consequences

- Engine work proceeds from ECMA-376 / OPC and InjOffice contracts, not from a
  third-party source tree.
- Architecture CI remains the gate against Electron and Mammoth as native authority.
