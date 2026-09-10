# Reference

Use the guides to complete a task and the source-derived reference to inspect the exact package or protocol boundary.

## Package and Go references

The [complete reference index](generated/) covers all 26 TypeScript packages and seven Go modules. Each package page includes its maintained README, examples, limitations, and an exported-name inventory derived from built declarations.

The inventory is a symbol index, not a complete parameter-level generated API manual. Follow linked contracts and use the `.d.ts` files from your installed version for exact signatures, properties, and overloads. These pages track the repository's `main` branch.

## Start with a core contract

- [Agent sessions and tools](generated/packages/agent-tools)
- [Format-specific agent adapters](generated/packages/agent-office)
- [Sheets models and mutation protocol](generated/packages/sheets)
- [Native DOCX contract](generated/contracts/docx-native-contract)
- [Native PPTX contract](generated/contracts/pptx-native-contract)
- [PPTX render tree](generated/contracts/pptx-render-tree)
- [Collaboration protocol](generated/contracts/collaboration-protocol)
- [Optional HTTP server](generated/go/injoffice-server)

## Compatibility and release information

- [Public-release status and checklist](generated/contracts/public-release)
- [Dependency transparency](generated/contracts/dependency-transparency)
- [Univer Sheets compatibility](generated/contracts/univer-sheets-compatibility)
- [Formula inventory](generated/contracts/functions)

Reference text is regenerated at documentation build time. Correct the linked source README/contract rather than editing an ignored generated page. Source-derived examples may be integration fragments; the separately maintained [guide examples](../examples/) identify which snippets are typechecked or executed.
