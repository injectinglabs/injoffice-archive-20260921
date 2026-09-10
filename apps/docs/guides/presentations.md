# Presentations / PPTX

Choose between two distinct workflows: edit an existing source-bound PPTX, or author a new deck model. A model created from scratch is not a lossless representation of an arbitrary imported PowerPoint file.

## Inspect an existing file

`@injoffice/pptx-wasm` extracts native models in a browser Worker. Use source-anchored IDs to choose targets:

<<< @/examples/pptx.ts

For a native write, retain the full extracted deck and original bytes. `client.apply` takes `expectedSourceRevision` plus source-bound operations. Text operations bind the element ID and `expectedFingerprintSha256`; replacement paragraphs carry explicit styling. See the complete [PPTX Worker example](../reference/generated/packages/pptx-wasm).

The native writer supports exact parsed text and AutoShape property replacements. It refuses general slide insertion/removal, pictures, charts, tables, groups, connectors, animation, and transition editing on imported files.

## Author a new deck

Use the DOM-free authoring entry and compile a supported `DeckSpec` into a native model:

<<< @/examples/presentation.ts

This returns JSON models and audit results, not `.pptx` file bytes. The Go PPTX module is the native archive generation/persistence boundary. Do not rename JSON output to `.pptx`.

## Render a preview

`@injoffice/pptx-render` compiles validated native slides into a renderer-neutral tree. Text needs explicit font-manifest, resolver, and shaper inputs. A Canvas host can replay commands, but browser measurements cannot substitute for the native layout contract.

The compiler can refuse unsupported content or incomplete font authority. Neither a rendered screenshot nor an authored preview proves PowerPoint compatibility.

## Add agent operations

Use `createNativePptxAgentAdapter` for source-bound native edits or `createAuthoredDeckAgentAdapter` for a `DeckSpec`. They have different capabilities and should not share an undocumented fallback. See the [agent-office reference](../reference/generated/packages/agent-office).

## References

- [Native PPTX contract](../reference/generated/contracts/pptx-native-contract)
- [Authored compiler](../reference/generated/packages/pptx-authored)
- [Render tree contract](../reference/generated/contracts/pptx-render-tree)
- [Go PPTX engine](../reference/generated/go/pptxpatch)
