# Architecture and runtimes

InjOffice separates document meaning, file persistence, and presentation. That separation lets the same source-bound operations run behind different editors or agent frameworks.

## Native Office files

1. Retain the original XLSX, DOCX, or PPTX archive.
2. Extract a versioned native JSON projection with the corresponding Go engine.
3. Validate and inspect the projection; preserve stable IDs and source fingerprints.
4. Construct a supported mutation bound to the original revision.
5. Apply it through the native writer to produce replacement bytes.
6. Re-extract the exact replacement and check the requested effect.

An editor model, canvas image, or preview cannot establish that an unsupported part of an Office file was preserved. Original archive bytes remain the authority. See the [native contracts](../reference/).

## Execution choices

| Layer | Browser | Node / Go host |
| --- | --- | --- |
| Plain JSON contracts and pure helpers | Supported narrow entries | Supported |
| Native XLSX/DOCX/PPTX extraction and writes | Optional Go WASM Worker packages | Go libraries or optional HTTP server |
| PDF page transformations | Browser-safe PDF entry | Same entry for pure byte operations |
| Native font providers and additional PDF tools | Only explicitly browser-safe providers | Node-specific integrations as documented |
| Agent sessions | Host-managed local session | Host-managed service session |
| Durable authorization, persistence, and multi-user ordering | Supplied by your application | Supplied by your application |

There is no automatic browser-to-server fallback. Choose a deployment boundary explicitly; do not upload user files merely because a Worker fails.

## Preview versus persistence

Native paint and render-tree compilers emit bounded preview commands. They can report unsupported content or refuse incomplete font/layout inputs. They are not Office-compatible rasterizers and are not serializers.

Authored `DeckSpec` data is a separate workflow from editing an imported PPTX archive. Likewise, an in-memory spreadsheet or collaboration sample is not evidence of native file round-trip support.

## AI without a model dependency

The agent packages operate on structured proposals. A model, deterministic function, or human can supply those proposals. The host owns any LLM integration, authentication, approval UI, and storage.

The public demo intentionally simulates proposal generation while executing real bounded document operations. A working mock flow proves the document tool boundary; it does not measure model reasoning quality.
