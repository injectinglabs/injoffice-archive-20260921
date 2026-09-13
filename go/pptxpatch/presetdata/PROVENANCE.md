# DrawingML preset catalog source

`preset-shapes.xml` contains the unchanged Apache POI preset resource from commit `338882ac8898df5c13a7d15f533204c5dd8607d6`:

https://github.com/apache/poi/blob/338882ac8898df5c13a7d15f533204c5dd8607d6/poi/src/main/resources/org/apache/poi/sl/draw/geom/presetShapeDefinitions.xml

SHA-256: `4a762444d8d85876881c02a5b1dedf6f73006fcd8acb7b4e393435615b37c780`.

The upstream `legal/LICENSE` and `legal/NOTICE` from the same revision accompany it unchanged. Only preset definitions are incorporated; the upstream legal files also mention components of the broader POI distribution.

The runtime derives evaluated paths from these definitions. It performs deterministic guide alpha-renaming for sequentially reassigned names and excludes validated nonvisual adjustment-handle/connection metadata from paint evaluation. The original resource remains byte-for-byte intact for provenance and audit.

`preset-shapes.xml.gz` is generated losslessly from these exact XML bytes by
`scripts/generate-pptx-preset-catalog.mjs` (Node 22). It has no filename or
timestamp and uses an OS-neutral header. The runtime embeds only this compressed
copy, bounds inflation to the original byte count, rejects trailing members,
and verifies the original SHA-256 before parsing. Run the generator with
`--check` to verify reproducibility under the pinned Node toolchain.
