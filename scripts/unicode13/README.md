# Unicode 13 deterministic projection

`ucd13-projection.txt` is the minimal runtime projection generated from the
vendored official Unicode 13.0.0 files under `ucd/`. Their exact SHA-256
digests are verified by the generator and recorded in both the projection and
generated TypeScript artifact.

Regenerate from verified source files with:

```sh
node scripts/generate-unicode13-tables.mjs
```

CI uses `--check` without network access to rederive and bind the official
sources, committed projection, generator source, and runtime encoding
byte-for-byte. `--from-ucd` may be used to verify a separately obtained copy.

`bidi/BidiCharacterTest-13.0.0.txt.gz` is the official Unicode 13 character
conformance corpus in reproducibly compressed form. Its compressed SHA-256 is
`cbb9d664ea46fcf22dc42bee73cda0f2ffef733e527b7ea375c3d93a50affc5a`;
the decompressed official file SHA-256 is
`b0939c352a162034a58e89c792cdcfa5bb4db5d9fe506e36201d5a78744cf714`.
The bidi tests verify both digests and every explicit-base, BMP-only vector
admitted by the native boundary.
