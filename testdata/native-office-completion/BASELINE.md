# Historical completion evidence

`baseline-420424b.json.gz` preserves the recursive tree metadata of InjOffice
commit `420424b57f658a034f7d8cdd76aab20f048a32f4` and the exact bytes needed by the
three completion matrices. Its tree is `34a2027f1bc0c13ffad8b5d4956e7bd42a1e025d`.

The snapshot has 1,215 tree entries, but only 189 file payloads: the 186 files
read by the matrices and their transitive authority checks, plus `LICENSE`,
`LICENSE-UNICODE.txt`, and `NOTICE`. Unrelated historical documentation, app code,
and debugging notes have no payload. Required evidence is never rewritten to
match current source or to remove original attribution.

Keeping the complete tree metadata lets import resolution identify missing
payloads and fail closed instead of silently skipping a historical dependency.
The validator uses the snapshot when the historical commit is unavailable,
including shallow clones and fresh-history repositories. No network fetch or
historical Git commit import is required. All matrix versions keep their original
baseline and fixture digests.

The gzip SHA-256 is
`82a2eef18c6754acd85bea0873b63b178231b74a9498d42aa93997002978699a`;
the reader pins and verifies it before using the snapshot.

The payload is UTF-8 JSON with `commit`, `tree`, and `files` fields. Entries follow
`git ls-tree -r -z` order and retain their original `mode`, `type`, and blob
`object` ID. Only retained evidence entries also have `base64` bytes from the
original blob. Tests verify every retained payload against its Git blob hash and
validate all three matrices without access to the historical commit.

JSON uses compact separators and unescaped Unicode; gzip uses compression level
9, no filename, and modification time 0. To reproduce payloads, read the original
blobs for entries that have `base64`; do not substitute current file contents.
The original licenses, notices, and file-level attributions remain applicable.
