# Historical completion evidence

`baseline-420424b.json.gz` preserves the exact file bytes and recursive tree
entries of InjOffice commit `420424b57f658a034f7d8cdd76aab20f048a32f4`.
Its tree is `34a2027f1bc0c13ffad8b5d4956e7bd42a1e025d`.

The completion validator uses this snapshot when the historical commit is
unavailable, including shallow clones and the repository's fresh-history
migration. It does not substitute current files for historical evidence, fetch
from the network, import Git commits, or relax any completion gates. All three
published matrix versions retain their original baseline and fixture digests.

The gzip SHA-256 is
`9fd082f8da616cc0eebc63230f8e107ad3c1a603bfe81e7fa614295145912f3b`;
the reader pins and verifies it before using the snapshot.

The payload is UTF-8 JSON with `commit`, `tree`, and `files` fields. Files follow
`git ls-tree -r -z` order. Each entry records the original `mode`, `type`, blob
`object` ID, and base64-encoded bytes from `git cat-file --batch`. There are
1,215 entries. JSON uses compact separators and unescaped Unicode; gzip uses
compression level 9, no filename, and modification time 0.

This is historical source/test data, not a second Git history. The original
licenses, notices, and file-level attributions are included in the snapshot and
remain applicable.
