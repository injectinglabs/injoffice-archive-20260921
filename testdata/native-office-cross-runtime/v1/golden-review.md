# Semantic golden review: source-bound table cell flows

Compared independently built revisions `ec5e089` and `04dc6c0`, using the same
locked dependencies and the unchanged checked-in qualification inputs.

The semantic report changed from
`d4d2592e18b4e86d4183c960478ed88a76dbbb78b82aefbc7c0248756b277a89` to
`ae83f79d5d7a34b506017c2f37411186b537b1433d371151115d48dd77931d76`.

Recursive comparison of the full DOCX pagination, paint request, and paint
output found only these changes:

- Page 0, paragraph slice 1 gained `table_cell_id: "cell:summary:1:1"`.
- Page 0, placed line 1 gained the same `table_cell_id`.
- The request's embedded pagination contains those same two additions.
- The request's pagination integrity hash and the output's pagination
  provenance hash changed from
  `sha256:a7c38153d98476131a4e71ce27a64d5307fcae1f5b4e7120febbbbc2bff4ce22`
  to `sha256:7e425eb81fd80429c949a66bd9b8757b74f10cce432a2d79c4537f78076bf787`.

All page coordinates, glyph paths, paint commands, resources, source extraction,
shaping, refusal results, and complete PPTX/XLSX reports were unchanged.
The additions bind adjacent table-cell flows to source identities; they prevent
valid side-by-side cells from being rejected as overlapping ordinary paragraphs.
The qualification script now also asserts these exact cell identity joins.

For future structural investigation, run
`node scripts/qualify-native-office-repro.mjs --worker --inspect-docx` in separately
built revisions. This read-only diagnostic emits the full original-fixture DOCX
objects instead of their digest summary. Normal qualification still compares
the complete semantic report across all manifest profiles and against its pin.
This is a deterministic regression baseline, not an external Office reference.
