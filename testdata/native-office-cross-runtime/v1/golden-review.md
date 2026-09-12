# Semantic golden review: explicit font-selection evidence

Compared independently built revisions `de7fc902` and `54f9380` with unchanged
qualification inputs. The report changes from
`ae83f79d5d7a34b506017c2f37411186b537b1433d371151115d48dd77931d76` to
`91baa725e6cbf4f6d293a8ee4a1bc479b87c3bfc86025a625e8e74fa5f17c2fc`.

The only report change is `pptx.imported.render_tree_sha256`, from
`d60d513ac19abeca0716658a279c6b501c0b1b3fba1bbc9148f0c2ee49fec485` to
`c24cb8f335100f1a638e7ba04a2d5ed1cd93352d697079a464d150726e6f83d7`.
Recursive comparison of the complete imported render trees found exactly four
additions: `/nodes/4/cells/{0,1,2,3}/paragraph/runs/0/fontSelection`, each
`{sourceFamily: 'Fixture Sans', selectedFamily: 'Fixture Sans', resolution: 'exact'}`.
All coordinates, glyphs, diagnostics, fidelity labels and other fields are
unchanged. Complete DOCX and XLSX reports are unchanged. This is additive font
provenance, not a visual-fidelity improvement or an external Office reference.
The existing three locale/timezone profiles continue to require identical output.

## Previous review: source-bound table cell flows

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
