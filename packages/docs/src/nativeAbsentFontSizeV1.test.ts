import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import {
  DOCX_ABSENT_FONT_SIZE_WARNING,
  DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1,
  projectNativeDocxAbsentFontSizesV1,
  validNativeDocxAbsentDefaultSizeShapeV1,
  validNativeDocxAbsentFontSizesV1,
  validNativeDocxApproximatedFontSizesV1,
  validNativeDocxHostDefaultSizePolicyV1,
} from "./nativeAbsentFontSizeV1.js";
import {
  decodeNativeDocxResolvedLayout,
  nativeDocxResolvedNumberingDefinitionSha256V1,
  nativeDocxResolvedNumberingModelSha256V1,
  type NativeDocxResolvedLayoutInputV1,
} from "./nativeResolvedLayout.js";

const hash = `sha256:${"a".repeat(64)}`;
const fact = {
  scope_kind: "paragraph-mark" as const,
  scope_id: "paragraph:1",
  part_name: "word/document.xml",
  path: "/w:document[1]/w:body[1]/w:p[1]",
  package_sha256: hash,
};
describe("declared host-size source evidence", () => {
  it("pins each host default to the size Microsoft Word 16.112 exports for that shape", () => {
    // Read out of the Tf operators of Word 16.112's own PDF exports, which
    // place text on a 1/300 in grid: listWithLgl.docx and
    // sdt_after_section_break.docx (no w:docDefaults record) are written at 50
    // units = 12 pt, NumberedList.docx and ImageCrop.docx (a w:docDefaults
    // record stating no w:sz) at 42 units = 10 pt. Neither is 11 pt.
    expect(DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1).toEqual({
      "absent-document-defaults": 24,
      "sizeless-document-defaults": 20,
    });
    // The disclosure is derived, so it can never name a size nobody applies.
    expect(DOCX_ABSENT_FONT_SIZE_WARNING).toContain("10 or 12 pt");
    for (const shape of ["absent-document-defaults", "sizeless-document-defaults"])
      expect(validNativeDocxAbsentDefaultSizeShapeV1(shape)).toBe(true);
    for (const shape of [undefined, "", "absent", 24, null])
      expect(validNativeDocxAbsentDefaultSizeShapeV1(shape)).toBe(false);
  });
  it("bounds unique source facts and rejects unknown/foreign fields", () => {
    expect(validNativeDocxAbsentFontSizesV1([fact], hash)).toBe(true);
    for (const candidate of [
      [fact, fact],
      [{ ...fact, package_sha256: `sha256:${"b".repeat(64)}` }],
      [{ ...fact, scope_kind: "global" }],
      [{ ...fact, part_name: "../document.xml" }],
      [{ ...fact, path: "/wrong" }],
      [{ ...fact, guessed: true }],
      Array.from({ length: 1001 }, (_, i) => ({
        ...fact,
        scope_id: `paragraph:${i}`,
      })),
    ])
      expect(validNativeDocxAbsentFontSizesV1(candidate, hash)).toBe(false);
  });
  it("requires an explicit declared host choice and full retained source coverage", () => {
    for (const halfPoints of Object.values(DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1))
      expect(
        validNativeDocxHostDefaultSizePolicyV1({
          kind: "host-default-size-v1",
          half_points: halfPoints,
        }),
      ).toBe(true);
    for (const policy of [
      undefined,
      // 11 pt was the host's own invention and matches no Word reference.
      { kind: "host-default-size-v1", half_points: 22 },
      { kind: "word-default", half_points: 24 },
      { kind: "host-default-size-v1", half_points: 20, exact: true },
    ])
      expect(validNativeDocxHostDefaultSizePolicyV1(policy)).toBe(false);
    const applied = [{ ...fact, chosen_half_points: 20 }];
    const shape = "sizeless-document-defaults";
    expect(
      validNativeDocxApproximatedFontSizesV1(applied, [fact], hash, shape),
    ).toBe(true);
    expect(
      validNativeDocxApproximatedFontSizesV1(applied, [], hash, shape),
    ).toBe(false);
    // The other shape's Word-derived size is still a forgery for this package.
    expect(
      validNativeDocxApproximatedFontSizesV1(
        [{ ...fact, chosen_half_points: 24 }],
        [fact],
        hash,
        shape,
      ),
    ).toBe(false);
    expect(
      validNativeDocxApproximatedFontSizesV1(
        [{ ...fact, chosen_half_points: 22 }],
        [fact],
        hash,
        shape,
      ),
    ).toBe(false);
    expect(
      validNativeDocxApproximatedFontSizesV1(applied, [fact], hash, undefined),
    ).toBe(false);
    expect(
      validNativeDocxApproximatedFontSizesV1(
        [{ ...fact, chosen_half_points: 24 }],
        [fact],
        hash,
        "absent-document-defaults",
      ),
    ).toBe(true);
    expect(
      validNativeDocxApproximatedFontSizesV1(
        applied,
        [{ ...fact, scope_id: "paragraph:other" }],
        hash,
        shape,
      ),
    ).toBe(false);
  });
});

/**
 * The separator story the extractor qualifies may carry ordinary paragraphs
 * after its instruction paragraph, and the extractor emits missing-size facts
 * for them. This module keeps its own copy of that qualification, and when the
 * two drifted the carried paragraph was excluded from the joinable set and the
 * exact-join threw `Missing-size scope anchor does not exact-join` — which is
 * what took `NumberedList.docx`, one of the two 10 pt packages the host-size
 * policy was measured on, from painting to HTTP 422.
 */
describe("missing-size facts inside a reserved separator story", () => {
  const source = JSON.parse(readFileSync(new URL("../../../testdata/docx-native/document-v1.json", import.meta.url), "utf8"));
  const notePart = "word/footnotes.xml";
  const noteAnchor = (path: string) => ({ part_name: notePart, path, start_byte: 10, end_byte: 20, xml_sha256: `sha256:${"0".repeat(64)}` });
  const readOnly = { mode: "read-only", allowed_operations: [], refusal: { code: "NATIVE_READ_ONLY", message: "Reserved separator source remains authoritative.", preservation: "refuse-mutation" } };
  const separatorStory = (carried: boolean) => ({
    id: "story:footnote:separator", kind: "footnote", part_name: notePart, native_story_id: "-1", relationship_id: "rIdFootnotes",
    note_role: "separator", anchor: noteAnchor("/w:footnotes[1]/w:footnote[2]"),
    blocks: [
      { kind: "paragraph", id: "paragraph:separator:instruction", paragraph: { id: "paragraph:separator:instruction", anchor: noteAnchor("/w:footnotes[1]/w:footnote[2]/w:p[1]"), edit_policy: readOnly, properties: {}, runs: [] } },
      ...(carried ? [{ kind: "paragraph", id: "paragraph:separator:carried", paragraph: { id: "paragraph:separator:carried", anchor: noteAnchor("/w:footnotes[1]/w:footnote[2]/w:p[2]"), edit_policy: readOnly, properties: {}, runs: [] } }] : []),
    ],
  });
  const document = (carried: boolean) => ({ ...structuredClone(source), notes: [separatorStory(carried)], unsupported: [] });
  const paragraphID = carriedFact().scope_id;
  function carriedFact() {
    return { scope_kind: "paragraph-mark" as const, scope_id: "paragraph:separator:carried", part_name: notePart, path: "/w:footnotes[1]/w:footnote[2]/w:p[2]", package_sha256: source.source.package_sha256 };
  }
  const layout = {
    protocol: "injoffice.docx.resolved-layout", version: 1, document_id: source.document_id, revision: source.revision,
    source_parts: { main_part: source.source.main_part },
    paragraphs: [{ paragraph_id: paragraphID, applied_styles: [], properties: {}, paragraph_mark_properties: {} }],
    runs: [], tables: [], fonts: [], diagnostics: [],
  };

  it("applies the host default size to a paragraph the separator story carries", () => {
    const result = projectNativeDocxAbsentFontSizesV1(document(true), layout, [carriedFact()], { kind: "host-default-size-v1", half_points: DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1["sizeless-document-defaults"] }, "sizeless-document-defaults");
    expect(result.applied.map((entry) => [entry.scope_id, entry.chosen_half_points])).toEqual([[paragraphID, 20]]);
    expect(result.resolved.paragraphs[0]!.paragraph_mark_properties.font_size_half_points).toBe(20);
  });

  it("refuses a fact whose scope no story projects", () => {
    expect(() => projectNativeDocxAbsentFontSizesV1(document(false), layout, [carriedFact()], { kind: "host-default-size-v1", half_points: 20 }, "sizeless-document-defaults"))
      .toThrow("Missing-size scope anchor does not exact-join");
  });
});

/**
 * `listWithLgl.docx` numbers every one of its paragraphs, so until the list
 * marker had a scope of its own the package produced no missing-size evidence
 * at all and refused on `missing-run-size`. Microsoft Word 16.112's own PDF
 * export writes its markers at 12 pt — the host default for a package that
 * carries no `w:docDefaults` record — at x = 72.0 pt for `ilvl=0` and
 * x = 108.0 pt for `ilvl=1`.
 */
describe("missing-size evidence for a list marker", () => {
  const source = JSON.parse(readFileSync(new URL("../../../testdata/docx-native/document-v1.json", import.meta.url), "utf8"));
  const relationshipsPart = "word/_rels/document.xml.rels";
  const numberingPart = "word/numbering.xml";
  const digest = (byte: string) => `sha256:${byte.repeat(64)}`;
  const policy = { kind: "host-default-size-v1", half_points: DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1["absent-document-defaults"] };
  const shape = "absent-document-defaults";

  function model(markerSizeHalfPoints?: number) {
    const document = structuredClone(source);
    const paragraph = document.body.blocks[0].paragraph;
    document.passthrough_parts = [
      ...document.passthrough_parts,
      { part_name: relationshipsPart, content_type: "application/vnd.openxmlformats-package.relationships+xml", byte_length: 1, sha256: digest("1"), policy: "preserve-verbatim" },
      { part_name: numberingPart, content_type: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", byte_length: 1, sha256: digest("2"), policy: "preserve-verbatim" },
    ];
    const numbering = {
      marker_id: "marker:paragraph:intro", definition_sha256: digest("0"),
      num_id: "7", abstract_num_id: "3", level: 0, start: 1,
      format: "decimal", text: "%1.", suffix: "tab", alignment: "start", never_restart: true,
      counter_value: 1, counter_values: [{ level: 0, value: 1, format: "decimal" }], resolved_text: "1.",
      label_start_twips: 0, label_end_twips: 720, text_start_twips: 720,
      marker_properties: markerSizeHalfPoints === undefined ? {} : { font_size_half_points: markerSizeHalfPoints },
    };
    const resolved = {
      protocol: "injoffice.docx.resolved-layout", version: 1,
      document_id: document.document_id, revision: document.revision,
      source_parts: { main_part: document.source.main_part, numbering_part: numberingPart },
      paragraphs: [{ paragraph_id: paragraph.id, applied_styles: [], properties: {}, paragraph_mark_properties: {}, numbering }],
      runs: [], tables: [], fonts: [], diagnostics: [],
    } as unknown as NativeDocxResolvedLayoutInputV1;
    const attest = {
      relationships_part: relationshipsPart, relationships_sha256: digest("1"),
      relationship_id: "rIdNumbering", relationship_type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
      relationship_target: "numbering.xml", part_name: numberingPart,
      content_type: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" as const, part_sha256: digest("2"),
    };
    numbering.definition_sha256 = nativeDocxResolvedNumberingDefinitionSha256V1(resolved.paragraphs[0]!.numbering!, attest.part_sha256);
    resolved.numbering_source = { ...attest, model_sha256: nativeDocxResolvedNumberingModelSha256V1(resolved.paragraphs, attest) };
    const fact = { scope_kind: "numbering-marker" as const, scope_id: paragraph.id, part_name: paragraph.anchor.part_name, path: paragraph.anchor.path, package_sha256: source.source.package_sha256 };
    return { document, resolved, fact };
  }

  it("applies the host default to the marker and re-attests the marker model", () => {
    const { document, resolved, fact } = model();
    expect(decodeNativeDocxResolvedLayout(resolved).ok).toBe(true);
    const projected = projectNativeDocxAbsentFontSizesV1(document, resolved, [fact], policy, shape);
    expect(projected.applied).toEqual([{ ...fact, chosen_half_points: 24 }]);
    expect(projected.resolved.paragraphs[0]!.numbering!.marker_properties.font_size_half_points).toBe(24);
    // marker_properties is part of the canonical marker model, so a projection
    // that left the attestation alone would hand the next stage a layout that
    // no longer decodes. The strict input is never mutated.
    expect(projected.resolved.numbering_source!.model_sha256).not.toBe(resolved.numbering_source!.model_sha256);
    expect(decodeNativeDocxResolvedLayout(projected.resolved).ok).toBe(true);
    expect(resolved.paragraphs[0]!.numbering!.marker_properties.font_size_half_points).toBeUndefined();
  });

  it("never overrides a marker size the source resolved, and needs a marker to apply to", () => {
    const sized = model(28);
    expect(() => projectNativeDocxAbsentFontSizesV1(sized.document, sized.resolved, [sized.fact], policy, shape)).toThrow("override an authored/resolved font size");
    const unnumbered = model();
    delete unnumbered.resolved.paragraphs[0]!.numbering;
    delete unnumbered.resolved.numbering_source;
    delete unnumbered.resolved.source_parts.numbering_part;
    expect(() => projectNativeDocxAbsentFontSizesV1(unnumbered.document, unnumbered.resolved, [unnumbered.fact], policy, shape)).toThrow("override an authored/resolved font size");
  });
});
