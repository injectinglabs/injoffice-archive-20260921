import { describe, expect, it } from "vitest";
import {
  DOCX_ABSENT_FONT_SIZE_WARNING,
  DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1,
  validNativeDocxAbsentDefaultSizeShapeV1,
  validNativeDocxAbsentFontSizesV1,
  validNativeDocxApproximatedFontSizesV1,
  validNativeDocxHostDefaultSizePolicyV1,
} from "./nativeAbsentFontSizeV1.js";

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
