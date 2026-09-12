import { describe, expect, it } from "vitest";
import {
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
  it("requires an explicit fixed host choice and full retained source coverage", () => {
    expect(
      validNativeDocxHostDefaultSizePolicyV1({
        kind: "host-default-size-v1",
        half_points: 22,
      }),
    ).toBe(true);
    for (const policy of [
      undefined,
      { kind: "host-default-size-v1", half_points: 24 },
      { kind: "word-default", half_points: 22 },
      { kind: "host-default-size-v1", half_points: 22, exact: true },
    ])
      expect(validNativeDocxHostDefaultSizePolicyV1(policy)).toBe(false);
    const applied = [{ ...fact, chosen_half_points: 22 }];
    expect(validNativeDocxApproximatedFontSizesV1(applied, [fact], hash)).toBe(
      true,
    );
    expect(validNativeDocxApproximatedFontSizesV1(applied, [], hash)).toBe(
      false,
    );
    expect(
      validNativeDocxApproximatedFontSizesV1(
        [{ ...fact, chosen_half_points: 24 }],
        [fact],
        hash,
      ),
    ).toBe(false);
    expect(
      validNativeDocxApproximatedFontSizesV1(
        applied,
        [{ ...fact, scope_id: "paragraph:other" }],
        hash,
      ),
    ).toBe(false);
  });
});
