import { describe, expect, it } from "vitest";
import {
  DOCX_AUTO_BORDER_POLICY,
  validNativeDocxAutomaticBorderEvidenceV1,
  type NativeDocxAutomaticBorderEvidenceV1,
} from "./nativeAutomaticBorderEvidenceV1.js";

function fixture(): NativeDocxAutomaticBorderEvidenceV1 {
  return {
    policy: DOCX_AUTO_BORDER_POLICY,
    read_only: true,
    package_sha256: `sha256:${"a".repeat(64)}`,
    page_background: "absent-on-white-preview",
    background_rgb: "FFFFFF",
    source_part: "word/document.xml",
    source_path: "/w:document[1]/w:body[1]/w:tbl[1]/w:tblPr[1]/w:tblBorders[1]",
    source_sha256: `sha256:${"b".repeat(64)}`,
    borders: {
      top: { style: "single", size_eighth_points: 8, color_rgb: "000000" },
      bottom: { style: "single", size_eighth_points: 8, color_rgb: "FF0000" },
    },
    automatic_edges: ["top"],
    cell_ids: ["cell:1"],
    source_diagnostics: [
      {
        code: "UNMODELED_TABLE_PROPERTY",
        scope_id: "table:1",
        part_name: "word/document.xml",
        path: "/w:document[1]/w:body[1]/w:tbl[1]/w:tblPr[1]/w:tblBorders[1]",
      },
    ],
  };
}

describe("source-bound automatic table-border evidence wire", () => {
  it("retains explicit nonautomatic color siblings and requires literal policy", () => {
    expect(validNativeDocxAutomaticBorderEvidenceV1(fixture())).toBe(true);
    for (const patch of [
      { policy: "auto-is-black" },
      { read_only: false },
      { background_rgb: "000000" },
      { page_background: "assumed-white" },
      { package_sha256: "wrong" },
      { source_part: "../document.xml" },
      { source_path: "/wrong" },
      { automatic_edges: [] },
      { automatic_edges: ["top", "top"] },
      { automatic_edges: ["left"] },
      { cell_ids: [] },
      { cell_ids: ["cell:1", "cell:1"] },
      { source_diagnostics: [] },
      { unknown: true },
    ])
      expect(
        validNativeDocxAutomaticBorderEvidenceV1({ ...fixture(), ...patch }),
      ).toBe(false);
  });

  it("enforces edge, style, numeric and cumulative evidence budgets", () => {
    for (const size of [0, -0, -1, 769, 0.5, NaN, Infinity, "8"]) {
      const value = fixture();
      Object.assign(value.borders.top!, { size_eighth_points: size });
      expect(validNativeDocxAutomaticBorderEvidenceV1(value)).toBe(false);
    }
    const wrongColor = fixture();
    wrongColor.borders.top!.color_rgb = "FFFFFF";
    expect(validNativeDocxAutomaticBorderEvidenceV1(wrongColor)).toBe(false);
    const missing = fixture();
    delete missing.borders.top;
    expect(validNativeDocxAutomaticBorderEvidenceV1(missing)).toBe(false);
    const huge = fixture();
    huge.cell_ids = Array.from(
      { length: 10001 },
      (_, index) => `cell:${index}`,
    );
    expect(validNativeDocxAutomaticBorderEvidenceV1(huge)).toBe(false);
    const duplicate = fixture();
    duplicate.source_diagnostics.push(duplicate.source_diagnostics[0]!);
    expect(validNativeDocxAutomaticBorderEvidenceV1(duplicate)).toBe(false);
    const broad = fixture();
    broad.source_diagnostics[0]!.code = "UNKNOWN_CONTENT";
    expect(validNativeDocxAutomaticBorderEvidenceV1(broad)).toBe(false);
  });
});
