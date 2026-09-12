import type { NativeDocxTableBordersV1 } from "./nativeContract.js";

export const DOCX_AUTO_BORDER_POLICY =
  "auto-border-on-qualified-white-v1" as const;
export const DOCX_AUTO_BORDER_WARNING =
  "Approximate read-only preview: automatic table borders use black against a source-qualified white preview background; this is not Microsoft Word color fidelity." as const;
export const DOCX_AUTO_BORDER_EDGES = [
  "top",
  "right",
  "bottom",
  "left",
  "inside_horizontal",
  "inside_vertical",
] as const;
export type NativeDocxAutomaticBorderEdgeV1 =
  (typeof DOCX_AUTO_BORDER_EDGES)[number];
export interface NativeDocxAutomaticBorderDiagnosticV1 {
  code: string;
  scope_id: string;
  part_name: string;
  path: string;
}
/** Native source-derived evidence; it does not alter the strict border model. */
export interface NativeDocxAutomaticBorderEvidenceV1 {
  policy: typeof DOCX_AUTO_BORDER_POLICY;
  read_only: true;
  package_sha256: string;
  page_background: "absent-on-white-preview";
  background_rgb: "FFFFFF";
  source_part: string;
  source_path: string;
  source_sha256: string;
  borders: NativeDocxTableBordersV1;
  automatic_edges: NativeDocxAutomaticBorderEdgeV1[];
  cell_ids: string[];
  source_diagnostics: NativeDocxAutomaticBorderDiagnosticV1[];
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}
function keys(
  value: Record<string, unknown>,
  expected: readonly string[],
): boolean {
  return Object.keys(value).sort().join(",") === [...expected].sort().join(",");
}
const ID = /^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/;
const HASH = /^sha256:[0-9a-f]{64}$/;
const safeText = (value: unknown, max: number): value is string =>
  typeof value === "string" &&
  value.length > 0 &&
  value.length <= max &&
  !/[\u0000-\u001f\u007f]/.test(value);
const part = (value: unknown): value is string =>
  safeText(value, 1024) &&
  !value.startsWith("/") &&
  !value.includes("\\") &&
  value
    .split("/")
    .every((segment) => segment !== "" && segment !== "." && segment !== "..");

/** Pure, browser-safe structural validation. Source joins happen separately. */
export function validNativeDocxAutomaticBorderEvidenceV1(
  value: unknown,
): value is NativeDocxAutomaticBorderEvidenceV1 {
  if (
    !record(value) ||
    !keys(value, [
      "policy",
      "read_only",
      "package_sha256",
      "page_background",
      "background_rgb",
      "source_part",
      "source_path",
      "source_sha256",
      "borders",
      "automatic_edges",
      "cell_ids",
      "source_diagnostics",
    ])
  )
    return false;
  if (
    value.policy !== DOCX_AUTO_BORDER_POLICY ||
    value.read_only !== true ||
    value.page_background !== "absent-on-white-preview" ||
    value.background_rgb !== "FFFFFF" ||
    typeof value.package_sha256 !== "string" ||
    !HASH.test(value.package_sha256) ||
    typeof value.source_sha256 !== "string" ||
    !HASH.test(value.source_sha256) ||
    !part(value.source_part) ||
    !safeText(value.source_path, 4096) ||
    !value.source_path.endsWith("/w:tblBorders[1]")
  )
    return false;
  if (
    !Array.isArray(value.cell_ids) ||
    value.cell_ids.length < 1 ||
    value.cell_ids.length > 10_000 ||
    value.cell_ids.some((id) => typeof id !== "string" || !ID.test(id)) ||
    new Set(value.cell_ids).size !== value.cell_ids.length
  )
    return false;
  if (
    !Array.isArray(value.automatic_edges) ||
    value.automatic_edges.length < 1 ||
    value.automatic_edges.length > 6 ||
    value.automatic_edges.some(
      (edge) => !DOCX_AUTO_BORDER_EDGES.includes(edge),
    ) ||
    new Set(value.automatic_edges).size !== value.automatic_edges.length ||
    !record(value.borders) ||
    Object.keys(value.borders).length < 1 ||
    Object.keys(value.borders).some(
      (edge) =>
        !DOCX_AUTO_BORDER_EDGES.includes(
          edge as NativeDocxAutomaticBorderEdgeV1,
        ),
    )
  )
    return false;
  for (const [edge, border] of Object.entries(value.borders)) {
    if (
      !record(border) ||
      !keys(
        border,
        border.style === "none"
          ? ["style", "size_eighth_points"]
          : ["style", "size_eighth_points", "color_rgb"],
      )
    )
      return false;
    if (border.style === "none") {
      if (
        border.size_eighth_points !== 0 ||
        Object.is(border.size_eighth_points, -0) ||
        value.automatic_edges.includes(edge)
      )
        return false;
    } else if (
      border.style !== "single" ||
      !Number.isSafeInteger(border.size_eighth_points) ||
      (border.size_eighth_points as number) < 1 ||
      (border.size_eighth_points as number) > 768 ||
      typeof border.color_rgb !== "string" ||
      !/^[0-9A-F]{6}$/.test(border.color_rgb) ||
      (value.automatic_edges.includes(edge) && border.color_rgb !== "000000")
    )
      return false;
  }
  if (
    value.automatic_edges.some(
      (edge) => !(edge in (value.borders as Record<string, unknown>)),
    )
  )
    return false;
  if (
    !Array.isArray(value.source_diagnostics) ||
    value.source_diagnostics.length < 1 ||
    value.source_diagnostics.length > 2
  )
    return false;
  const diagnostics = new Set<string>();
  for (const diagnostic of value.source_diagnostics) {
    if (
      !record(diagnostic) ||
      !keys(diagnostic, ["code", "scope_id", "part_name", "path"]) ||
      !["UNMODELED_TABLE_PROPERTY", "TABLE_STYLE_EFFECTS_PRESERVED"].includes(
        diagnostic.code as string,
      ) ||
      typeof diagnostic.scope_id !== "string" ||
      !ID.test(diagnostic.scope_id) ||
      !part(diagnostic.part_name) ||
      !safeText(diagnostic.path, 4096)
    )
      return false;
    const key = JSON.stringify([
      diagnostic.code,
      diagnostic.scope_id,
      diagnostic.part_name,
      diagnostic.path,
    ]);
    if (diagnostics.has(key)) return false;
    diagnostics.add(key);
  }
  return true;
}
