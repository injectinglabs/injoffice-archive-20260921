import { decodeNativeDocxDocument } from "./nativeContract.js";
import { decodeNativeDocxResolvedLayout } from "./nativeResolvedLayout.js";

/** The single definition of the read-only host default size, one value per
 * proven source shape. Both numbers were read directly out of the `Tf`
 * operators of Microsoft Word 16.112's own PDF exports of corpus packages that
 * state no size — Word lays text out on a 1/300 in grid, so its 50 units are
 * 12 pt and its 42 units are 10 pt. A package that carries no `w:docDefaults`
 * record at all is laid out at 12 pt; one whose `w:docDefaults` exists and
 * states no `w:sz` is laid out at 10 pt. They are NOT interchangeable, so no
 * single host default can match Word. Nothing else may spell these numbers:
 * the policy type, the validators and the disclosure all derive from here. */
export const DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1 = {
  /** No `w:docDefaults` record in the package (ECMA-376 17.7.2 makes both the
   * styles part and the element optional). Word 16.112 paints 12 pt. */
  "absent-document-defaults": 24,
  /** A `w:docDefaults` record that states no `w:sz`. Word 16.112 paints 10 pt. */
  "sizeless-document-defaults": 20,
} as const;
export type NativeDocxAbsentDefaultSizeShapeV1 =
  keyof typeof DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1;
export type NativeDocxHostDefaultSizeHalfPointsV1 =
  (typeof DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1)[NativeDocxAbsentDefaultSizeShapeV1];
const hostDefaultHalfPoints: readonly number[] = Object.values(
  DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1,
);
export function validNativeDocxAbsentDefaultSizeShapeV1(
  value: unknown,
): value is NativeDocxAbsentDefaultSizeShapeV1 {
  return (
    typeof value === "string" &&
    Object.hasOwn(DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1, value)
  );
}
export const DOCX_ABSENT_FONT_SIZE_WARNING =
  `Approximate read-only preview: source-absent font sizes use an explicitly selected host default of ${[...hostDefaultHalfPoints].sort((a, b) => a - b).map((halfPoints) => halfPoints / 2).join(" or ")} pt, selected per proven source shape from Microsoft Word 16.112 references; this is not an authored size.` as const;
export interface NativeDocxAbsentFontSizeV1 {
  scope_kind: "run" | "paragraph-mark";
  scope_id: string;
  part_name: string;
  path: string;
  package_sha256: string;
}
export interface NativeDocxHostDefaultSizePolicyV1 {
  kind: "host-default-size-v1";
  half_points: NativeDocxHostDefaultSizeHalfPointsV1;
}
export interface NativeDocxApproximatedFontSizeV1
  extends NativeDocxAbsentFontSizeV1 {
  chosen_half_points: NativeDocxHostDefaultSizeHalfPointsV1;
}
export function validNativeDocxHostDefaultSizePolicyV1(
  value: unknown,
): value is NativeDocxHostDefaultSizePolicyV1 {
  if (!value || typeof value !== "object") return false;
  const p = value as Record<string, unknown>;
  return (
    Object.keys(p).sort().join(",") === "half_points,kind" &&
    p.kind === "host-default-size-v1" &&
    typeof p.half_points === "number" &&
    hostDefaultHalfPoints.includes(p.half_points)
  );
}
export function validNativeDocxAbsentFontSizesV1(
  value: unknown,
  packageSHA256: string,
): value is NativeDocxAbsentFontSizeV1[] {
  if (!Array.isArray(value) || value.length > 1000) return false;
  const ids = new Set<string>();
  for (const fact of value) {
    if (
      !fact ||
      typeof fact !== "object" ||
      Object.keys(fact).sort().join(",") !==
        "package_sha256,part_name,path,scope_id,scope_kind" ||
      !["run", "paragraph-mark"].includes(fact.scope_kind) ||
      typeof fact.scope_id !== "string" ||
      !/^[A-Za-z0-9][A-Za-z0-9._:-]{0,255}$/.test(fact.scope_id) ||
      fact.package_sha256 !== packageSHA256 ||
      !/^sha256:[0-9a-f]{64}$/.test(fact.package_sha256) ||
      typeof fact.part_name !== "string" ||
      fact.part_name.length < 1 ||
      fact.part_name.length > 1024 ||
      fact.part_name.startsWith("/") ||
      /[\\\u0000-\u001f\u007f]/.test(fact.part_name) ||
      fact.part_name
        .split("/")
        .some((s: string) => !s || s === "." || s === "..") ||
      typeof fact.path !== "string" ||
      fact.path.length > 4096 ||
      !fact.path.startsWith("/w:") ||
      /[\u0000-\u001f\u007f]/.test(fact.path)
    )
      return false;
    const id = `${fact.scope_kind}:${fact.scope_id}`;
    if (ids.has(id)) return false;
    ids.add(id);
  }
  return true;
}
export function validNativeDocxApproximatedFontSizesV1(
  value: unknown,
  source: readonly NativeDocxAbsentFontSizeV1[],
  packageSHA256: string,
  shape: unknown,
): value is NativeDocxApproximatedFontSizeV1[] {
  // Every entry must carry the one host value declared for the shape the
  // source evidence proved; the other shape's Word-derived size is a forgery.
  if (!validNativeDocxAbsentDefaultSizeShapeV1(shape)) return false;
  const chosen = DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1[shape];
  if (
    !Array.isArray(value) ||
    value.length !== source.length ||
    value.length === 0 ||
    value.some(
      (f) =>
        !f ||
        typeof f !== "object" ||
        f.chosen_half_points !== chosen ||
        Object.keys(f).sort().join(",") !==
          "chosen_half_points,package_sha256,part_name,path,scope_id,scope_kind",
    )
  )
    return false;
  const facts = value.map(({ chosen_half_points: _, ...fact }) => fact);
  return (
    validNativeDocxAbsentFontSizesV1(facts, packageSHA256) &&
    facts.every((fact, index) =>
      Object.keys(fact).every(
        (key) =>
          fact[key as keyof NativeDocxAbsentFontSizeV1] ===
          source[index]?.[key as keyof NativeDocxAbsentFontSizeV1],
      ),
    )
  );
}

/** @internal Applies an explicit host choice only to proven source omissions.
 * All original unsupported diagnostics survive in the private projection. */
export function projectNativeDocxAbsentFontSizesV1(
  documentValue: unknown,
  resolvedValue: unknown,
  facts: readonly NativeDocxAbsentFontSizeV1[],
  policy: unknown,
  shape: unknown,
) {
  // The one place a host value is bound to the source shape it was measured
  // for. A policy carrying the other shape's size is refused rather than
  // applied, so the two Word-derived numbers cannot be swapped downstream.
  if (
    !validNativeDocxAbsentDefaultSizeShapeV1(shape) ||
    !validNativeDocxHostDefaultSizePolicyV1(policy) ||
    policy.half_points !== DOCX_HOST_DEFAULT_SIZE_HALF_POINTS_V1[shape]
  )
    throw new TypeError(
      "Missing-size approximation requires the host size policy declared for the proven source shape",
    );
  const document = decodeNativeDocxDocument(documentValue),
    layout = decodeNativeDocxResolvedLayout(resolvedValue);
  if (
    !document.ok ||
    !layout.ok ||
    document.value.document_id !== layout.value.document_id ||
    document.value.revision !== layout.value.revision ||
    !validNativeDocxAbsentFontSizesV1(
      facts,
      document.value.source.package_sha256,
    )
  )
    throw new TypeError("Missing-size source evidence does not exact-join");
  const paragraphs = [
    ...document.value.body.blocks,
    ...document.value.headers.flatMap((s) => s.blocks),
    ...document.value.footers.flatMap((s) => s.blocks),
    // Raw sentinel qualification is performed by the source producer. Retain
    // the decoded reserved identity, empty content and clean-part boundaries.
    ...document.value.notes.filter((s) =>
      (s.note_role === "separator" || s.note_role === "continuation-separator") &&
      s.blocks.length === 1 &&
      s.blocks[0]?.paragraph?.runs.length === 0 &&
      !document.value.unsupported.some((d) => d.scope_id === s.id || d.anchor?.part_name === s.part_name)
    ).flatMap((s) => s.blocks),
  ].flatMap((block) =>
    block.paragraph
      ? [block.paragraph]
      : (block.table?.rows ?? []).flatMap((row) =>
          row.cells.flatMap((cell) => cell.paragraphs),
        ),
  );
  const resolved = structuredClone(layout.value);
  const applied: NativeDocxApproximatedFontSizeV1[] = [];
  for (const fact of facts) {
    const candidates =
      fact.scope_kind === "paragraph-mark"
        ? paragraphs.filter((p) => p.id === fact.scope_id)
        : paragraphs.flatMap((p) =>
            p.runs.filter((r) => r.id === fact.scope_id),
          );
    if (
      candidates.length !== 1 ||
      candidates[0]!.anchor.part_name !== fact.part_name ||
      candidates[0]!.anchor.path !== fact.path
    )
      throw new TypeError("Missing-size scope anchor does not exact-join");
    const targets =
      fact.scope_kind === "paragraph-mark"
        ? resolved.paragraphs
            .filter((p) => p.paragraph_id === fact.scope_id)
            .map((p) => p.paragraph_mark_properties)
        : resolved.runs
            .filter((r) => r.run_id === fact.scope_id)
            .map((r) => r.properties);
    if (
      targets.length !== 1 ||
      !targets[0] ||
      targets[0].font_size_half_points !== undefined
    )
      throw new TypeError(
        "Host size policy cannot override an authored/resolved font size",
      );
    targets[0].font_size_half_points = policy.half_points;
    applied.push({ ...fact, chosen_half_points: policy.half_points });
  }
  return { resolved, applied };
}
