import { decodeNativeDocxDocument } from "./nativeContract.js";
import { decodeNativeDocxResolvedLayout } from "./nativeResolvedLayout.js";

export const DOCX_ABSENT_FONT_SIZE_WARNING =
  "Approximate read-only preview: source-absent font sizes use the explicitly selected 11 pt host default; this is not an authored size or Microsoft Word default." as const;
export interface NativeDocxAbsentFontSizeV1 {
  scope_kind: "run" | "paragraph-mark";
  scope_id: string;
  part_name: string;
  path: string;
  package_sha256: string;
}
export interface NativeDocxHostDefaultSizePolicyV1 {
  kind: "host-default-size-v1";
  half_points: 22;
}
export interface NativeDocxApproximatedFontSizeV1
  extends NativeDocxAbsentFontSizeV1 {
  chosen_half_points: 22;
}
export function validNativeDocxHostDefaultSizePolicyV1(
  value: unknown,
): value is NativeDocxHostDefaultSizePolicyV1 {
  if (!value || typeof value !== "object") return false;
  const p = value as Record<string, unknown>;
  return (
    Object.keys(p).sort().join(",") === "half_points,kind" &&
    p.kind === "host-default-size-v1" &&
    p.half_points === 22
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
): value is NativeDocxApproximatedFontSizeV1[] {
  if (
    !Array.isArray(value) ||
    value.length !== source.length ||
    value.length === 0 ||
    value.some(
      (f) =>
        !f ||
        typeof f !== "object" ||
        f.chosen_half_points !== 22 ||
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
) {
  if (!validNativeDocxHostDefaultSizePolicyV1(policy))
    throw new TypeError(
      "Missing-size approximation requires the explicit 11 pt host policy",
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
      ((s.note_role === "separator" && s.native_story_id === "-1") ||
        (s.note_role === "continuation-separator" && s.native_story_id === "0")) &&
      s.blocks.length === 1 &&
      s.blocks[0]?.paragraph?.runs.length === 0 &&
      !document.value.unsupported.some((d) => d.scope_id === s.id || d.anchor?.part_name === s.part_name)
    ).flatMap((s) => s.blocks),
  ].flatMap((block) => block.paragraph ? [block.paragraph] : []);
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
