import type { NativeDocxDocumentV1, NativeDocxParagraphV1, NativeDocxParagraphPropertiesV1, NativeDocxRunV1, NativeDocxRunPropertiesV1 } from '../../../packages/docs/src/nativeContract';
type StyleSource = NativeDocxDocumentV1;
function mergeParagraph(base: NativeDocxParagraphPropertiesV1, next: NativeDocxParagraphPropertiesV1 = {}) {
  const result = {...base,...next};
  if (next.first_line_twips !== undefined) delete result.hanging_twips;
  if (next.hanging_twips !== undefined) delete result.first_line_twips;
  return result;
}
const catalogs = new WeakMap<NativeDocxDocumentV1, Map<string, NonNullable<NativeDocxDocumentV1['paragraph_styles']>[number]>>();
/** Preview cascade for the modeled style subset; cycles/absent bases are never invented. */
export function paragraphAppearance(document: StyleSource, paragraph: NativeDocxParagraphV1) {
  let catalog = catalogs.get(document);
  if (!catalog) { catalog = new Map(document.paragraph_styles?.map(style => [style.id, style]) ?? []); catalogs.set(document,catalog); }
  const chain: NonNullable<NativeDocxDocumentV1['paragraph_styles']> = [], seen = new Set<string>();
  let id = paragraph.properties.paragraph_style_id ?? document.default_paragraph_style_id;
  let complete = true;
  while (id) { const style = catalog.get(id); if (!style || seen.has(id) || chain.length >= 64) { complete = false; break; } seen.add(id); chain.unshift(style); id = style.based_on; }
  let run: NativeDocxRunPropertiesV1 = { ...document.default_run_properties }, properties: NativeDocxParagraphPropertiesV1 = { ...document.default_paragraph_properties };
  if (complete) for (const style of chain) {
    properties = mergeParagraph(properties, {...style.paragraph_properties, ...(style.outline_level !== undefined ? {outline_level:style.outline_level} : {})});
    const { bold, italic, ...rest } = style.run_properties ?? {};
    // OOXML style bold/italic are toggles; direct run properties below are absolute.
    // https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.bold
    run = { ...run, ...rest, ...(bold ? {bold:!run.bold} : {}), ...(italic ? {italic:!run.italic} : {}) };
  }
  const numbering = paragraph.properties.numbering ?? properties.numbering;
  const level = document.numbering_definitions?.find(item => item.num_id === numbering?.num_id)?.levels.find(item => item.level === numbering?.level);
  const resolved = mergeParagraph(mergeParagraph(properties, level?.paragraph_properties), paragraph.properties);
  return { run, paragraph: resolved, complete };
}
export function runAppearance(document: StyleSource, paragraph: NativeDocxParagraphV1, run: NativeDocxRunV1): NativeDocxRunPropertiesV1 { return { ...paragraphAppearance(document, paragraph).run, ...run.properties }; }

/**
 * The style name Word shows in its status bar: the paragraph's own style, else the document
 * default. Unknown ids are shown as authored rather than replaced with an invented name.
 */
export function paragraphStyleName(document: NativeDocxDocumentV1, paragraph?: NativeDocxParagraphV1) {
  const id = paragraph?.properties.paragraph_style_id ?? document.default_paragraph_style_id;
  if (!id) return 'Normal';
  return document.paragraph_styles?.find(style => style.id === id)?.name ?? id;
}

/** Number only qualified lists, separately for each story, in authored paragraph order. */
export function paragraphListLabels(document: NativeDocxDocumentV1, paragraphs: NativeDocxParagraphV1[]) {
  const labels = new Map<string, {text:string;suffix:string;properties?:NativeDocxRunPropertiesV1}>();
  const counters = new Map<string, Map<number,number>>();
  for (const paragraph of paragraphs) {
    const reference = paragraphAppearance(document,paragraph).paragraph.numbering;
    if (!reference || reference.num_id === '0') continue;
    const definition = document.numbering_definitions?.find(item => item.num_id === reference.num_id);
    const level = definition?.levels.find(item => item.level === reference.level);
    if (!level || !definition) continue;
    const counts = counters.get(reference.num_id) ?? new Map<number,number>(); counters.set(reference.num_id,counts);
    counts.set(level.level, counts.has(level.level) ? counts.get(level.level)! + 1 : level.start);
    for (const key of counts.keys()) if (key > level.level) counts.delete(key);
    let qualified = true;
    const text = level.text.replace(/%([1-9])/g, (_, value:string) => {
      const index = Number(value) - 1, source = definition.levels.find(item => item.level === index);
      if (!source || index > level.level || source.format !== 'decimal') {qualified = false; return '';}
      return String(counts.get(index) ?? source.start);
    });
    if (qualified) labels.set(`${paragraph.anchor.part_name}\0${paragraph.id}`, {text,suffix:level.suffix,properties:level.run_properties});
  }
  return labels;
}

/**
 * Whether the native engine can format part of this paragraph, or part of one
 * of its runs, without rewriting anything it cannot split exactly. It mirrors
 * the engine's own refusals: the paragraph must admit the run-property patch,
 * and each run the selection can touch must be a modeled text run that is the
 * only content of a `w:r` sitting directly under the paragraph. Hyperlinked,
 * wrapped and revision-tracked runs, and a run sharing its `w:r` with a tab or
 * a break, are all out.
 */
export function canFormatRun(paragraph: NativeDocxParagraphV1, run: NativeDocxRunV1): boolean {
  // A preview built from partial inspection can omit the policy; absence is a refusal, never a grant.
  if (paragraph.edit_policy?.mode !== 'read-write' || !paragraph.edit_policy.allowed_operations.includes('properties.patch')) return false;
  if (run.kind !== 'text' || run.anchor.part_name !== paragraph.anchor.part_name) return false;
  const owner = run.anchor.path.slice(0, run.anchor.path.lastIndexOf('/'));
  const prefix = `${paragraph.anchor.path}/`;
  if (!owner.startsWith(prefix) || owner.slice(prefix.length).includes('/') || !/(^|:)r\[\d+\]$/.test(owner.slice(prefix.length))) return false;
  return paragraph.runs.filter(sibling => sibling.anchor.path.slice(0, sibling.anchor.path.lastIndexOf('/')) === owner).length === 1;
}

/** Whether a selection spanning this paragraph's runs can be formatted. */
export function canFormatParagraphRange(paragraph: NativeDocxParagraphV1): boolean {
  const text = paragraph.runs.filter(run => run.kind === 'text');
  return text.length > 0 && text.every(run => canFormatRun(paragraph, run));
}
