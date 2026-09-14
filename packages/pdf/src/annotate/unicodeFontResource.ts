import { CustomFontEmbedder, PDFFont, PDFHexString, PDFName, PDFString } from 'pdf-lib'
import type { PDFContext, PDFDict, PDFDocument, PDFRef } from 'pdf-lib'
import type { UnicodeRun } from './unicodeShaping.js'
import type { UnicodeCffFont } from './unicodeCffFont.js'

export interface OutlineCommand { command: string; args: number[] }
export interface EncodedUnicodeRun { run: UnicodeRun; cids: readonly number[]; encoded: PDFHexString; ink: { minX: number; minY: number; maxX: number; maxY: number }; outlines: readonly (readonly OutlineCommand[] | undefined)[] }
interface Entry { gid: number; text: string | undefined; width: number }
const cidHex = (cid: number) => cid.toString(16).padStart(4, '0').toUpperCase()
const unicodeHex = (value: string) => Array.from({ length: value.length }, (_, i) => value.charCodeAt(i).toString(16).padStart(4, '0').toUpperCase()).join('')

/** A character code identifies source semantics as well as a glyph. TrueType
 * uses matching CIDs; CFF maps semantic codes to the program's original CIDs. */
export class UnicodeFontResource extends CustomFontEmbedder {
  private entries: Entry[] = []
  private keys = new Map<string, number>()
  private cmapRef: PDFRef | undefined
  private mapRef: PDFRef | undefined
  private encodingRef: PDFRef | undefined
  private cidDictionary: PDFDict | undefined
  private current: EncodedUnicodeRun | undefined
  private pending: { entries: Entry[]; keys: Map<string, number> } | undefined

  constructor(font: CustomFontEmbedder['font'], bytes: Uint8Array, private readonly document: PDFDocument, private readonly cff?: UnicodeCffFont) {
    super(font, bytes, cff?.name)
  }

  /** Preflight into local copies before publishing a run or any resource state. */
  qualify(run: UnicodeRun): EncodedUnicodeRun {
    const entries = this.entries.slice()
    const keys = new Map(this.keys)
    const groups = new Map<number, typeof run.glyphs[number][]>()
    for (const glyph of run.glyphs) {
      const group = groups.get(glyph.cluster) ?? []
      group.push(glyph); groups.set(glyph.cluster, group)
    }
    const sources = new Map<number, string[]>()
    for (const group of groups.values()) {
      const glyph = group[0]!
      const text = run.value.slice(glyph.cluster, glyph.end)
      const characters = [...text]
      const partitionable = characters.length === group.length && group.every((g, i) => this.font.glyphForCodePoint(characters[i]!.codePointAt(0)!).id === g.id)
      sources.set(glyph.cluster, partitionable ? characters : [text])
    }
    const indices = new Map<number, number>()
    const outlines: (readonly OutlineCommand[] | undefined)[] = []
    let pathBudget = 65536
    const cids = run.glyphs.map(glyph => {
      const index = indices.get(glyph.cluster) ?? 0
      indices.set(glyph.cluster, index + 1)
      const text = sources.get(glyph.cluster)![index]
      if (text === undefined) {
        // Reader-safe continuation: no empty/unmapped text CID that extractors
        // can reinterpret as a character. Preserve the exact glyph outline.
        const commands = (this.font.getGlyph(glyph.id).path as unknown as { commands: OutlineCommand[] }).commands
        if (!Array.isArray(commands)) throw new Error('invalid embedded continuation outline')
        pathBudget -= commands.length
        if (pathBudget < 0) throw new Error('embedded appearance outline limit exceeded')
        const counts: Record<string, number> = { moveTo: 2, lineTo: 2, quadraticCurveTo: 4, bezierCurveTo: 6, closePath: 0 }
        if (commands.some(c => !c || !Array.isArray(c.args) || counts[c.command] !== c.args.length || c.args.some(n => !Number.isFinite(n) || Math.abs(n) > run.unitsPerEm * 100))) throw new Error('invalid embedded continuation outline')
        outlines.push(commands.map(c => ({ command: c.command, args: [...c.args] })))
        return 0
      }
      if (text.length > 256) throw new Error('embedded appearance cluster exceeds the 512-byte ToUnicode limit')
      outlines.push(undefined)
      const key = JSON.stringify([glyph.id, text ?? null])
      let cid = keys.get(key)
      if (cid === undefined) {
        cid = entries.length + 1
        if (cid > 65535) throw new Error('embedded appearance CID limit exceeded')
        const width = this.font.getGlyph(glyph.id).advanceWidth * this.scale
        if (!Number.isFinite(width) || width < 0) throw new Error('invalid embedded appearance glyph width')
        entries.push({ gid: glyph.id, text, width })
        keys.set(key, cid)
      }
      return cid
    })
    const ink = { minX: 0, minY: 0, maxX: run.width, maxY: 0 }
    for (const glyph of run.glyphs) {
      const outline = this.font.getGlyph(glyph.id)
      const box = outline.bbox
      if ((outline.path as unknown as { commands: OutlineCommand[] }).commands.length === 0) continue
      if (![box.minX, box.minY, box.maxX, box.maxY].every(Number.isFinite) || box.minX > box.maxX || box.minY > box.maxY) throw new Error('invalid embedded glyph ink bounds')
      ink.minX = Math.min(ink.minX, glyph.x + box.minX)
      ink.minY = Math.min(ink.minY, glyph.y + box.minY)
      ink.maxX = Math.max(ink.maxX, glyph.x + box.maxX)
      ink.maxY = Math.max(ink.maxY, glyph.y + box.maxY)
    }
    const result = { run, cids, ink, encoded: PDFHexString.of(cids.filter(cid => cid !== 0).map(cidHex).join('')), outlines }
    this.pending = { entries, keys }
    this.current = result
    return result
  }

  commit(): void {
    if (!this.pending) return
    this.entries = this.pending.entries
    this.keys = this.pending.keys
    this.pending = undefined
    this.refresh(this.document.context)
  }

  override encodeText(value: string): PDFHexString {
    if (!this.current || this.current.run.value !== value) throw new Error('unqualified embedded Unicode encoding')
    return this.current.encoded
  }
  override widthOfTextAtSize(value: string, size: number): number {
    if (!this.current || this.current.run.value !== value) throw new Error('unqualified embedded Unicode metrics')
    return this.current.run.width * size / this.current.run.unitsPerEm
  }
  async embed(): Promise<PDFFont> {
    this.commit()
    const ref = this.document.context.nextRef()
    await this.embedIntoContext(this.document.context, ref)
    // Resources are written directly and refreshed at the same stable refs.
    // No private PDFDocument font registry or existing font is mutated.
    return PDFFont.of(ref, this.document, this)
  }
  protected override isCFF(): boolean { return Boolean(this.cff) }
  protected override async embedFontStream(context: PDFContext): Promise<PDFRef> {
    if (!this.cff) return super.embedFontStream(context)
    return context.register(context.flateStream(this.cff.bytes, { Subtype: 'CIDFontType0C' }))
  }
  protected override async embedFontDict(context: PDFContext, ref?: PDFRef): Promise<PDFRef> {
    if (!this.cff) return super.embedFontDict(context, ref)
    const descendant = await this.embedCIDFontDict(context), unicode = this.embedUnicodeCmap(context)
    this.encodingRef = context.nextRef()
    this.refresh(context)
    const dictionary = context.obj({ Type: 'Font', Subtype: 'Type0', BaseFont: this.baseFontName, Encoding: this.encodingRef, DescendantFonts: [descendant], ToUnicode: unicode })
    if (ref) { context.assign(ref, dictionary); return ref }
    return context.register(dictionary)
  }
  protected override async embedCIDFontDict(context: PDFContext): Promise<PDFRef> {
    const descriptor = await this.embedFontDescriptor(context)
    if (!this.cff) this.mapRef = context.nextRef()
    this.cidDictionary = context.obj({
      Type: 'Font', Subtype: this.cff ? 'CIDFontType0' : 'CIDFontType2', BaseFont: this.baseFontName,
      CIDSystemInfo: { Registry: PDFString.of(this.cff?.registry ?? 'Adobe'), Ordering: PDFString.of(this.cff?.ordering ?? 'Identity'), Supplement: this.cff?.supplement ?? 0 },
      FontDescriptor: descriptor, CIDToGIDMap: this.mapRef, W: [],
    })
    this.refresh(context)
    return context.register(this.cidDictionary)
  }
  protected override embedUnicodeCmap(context: PDFContext): PDFRef {
    this.cmapRef = context.nextRef()
    this.refresh(context)
    return this.cmapRef
  }
  private refresh(context: PDFContext): void {
    if (this.mapRef) {
      const bytes = new Uint8Array((this.entries.length + 1) * 2)
      this.entries.forEach((entry, i) => { bytes[(i + 1) * 2] = entry.gid >>> 8; bytes[(i + 1) * 2 + 1] = entry.gid & 255 })
      context.assign(this.mapRef, context.flateStream(bytes))
    }
    if (this.cff) {
      const widths = new Map<number, number>()
      for (const entry of this.entries) widths.set(this.cff.glyphCids[entry.gid]!, entry.width)
      this.cidDictionary?.set(PDFName.of('W'), context.obj([...widths].sort((a, b) => a[0] - b[0]).flatMap(([cid, width]) => [cid, [width]])))
    } else this.cidDictionary?.set(PDFName.of('W'), context.obj(this.entries.length ? [1, this.entries.map(entry => entry.width)] : []))
    if (this.encodingRef && this.cff) {
      const mappings = this.entries.map((entry, index) => `<${cidHex(index + 1)}> ${this.cff!.glyphCids[entry.gid]}`), sections: string[] = []
      for (let i = 0; i < mappings.length; i += 100) { const batch = mappings.slice(i, i + 100); sections.push(`${batch.length} begincidchar\n${batch.join('\n')}\nendcidchar`) }
      context.assign(this.encodingRef, context.flateStream(`/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry ${PDFString.of(this.cff.registry)} /Ordering ${PDFString.of(this.cff.ordering)} /Supplement ${this.cff.supplement} >> def\n/CMapName /InjofficeCff def\n/CMapType 1 def\n/WMode 0 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${sections.join('\n')}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend`))
    }
    if (this.cmapRef) {
      const mappings = this.entries.flatMap((entry, index) => entry.text === undefined ? [] : [`<${cidHex(index + 1)}> <${unicodeHex(entry.text)}>`])
      const sections: string[] = []
      for (let i = 0; i < mappings.length; i += 100) {
        const batch = mappings.slice(i, i + 100)
        sections.push(`${batch.length} beginbfchar\n${batch.join('\n')}\nendbfchar`)
      }
      context.assign(this.cmapRef, context.flateStream(`/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /InjofficeUnicode def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${sections.join('\n')}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend`))
    }
  }
}
