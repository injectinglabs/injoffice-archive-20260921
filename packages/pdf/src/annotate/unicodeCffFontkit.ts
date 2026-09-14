import type { CustomFontEmbedder } from 'pdf-lib'
import type { UnicodeCffFont } from './unicodeCffFont.js'

/** Only call on a fresh, exclusively owned @pdf-lib/fontkit1.1.1 parser.
 * Its format3 FDSelect binary search uses > instead of >= at some boundaries.
 * Normalize the private decoded view to format0 using our validated source
 * mapping before any glyph paths are cached. Never change bytes/prototypes.
 * The installed version is pinned by the dependency and adapter regression. */
export function normalizeOwnedCffFontkit(font: CustomFontEmbedder['font'], validated: UnicodeCffFont): void {
  if (!validated.sourceGlyphFDs) return // Name-keyed source has no FDSelect.
  type Selection = { version: number; fds?: number[]; ranges?: { first: number; fd: number }[]; sentinel?: number; nRanges?: number }
  type PrivateCff = { version: number; topDict: { FDSelect: Selection; FDArray: unknown[] }; fdForGlyph: unknown; privateDictForGlyph: unknown }
  const cff = (font as unknown as { 'CFF '?: PrivateCff })['CFF ']
  const fail = () => { throw new Error('unsupported fontkit CFF FDSelect parser shape') }
  if (!cff || cff.version !== 1 || typeof cff.fdForGlyph !== 'function' || typeof cff.privateDictForGlyph !== 'function' || !cff.topDict || !Array.isArray(cff.topDict.FDArray)) return fail()
  const selected = validated.sourceGlyphFDs, select = cff.topDict.FDSelect
  if (selected.length !== font.numGlyphs || selected.some(fd => !Number.isInteger(fd) || fd < 0 || fd >= cff.topDict.FDArray.length) || !select) return fail()
  if (select.version === 0) {
    if (!Array.isArray(select.fds) || select.fds.length !== selected.length || select.fds.some((fd, i) => fd !== selected[i])) return fail()
  } else if (select.version === 3) {
    const ranges = select.ranges
    if (!Array.isArray(ranges) || !ranges.length || select.nRanges !== ranges.length || select.sentinel !== selected.length || ranges[0]?.first !== 0) return fail()
    for (let i = 0; i < ranges.length; i++) {
      const range = ranges[i]!, end = ranges[i + 1]?.first ?? selected.length
      if (!Number.isInteger(range.first) || range.first < 0 || range.first >= end || end > selected.length) return fail()
      for (let gid = range.first; gid < end; gid++) if (selected[gid] !== range.fd) return fail()
    }
  } else return fail()
  const descriptor = Object.getOwnPropertyDescriptor(cff.topDict, 'FDSelect')
  if (!descriptor || !descriptor.configurable || !('value' in descriptor)) return fail()
  Object.defineProperty(cff.topDict, 'FDSelect', { value: Object.freeze({ version: 0, fds: Object.freeze([...selected]) }), enumerable: true, configurable: true })
}
