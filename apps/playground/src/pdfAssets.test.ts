import { describe, expect, it } from 'vitest'
import { pdfAssets, pdfResourceFiles } from '../pdfAssets'

describe('self-hosted PDF assets', () => {
  it('includes packed maps, fonts, decoders and license files from one installed version', () => {
    const files = pdfResourceFiles()
    const names = [...files.keys()]
    expect(new Set(names.map(name => name.split('/')[1])).size).toBe(1)
    for (const suffix of ['cmaps/Adobe-Japan1-UCS2.bcmap', 'standard_fonts/LiberationSans-Regular.ttf', 'wasm/openjpeg.wasm', 'wasm/qcms_bg.wasm', 'wasm/LICENSE_OPENJPEG']) {
      expect(names.some(name => name.endsWith('/' + suffix))).toBe(true)
    }
    expect(names.every(name => !name.includes('..'))).toBe(true)
    expect([...files.values()].every(bytes => bytes.length > 0)).toBe(true)
  })

  it('emits host options under a non-root base without loading document globals at module import', () => {
    const plugin = pdfAssets()
    const configure = plugin.configResolved as (config: { base: string }) => void
    const load = plugin.load as (id: string) => string
    configure({ base: '/injoffice/' })
    const source = load('\0virtual:injoffice-pdf-resources')
    expect(source).toContain('export function getPdfLoadOptions()')
    expect(source).toContain('/injoffice/assets/pdfjs-')
    expect(source).toContain('useSystemFonts: false')
    expect(source).not.toContain('https://')
  })
})
