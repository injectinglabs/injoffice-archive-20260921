import { describe, expect, it } from 'vitest'
import { isDocsHash, parseSheetsView, parseSurface, surfaceHref, SURFACES } from './route'

describe('playground surfaces', () => {
  it('defaults unknown hashes to the lightweight overview and round-trips hrefs', () => {
    expect(parseSurface('')).toBe('overview')
    expect(parseSurface('#/nope')).toBe('overview')
    for (const surface of SURFACES) {
      expect(parseSurface(surfaceHref(surface))).toBe(surface)
    }
  })

  it('ignores query strings and normalizes surface case', () => {
    expect(parseSurface('#/COLLAB?artifact=art_123')).toBe('collab')
    expect(parseSurface('#/PPTX-NATIVE/')).toBe('pptx-native')
  })

  it('keeps the old native link working through the spreadsheets page', () => {
    expect(parseSurface('#/native/')).toBe('sheets')
  })

  it('keeps documentation on the same 3100 hash router without colliding with the DOCX lab', () => {
    expect(isDocsHash('#/guides')).toBe(true)
    expect(isDocsHash('#/guides/charts?section=install')).toBe(true)
    expect(isDocsHash('#/docs')).toBe(false)
    expect(parseSurface('#/guides/charts')).toBe('overview')
    expect(parseSurface('#/docs')).toBe('docs')
  })

  it('deep-links to a spreadsheet proof without changing the surface parser', () => {
    expect(parseSurface('#/sheets?view=native')).toBe('sheets')
    expect(parseSheetsView('#/sheets?view=native')).toBe('native')
    expect(parseSheetsView('#/sheets?view=tools')).toBe('tools')
    expect(parseSheetsView('#/sheets?view=unknown')).toBe('editor')
  })
})
