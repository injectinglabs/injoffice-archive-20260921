import { describe, expect, it } from 'vitest'
import { agentFormatFromTool, agentHref, isDesignSystemHash, isDocsHash, parseAgentTool, parseSheetsView, parseSurface, surfaceHref, SURFACES } from './route'

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

  it('serves the proposed design system without colliding with guides or the DOCX lab', () => {
    expect(isDesignSystemHash('#/design-system')).toBe(true)
    expect(isDesignSystemHash('#/design-system?section=buttons')).toBe(true)
    expect(isDesignSystemHash('#/guides')).toBe(false)
    expect(isDesignSystemHash('#/docs')).toBe(false)
    expect(parseSurface('#/design-system')).toBe('overview')
  })

  it('deep-links to a spreadsheet proof without changing the surface parser', () => {
    expect(parseSurface('#/sheets?view=native')).toBe('sheets')
    expect(parseSheetsView('#/sheets?view=native')).toBe('native')
    expect(parseSheetsView('#/sheets?view=tools')).toBe('tools')
    expect(parseSheetsView('#/sheets?view=unknown')).toBe('editor')
  })

  it('deep-links the AI change-set demo to sheets, docs, slides, and pdf', () => {
    expect(parseSurface('#/agent?format=docs')).toBe('agent')
    expect(parseAgentTool('#/agent')).toBe('sheets')
    expect(parseAgentTool('#/agent?format=sheets')).toBe('sheets')
    expect(parseAgentTool('#/agent?format=docs')).toBe('docs')
    expect(parseAgentTool('#/agent?format=slides')).toBe('slides')
    expect(parseAgentTool('#/agent?format=pdf')).toBe('pdf')
    expect(parseAgentTool('#/agent?format=pptx')).toBe('slides')
    expect(agentFormatFromTool('sheets')).toBe('xlsx')
    expect(agentFormatFromTool('docs')).toBe('docx')
    expect(agentFormatFromTool('slides')).toBe('pptx')
    expect(agentFormatFromTool('pdf')).toBe('pdf')
    expect(agentHref('sheets')).toBe('#/agent?format=sheets')
    expect(agentHref('docs')).toBe('#/agent?format=docs')
    expect(agentHref('slides')).toBe('#/agent?format=slides')
    expect(agentHref('pdf')).toBe('#/agent?format=pdf')
  })
})
