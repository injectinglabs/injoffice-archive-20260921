import { describe, expect, it } from 'vitest'
import { DEMOS } from './demoRegistry'
import { AGENT_TOOLS, agentHref, surfaceHref } from './route'
import { activeSectionKey, SCROLL_SECTIONS, sectionForHash, workspaceNavigationHash } from './scrollSections'
import { TOOL_WORKSPACES, workspaceHref } from './toolWorkspaces'

describe('four-tool continuous showcase', () => {
  it('starts directly with Sheets and provides only four live workspaces', () => {
    expect(SCROLL_SECTIONS.map(section => section.key)).toEqual(['sheets', 'docs', 'slides', 'pdf'])
    expect(new Set(SCROLL_SECTIONS.map(section => section.href)).size).toBe(4)
    expect(SCROLL_SECTIONS[0]?.href).toBe('#/sheets')
    expect(SCROLL_SECTIONS.every(section => section.demo?.surface === section.surface)).toBe(true)
    for (const section of SCROLL_SECTIONS) expect(sectionForHash(section.href)).toBe(section)
  })

  it('keeps every capability reachable within the correct tool rather than a separate section', () => {
    for (const workspace of TOOL_WORKSPACES) {
      for (const feature of workspace.features) expect(sectionForHash(workspaceHref(workspace.tool, feature.id)).key).toBe(workspace.tool)
    }
    for (const { tool } of AGENT_TOOLS) expect(sectionForHash(agentHref(tool)).key).toBe(tool)
    for (const demo of DEMOS) expect(sectionForHash(surfaceHref(demo.surface)).key).not.toBe('overview')
  })

  it.each([
    ['#/agent', 'sheets'], ['#/agent?format=docx', 'docs'], ['#/agent?format=pptx', 'slides'], ['#/agent?format=PDF', 'pdf'],
    ['#/native', 'sheets'], ['#/sheets?view=native', 'sheets'], ['#/charts', 'sheets'], ['#/history', 'sheets'],
    ['#/pptx-render', 'slides'], ['#/font-metrics', 'docs'], ['#/collab?format=pdf', 'pdf'],
    ['', 'sheets'], ['#/overview', 'sheets'], ['#/missing', 'sheets'],
  ])('resolves legacy bookmark %s to %s', (hash, tool) => {
    expect(sectionForHash(hash).key).toBe(tool)
  })

  it('normalizes the removed introduction without losing capability or room deep links', () => {
    for (const hash of ['', '#/', '#/overview', '#/overview?anything=true', '#/missing']) expect(workspaceNavigationHash(hash)).toBe('#/sheets')
    for (const hash of ['#/docs', '#/charts', '#/agent?format=pdf', '#/collab?format=docs&artifact=shared-room', '#/sheets?feature=native']) expect(workspaceNavigationHash(hash)).toBe(hash)
  })
})

describe('scroll section active probe', () => {
  it('selects a short final section when the visitor reaches the document end', () => {
    expect(activeSectionKey([{ key: 'slides', top: -900 }, { key: 'pdf', top: 90 }], 80, true)).toBe('pdf')
    expect(activeSectionKey([{ key: 'slides', top: -900 }, { key: 'pdf', top: Infinity }], 80, true)).toBe('slides')
  })
  const sections = [{ key: 'overview', top: -800 }, { key: 'sheets', top: 80 }, { key: 'docs', top: 1200 }]
  it('selects the last section crossing the probe without moving focus or mutating coordinates', () => {
    expect(activeSectionKey(sections, 79)).toBe('overview')
    expect(activeSectionKey(sections, 80)).toBe('sheets')
    expect(activeSectionKey(sections, 1199)).toBe('sheets')
    expect(activeSectionKey(sections, 1200)).toBe('docs')
    expect(activeSectionKey(sections, 5000)).toBe('docs')
    expect(activeSectionKey([], 80)).toBeUndefined()
    expect(activeSectionKey(sections, -1000)).toBe('overview')
  })
  it('ignores unavailable geometry and preserves immutable inputs', () => {
    const frozen = Object.freeze([Object.freeze({ key: 'ready', top: -10 }), Object.freeze({ key: 'missing', top: Infinity })])
    expect(activeSectionKey(frozen, 80)).toBe('ready')
    expect(frozen).toEqual([{ key: 'ready', top: -10 }, { key: 'missing', top: Infinity }])
  })
})
