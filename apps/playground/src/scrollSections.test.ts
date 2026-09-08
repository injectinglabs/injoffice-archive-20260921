import { describe, expect, it } from 'vitest'
import { DEMOS, DEMO_GROUPS } from './demoRegistry'
import { AGENT_TOOLS, agentHref, surfaceHref } from './route'
import { activeSectionKey, SCROLL_SECTIONS, sectionForHash } from './scrollSections'

describe('continuous showcase section contracts', () => {
  it('provides twenty distinct anchors: the overview plus nineteen live examples', () => {
    expect(SCROLL_SECTIONS).toHaveLength(20)
    expect(new Set(SCROLL_SECTIONS.map(section => section.key)).size).toBe(20)
    expect(new Set(SCROLL_SECTIONS.map(section => section.href)).size).toBe(20)
    expect(SCROLL_SECTIONS[0]).toEqual({ key: 'overview', surface: 'overview', href: '#/overview' })
    expect(SCROLL_SECTIONS.slice(1).every(section => section.demo?.surface === section.surface)).toBe(true)
    for (const section of SCROLL_SECTIONS) expect(sectionForHash(section.href)).toBe(section)
  })

  it('matches the sidebar group order and keeps all four AI formats independent', () => {
    const expected = ['overview', ...DEMO_GROUPS.flatMap(group => DEMOS.filter(demo => demo.group === group).flatMap(demo =>
      demo.surface === 'agent' ? AGENT_TOOLS.map(({ tool }) => `agent-${tool}`) : [demo.surface],
    ))]
    expect(SCROLL_SECTIONS.map(section => section.key)).toEqual(expected)
    expect(SCROLL_SECTIONS.slice(1, 5).map(section => section.key)).toEqual(['agent-sheets', 'agent-docs', 'agent-slides', 'agent-pdf'])
    for (const { tool } of AGENT_TOOLS) expect(sectionForHash(agentHref(tool))).toMatchObject({ key: `agent-${tool}`, surface: 'agent', tool })
  })

  it.each([
    ['#/agent', 'agent-sheets'],
    ['#/agent?format=xlsx', 'agent-sheets'],
    ['#/agent?format=docx', 'agent-docs'],
    ['#/agent?format=pptx', 'agent-slides'],
    ['#/agent?format=PDF', 'agent-pdf'],
    ['#/agent?format=unknown', 'agent-sheets'],
    ['#/native', 'sheets'],
    ['#/sheets?view=native', 'sheets'],
    ['#/sheets?view=tools', 'sheets'],
    ['', 'overview'],
    ['#/missing', 'overview'],
  ])('resolves shared legacy deep link %s to %s', (hash, key) => {
    expect(sectionForHash(hash).key).toBe(key)
  })

  it('resolves every ordinary surface without adding artificial subroutes', () => {
    for (const demo of DEMOS.filter(demo => demo.surface !== 'agent')) {
      expect(sectionForHash(surfaceHref(demo.surface))).toMatchObject({ key: demo.surface, surface: demo.surface, demo })
    }
  })
})

describe('scroll section active probe', () => {
  const sections = [{ key: 'overview', top: -800 }, { key: 'agent-sheets', top: 80 }, { key: 'agent-docs', top: 1200 }]

  it('uses the last section whose top has crossed the probe, including the exact boundary', () => {
    expect(activeSectionKey(sections, 79)).toBe('overview')
    expect(activeSectionKey(sections, 80)).toBe('agent-sheets')
    expect(activeSectionKey(sections, 1199)).toBe('agent-sheets')
    expect(activeSectionKey(sections, 1200)).toBe('agent-docs')
    expect(activeSectionKey(sections, 5000)).toBe('agent-docs')
  })

  it('handles empty, above-first, single-section, and overlapping section coordinates', () => {
    expect(activeSectionKey([], 80)).toBeUndefined()
    expect(activeSectionKey(sections, -1000)).toBe('overview')
    expect(activeSectionKey([{ key: 'only', top: 100 }], 0)).toBe('only')
    expect(activeSectionKey([{ key: 'first', top: 0 }, { key: 'second', top: 0 }], 0)).toBe('second')
  })

  it('ignores unavailable section geometry and does not mutate measured coordinates', () => {
    const frozen = Object.freeze([Object.freeze({ key: 'ready', top: -10 }), Object.freeze({ key: 'missing', top: Infinity })])
    expect(activeSectionKey(frozen, 80)).toBe('ready')
    expect(frozen).toEqual([{ key: 'ready', top: -10 }, { key: 'missing', top: Infinity }])
  })
})
