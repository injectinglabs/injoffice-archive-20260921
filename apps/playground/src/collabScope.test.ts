import { describe, expect, it } from 'vitest'
import { applyCollabQuery, COLLAB_COPY, parseCollabFormat, parseCollabQuery } from './collabScope'

describe('collaboration demo boundaries', () => {
  it('names every missing production security boundary', () => {
    expect(COLLAB_COPY.security).toMatch(/no authentication/i)
    expect(COLLAB_COPY.security).toMatch(/authorization/i)
    expect(COLLAB_COPY.security).toMatch(/tenant isolation/i)
    expect(COLLAB_COPY.security).toMatch(/rate limiting/i)
    expect(COLLAB_COPY.security).toMatch(/localhost/i)
    expect(COLLAB_COPY.security).toMatch(/do not expose/i)
  })

  it('separates scoped completion from broader product coverage', () => {
    expect(COLLAB_COPY.completion).toMatch(/v3 checklist is complete/i)
    expect(COLLAB_COPY.completion).toMatch(/broader .* coverage remain partial/i)
  })

  it('describes the real editor and ordered protocol path', () => {
    expect(COLLAB_COPY.proof).toMatch(/real Univer sheet/i)
    expect(COLLAB_COPY.proof).toMatch(/ordered replay/i)
    expect(COLLAB_COPY.proof).toMatch(/server-sent events/i)
  })

  it('parses collaboration format query params with a sheets default', () => {
    expect(parseCollabFormat(undefined)).toBe('sheets')
    expect(parseCollabFormat(' SLIDES ')).toBe('slides')
    expect(parseCollabFormat('docs')).toBe('docs')
    expect(parseCollabFormat('pdf')).toBe('pdf')
    expect(parseCollabFormat('other')).toBe('sheets')
    expect(parseCollabQuery('?artifact=art_1&format=pdf')).toEqual({ artifact: 'art_1', format: 'pdf' })
    expect(parseCollabQuery('#/collab?format=slides&artifact=art_2')).toEqual({ artifact: 'art_2', format: 'slides' })
    expect(parseCollabQuery('http://localhost:3100/?artifact=keep#/collab')).toEqual({ artifact: 'keep', format: 'sheets' })
    expect(parseCollabQuery('http://localhost:3100/?format=pdf#/collab?format=slides')).toEqual({ artifact: '', format: 'pdf' })
  })

  it('updates format without dropping the artifact', () => {
    expect(applyCollabQuery('http://localhost:3100/?artifact=art_1#/collab', { artifact: 'art_1', format: 'slides' }))
      .toBe('/?artifact=art_1&format=slides#/collab')
    expect(applyCollabQuery('http://localhost:3100/#/collab?artifact=art_9', { artifact: 'art_9', format: 'docs' }))
      .toBe('/?artifact=art_9&format=docs#/collab')
  })
})
