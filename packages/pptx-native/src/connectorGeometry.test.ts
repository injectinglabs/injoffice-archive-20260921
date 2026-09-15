import { readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it } from 'vitest'
import { validateNativePptx } from './index'
import type { NativeEvaluatedGeometry, NativePptxDeck } from './types'

const fixtureRoot = resolve(import.meta.dirname, '../../../go/pptxpatch/testdata/native-contract')

function connectorDeck() {
  const deck = JSON.parse(readFileSync(resolve(fixtureRoot, 'valid/parsed-full.json'), 'utf8')) as NativePptxDeck
  const connector = deck.slides[0]!.elements.find((element) => element.kind === 'connector')!
  if (connector.kind !== 'connector') throw new Error('connector missing')
  return { deck, connector }
}

const elbow: NativeEvaluatedGeometry = {
  profile: 'drawingml-paths-v1',
  textRect: { x: 0, y: 0, cx: 1000000, cy: 500000 },
  paths: [{ fillMode: 'none', stroke: true, commands: [{ kind: 'moveTo', x: 0, y: 0 }, { kind: 'lineTo', x: 1000000, y: 0 }, { kind: 'lineTo', x: 1000000, y: 500000 }] }],
}

describe('connector preset geometry contract', () => {
  it('admits read-only evaluated connector geometry with a source affine and refuses editable or legacy-flagged variants', () => {
    const { deck, connector } = connectorDeck()
    const codes = () => { const result = validateNativePptx(deck); return result.ok ? [] : result.issues.map((issue) => issue.code) }
    connector.geometry = structuredClone(elbow)
    connector.compatibility = { status: 'preserveOnly', diagnostics: [{ severity: 'warning', code: 'pptx.connector-preset-preview', message: 'catalog preview' }] }
    connector.transform = { ...connector.transform, rotationAngle: 16200000, flipV: true }
    delete connector.flipH
    expect(codes()).toEqual([])

    connector.flipH = true
    expect(codes()).toContain('native.connectorGeometry')
    delete connector.flipH

    connector.compatibility = { status: 'editable', diagnostics: [] }
    expect(codes()).toContain('native.geometryAuthority')
    connector.compatibility = { status: 'preserveOnly', diagnostics: [] }

    connector.geometry = { ...structuredClone(elbow), paths: [{ ...elbow.paths[0]!, commands: [{ kind: 'lineTo', x: 1, y: 1 }] }] }
    expect(validateNativePptx(deck).ok).toBe(false)

    ;(connector as { preset?: string }).preset = 'rect'
    connector.geometry = structuredClone(elbow)
    expect(validateNativePptx(deck).ok).toBe(false)
  })

  it('keeps the exact straight connector contract unchanged without geometry', () => {
    const { deck } = connectorDeck()
    expect(validateNativePptx(deck).ok).toBe(true)
  })
})
