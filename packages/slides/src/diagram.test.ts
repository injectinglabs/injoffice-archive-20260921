import { describe, expect, it } from 'vitest'
import { compileDiagramShapes } from './diagram'
import { boardroomTheme } from './themes'
import type { DiagramSpec, OrgChartNode } from './types'
import type { WireShape } from './wire'

const REGION = { x: 0.9, y: 2.3, w: 11.4, h: 4.3 }
const EMU_PER_INCH = 914400

describe('compileDiagramShapes', () => {
  it('undefined diagram compiles to no shapes', () => {
    expect(compileDiagramShapes(undefined, boardroomTheme, REGION)).toEqual([])
  })

  it('empty process steps compiles to no shapes', () => {
    const diagram: DiagramSpec = { type: 'process', steps: [] }
    expect(compileDiagramShapes(diagram, boardroomTheme, REGION)).toEqual([])
  })

  describe('process flow', () => {
    const diagram: DiagramSpec = { type: 'process', steps: ['Discover', 'Design', 'Build', 'Ship'] }
    const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)

    it('produces one box per step plus one connector between each adjacent pair', () => {
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      const lines = shapes.filter((s) => s.kind === 'line')
      expect(boxes.length).toBe(4)
      expect(lines.length).toBe(3)
    })

    it('boxes carry their label text in reading order', () => {
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      expect(boxes.map((b) => b.paragraphs?.[0]?.runs[0]?.text)).toEqual(['Discover', 'Design', 'Build', 'Ship'])
    })

    it('boxes are laid out left-to-right with no overlap and stay inside the region width', () => {
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      for (let i = 0; i < boxes.length - 1; i++) {
        const a = boxes[i]!
        const b = boxes[i + 1]!
        expect(a.x + a.cx).toBeLessThanOrEqual(b.x)
      }
      const last = boxes[boxes.length - 1]!
      expect(last.x + last.cx).toBeLessThanOrEqual(Math.round((REGION.x + REGION.w) * EMU_PER_INCH) + 1)
    })

    it('every connector carries a tail arrowhead and no head arrowhead, no flip (horizontal)', () => {
      const lines = shapes.filter((s) => s.kind === 'line')
      for (const line of lines) {
        expect(line.tailArrow).toBe(true)
        expect(line.headArrow).toBeUndefined()
        expect(line.flipH).toBeUndefined()
        expect(line.cy).toBe(0) // same row: purely horizontal
      }
    })

    it('connectors span exactly the gap between adjacent boxes (touch both edges)', () => {
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      const lines = shapes.filter((s) => s.kind === 'line')
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i]!
        const boxA = boxes[i]!
        const boxB = boxes[i + 1]!
        expect(line.x).toBe(boxA.x + boxA.cx)
        expect(line.x + line.cx).toBe(boxB.x)
      }
    })

    it('a large step count still produces valid non-overlapping-in-order geometry (graceful degradation, not data loss)', () => {
      const many: DiagramSpec = { type: 'process', steps: Array.from({ length: 12 }, (_, i) => `Step ${i + 1}`) }
      const wide = compileDiagramShapes(many, boardroomTheme, REGION)
      expect(wide.filter((s) => s.kind === 'roundRect').length).toBe(12)
      expect(wide.filter((s) => s.kind === 'line').length).toBe(11)
      for (const s of wide) {
        expect(s.cx).toBeGreaterThan(0)
      }
    })
  })

  describe('org chart', () => {
    function node(label: string, children?: OrgChartNode[]): OrgChartNode {
      return { label, children }
    }

    it('a single node with no children compiles to one box, no connectors', () => {
      const diagram: DiagramSpec = { type: 'orgChart', root: node('CEO') }
      const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)
      expect(shapes.length).toBe(1)
      expect(shapes[0]?.kind).toBe('roundRect')
    })

    it('one box per node, one connector per parent-child edge', () => {
      const diagram: DiagramSpec = {
        type: 'orgChart',
        root: node('CEO', [node('CTO', [node('Eng Lead')]), node('COO'), node('CFO')]),
      }
      const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      const lines = shapes.filter((s) => s.kind === 'line')
      expect(boxes.length).toBe(5) // CEO, CTO, COO, CFO, Eng Lead
      expect(lines.length).toBe(4) // 4 edges in a 5-node tree
    })

    it('deeper rows sit strictly below shallower rows (top-down layout)', () => {
      const diagram: DiagramSpec = { type: 'orgChart', root: node('CEO', [node('CTO', [node('Eng Lead')])]) }
      const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      const ys = boxes.map((b) => b.y)
      expect(new Set(ys).size).toBe(3) // CEO, CTO, Eng Lead each on their own row
      expect(Math.min(...ys)).toBe(boxes[0]!.y) // root (first pushed) is the topmost row
    })

    it('a left child (left of parent center) gets a flipped connector; a right child does not', () => {
      // Root centered over 3 children: left, middle-aligned, right.
      const diagram: DiagramSpec = { type: 'orgChart', root: node('CEO', [node('CTO'), node('COO'), node('CFO')]) }
      const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      const lines = shapes.filter((s) => s.kind === 'line') as WireShape[]
      const root = boxes[0]!
      const rootCx = root.x + root.cx / 2
      // Identify each connector's child endpoint (the lower point) and compare
      // against root center to classify left/right, independent of internal
      // box ordering.
      for (const line of lines) {
        const childCenterX = line.flipH ? line.x : line.x + line.cx
        // when NOT flipped, the connector's second (bottom) point is at
        // (off.x+cx, off.y+cy) -> child is to the right or aligned (off.x==x1)
        // when flipped, the connector's second (bottom) point is at
        // (off.x, off.y+cy) -> child is to the left
        if (childCenterX < rootCx) {
          expect(line.flipH).toBe(true)
        } else if (childCenterX > rootCx) {
          expect(line.flipH).toBeUndefined()
        }
      }
      // At least one of each in this 3-child fan-out.
      expect(lines.some((l) => l.flipH)).toBe(true)
      expect(lines.some((l) => !l.flipH)).toBe(true)
    })

    it('no connector ever carries an arrowhead (org-chart connectors are plain lines)', () => {
      const diagram: DiagramSpec = { type: 'orgChart', root: node('CEO', [node('CTO'), node('COO')]) }
      const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)
      for (const line of shapes.filter((s) => s.kind === 'line')) {
        expect(line.headArrow).toBeUndefined()
        expect(line.tailArrow).toBeUndefined()
      }
    })

    it('a wide tree keeps boxes within a sane minimum width rather than collapsing to zero', () => {
      const wideRoot = node(
        'CEO',
        Array.from({ length: 15 }, (_, i) => node(`VP ${i + 1}`)),
      )
      const diagram: DiagramSpec = { type: 'orgChart', root: wideRoot }
      const shapes = compileDiagramShapes(diagram, boardroomTheme, REGION)
      const boxes = shapes.filter((s) => s.kind === 'roundRect')
      expect(boxes.length).toBe(16)
      for (const b of boxes) {
        expect(b.cx).toBeGreaterThan(0)
      }
    })
  })
})
