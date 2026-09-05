import { useEffect, useMemo, useRef, useState, type ReactNode } from 'react'
import type Konva from 'konva'
import { Arrow, Ellipse, Group, Layer, Line, RegularPolygon, Rect, Stage, Star, Text, Transformer } from 'react-konva'
import { compileDeckToWire } from './compile'
import {
  MIN_SHAPE_PX,
  canvasFontFamily,
  dragResultEmu,
  fitFontSizePx,
  makeViewport,
  paragraphLine,
  ptToPx,
  regularPolygonSides,
  resizeResultEmu,
  roundRectCornerRadiusPx,
  shapeRectPx,
  wrapLineToWidth,
} from './canvasGeometry'
import type { DeckSpec } from './types'
import type { WireDeck, WireShape, WireSlide } from './wire'

// DeckCanvasView — the OOXML-pivot canvas renderer (S7.3, editing added
// S8): renders the SAME compiled Deck model (compile.ts → go/pptxpatch's
// Deck shape) the gateway's POST /v1/pptx/build turns into a real .pptx,
// directly in the browser via Konva. This does NOT call the gateway or
// pptxpatch at all — it consumes compileDeckToWire's pure output client-
// side, so a deck previews correctly even before a save/export round trip.
// Sits alongside DeckView.tsx (the existing HTML/DOM renderer): same
// controlled at/onAtChange index API, so a host can swap between them
// without restructuring state.
//
// Deliberately minimal, matching the "canvas renderer is secondary to a
// correct file format" charter this pivot was built under:
//  - One Konva <Text> node per shape's whole paragraph list, not one per
//    paragraph/run. Bullet paragraphs get a glyph line-prefix
//    (canvasGeometry.paragraphLine); Konva's own wrap="word" handles
//    reflow. Mixed per-paragraph/per-run styling (bold mid-sentence, a
//    differently-styled bullet header) collapses to ONE style for the whole
//    shape (its first run's size/color/font) — real per-run rich text needs
//    multiple positioned Text nodes with measured heights, a follow-up.
//  - Autoshapes: rect/roundRect/ellipse/triangle/diamond/pentagon/hexagon
//    map onto Konva's own Rect/Ellipse/RegularPolygon; star5 → Konva Star;
//    rightArrow → a single-segment Konva Arrow (a pointer line, not the
//    filled block-arrow prstGeom shape @injoffice/shapes' geometry.ts
//    already has real path math for — porting that math to a Konva
//    <Shape sceneFunc> is a natural follow-up, not done here); line → Konva
//    Line.
//  - Fixed-width sizing (a `width` prop), no ResizeObserver-driven
//    responsive auto-fit like DeckView's cqw layout — a follow-up.
//
// S8 editing (basic slice — reposition + resize only, per shape, on the
// CURRENTLY VISIBLE slide):
//  - Every shape renders inside a <Group> whose x/y/width/height are always
//    TOP-LEFT px, regardless of the inner primitive's own Konva anchor
//    convention (Ellipse/Star/RegularPolygon are center-anchored; Rect/
//    Line/Text are top-left) — the inner primitive renders at LOCAL
//    coordinates relative to that Group's origin, so drag/resize logic
//    never special-cases shape kind. The px→EMU math for both is pure and
//    tested (canvasGeometry.dragResultEmu/resizeResultEmu); this file is
//    just the Konva event wiring around it.
//  - editable=false (the default): Groups aren't draggable, no Transformer
//    — identical rendering to before S8.
//  - editable=true: click a shape to select it (a Konva Transformer
//    attaches, giving 8 resize handles); drag to reposition; drag a handle
//    to resize (MIN_SHAPE_PX floor). Both commit ONLY on release
//    (dragend/transformend), not per-frame — Konva handles the live visual
//    drag/resize itself.
//  - NOT done here (explicitly out of scope this round): rotate, crop,
//    z-order, format-painter, multi-select, keyboard nudge/delete.
//  - Persistence (added round 4): every commit fires onShapeOverride(
//    slideIndex, key, patch) — `key` is the shape's STABLE ROLE within its
//    slide kind (compile.ts assigns one to every shape it produces, e.g.
//    'title'/'subtitle'/'bullets'; see types.ts's SlideSpec.shapeOverrides
//    and compile.ts's own doc comment for the full round trip). A host
//    folds that into SlideSpec.shapeOverrides[key] and saves the DeckSpec
//    through its normal path — the NEXT compile (this component's own
//    resync, or any other consumer of compileDeckToWire, including the
//    POST /v1/pptx/build export) already reflects it, no extra plumbing.
//    onDeckChange (the pre-existing callback) is still separate and still
//    session-local: it's for a host that wants the edited layout
//    IMMEDIATELY (e.g. an "export what I see" action) without waiting on
//    a save round trip.
//
// S9 partial start (added round 4): DeckCanvasView now takes an OPTIONAL
// `deck` prop — a pre-parsed/pre-compiled WireDeck to render directly,
// bypassing DeckSpec/compileDeckToWire entirely. This is how a real
// uploaded .pptx gets onto this same canvas: a host parses it server-side
// via the gateway's POST /v1/pptx/parse (pptxpatch.ParsePPTX, independently
// validated against a real python-pptx fixture — see the gateway repo) and
// hands the response straight to `deck`. `spec` becomes optional to make
// room for this — exactly one of spec/deck is the intended usage (see
// DeckCanvasViewProps' own doc comments for the exact precedence and what
// stops working in deck-only mode: onShapeOverride, since there's no
// SlideSpec to persist an override into). Deliberately NOT built this
// round: any client that actually calls POST /v1/pptx/parse and uses this
// prop (no upload UI exists yet), and any write-back path (parse → edit →
// re-export as a real .pptx) — this is the render-a-real-deck half of S9
// only, an honest partial start, not a finished round-trip editor.

// shapeText renders a shape's own paragraph text — pulled out of
// KonvaShapeInner (round 6 fix) because OUTLINE and TEXT are independent in
// real OOXML: any shape can carry BOTH an outline (fill/stroke/prstGeom) AND
// its own embedded text (a:txBody) at once — a callout box with a label
// typed inside it, say — 'textBox' (prstGeom="rect" + txBox="1") is just the
// one special case where there's conventionally no VISIBLE outline, not the
// only shape kind that can ever have text. The previous version gated this
// entire block behind `shape.kind === 'textBox'`, so any other kind's
// paragraphs (an autoshape's own text, or an unresolved/empty kind — e.g. a
// real .pptx's title/body placeholder, which never carries an explicit
// prstGeom of its own) were silently never painted at all, even though
// compileDeckToWire itself was fine — this only ever showed up parsing a
// REAL uploaded .pptx (compile.ts's own textBox()/rect() helpers never
// produce a non-textBox shape with paragraphs, so no DeckSpec-driven deck
// ever exercised this path). Confirmed live on staging: an autoshape with
// its own text rendered its fill/stroke correctly but the text inside it
// was completely absent; a title/body placeholder (kind "", no fill/stroke
// of its own either) rendered as literally nothing, for the same reason.
// Exported (not just internal to KonvaShapeInner) so its behavior is
// directly unit-testable without a full react-konva render: the load-
// bearing property this fix depends on is that shapeText's output depends
// ONLY on shape.paragraphs, never on shape.kind — see
// DeckCanvasView.shapeText.test.ts.
export function shapeText(shape: WireShape, w: number, h: number, scale: number) {
  const paras = shape.paragraphs ?? []
  if (paras.length === 0) return null
  const firstRun = paras[0]?.runs[0]
  const flatLines = paras.map((p) => paragraphLine(p.runs.map((run) => run.text).join(''), p.bullet, p.level))
  const align = paras[0]?.align === 'ctr' ? 'center' : paras[0]?.align === 'r' ? 'right' : 'left'
  const fontStyle = [firstRun?.bold ? 'bold' : '', firstRun?.italic ? 'italic' : ''].filter(Boolean).join(' ') || 'normal'
  const baseFontSize = ptToPx(firstRun?.sizePt || 14, scale)
  const fontSize = fitFontSizePx(flatLines.join('\n'), w, baseFontSize)
  // Pre-wrap ourselves at fontSize rather than asking Konva/the browser
  // canvas to decide line breaks (wrap="word" below is intentionally
  // OFF) — see canvasGeometry.ts's "Text overflow defense" section:
  // Konva's own width-based wrap depends on the browser canvas's own
  // text measurement, and a live-staging report showed that not being
  // trustworthy (title text rendering as one long unwrapped line
  // reaching the canvas edge) in a way this package's headless test
  // environment could not reproduce or fully diagnose. wrapLineToWidth
  // is deterministic and testable, independent of actual font metrics.
  const wrapped = flatLines.flatMap((line) => wrapLineToWidth(line, w, fontSize))
  return (
    <Text
      x={0}
      y={0}
      width={w}
      height={h}
      text={wrapped.join('\n')}
      fontSize={fontSize}
      fontFamily={canvasFontFamily(firstRun?.font)}
      fontStyle={fontStyle}
      fill={firstRun?.color || '#000000'}
      align={align}
      wrap="none"
      listening={false}
    />
  )
}

function KonvaShapeInner({ shape, w, h, scale }: { shape: WireShape; w: number; h: number; scale: number }) {
  const strokeWidth = shape.strokeWidthPt ? ptToPx(shape.strokeWidthPt, scale) : undefined

  // A plain textBox conventionally has no visible outline of its own — text
  // only, same as before this fix.
  if (shape.kind === 'textBox') {
    return shapeText(shape, w, h, scale)
  }

  // Every other kind: its own outline (when the kind maps to one), THEN its
  // own text on top (when it has any) — independent, not mutually
  // exclusive. Text renders last so it's never painted UNDER a fill.
  let outline: ReactNode

  if (shape.kind === 'line') {
    outline = <Line points={[0, 0, w, h]} stroke={shape.stroke || '#000000'} strokeWidth={strokeWidth ?? 1} listening={false} />
  } else if (shape.kind === 'ellipse') {
    outline = (
      <Ellipse
        x={w / 2}
        y={h / 2}
        radiusX={w / 2}
        radiusY={h / 2}
        fill={shape.fill || undefined}
        stroke={shape.stroke || undefined}
        strokeWidth={strokeWidth}
        listening={false}
      />
    )
  } else if (shape.kind === 'star5') {
    const outerR = Math.min(w, h) / 2
    outline = (
      <Star
        x={w / 2}
        y={h / 2}
        numPoints={5}
        innerRadius={outerR * 0.42}
        outerRadius={outerR}
        fill={shape.fill || undefined}
        stroke={shape.stroke || undefined}
        strokeWidth={strokeWidth}
        listening={false}
      />
    )
  } else if (shape.kind === 'rightArrow') {
    const midY = h / 2
    outline = (
      <Arrow
        points={[0, midY, w, midY]}
        pointerLength={Math.min(w * 0.3, 20)}
        pointerWidth={Math.min(h * 0.6, 20)}
        fill={shape.fill || shape.stroke || '#000000'}
        stroke={shape.stroke || shape.fill || '#000000'}
        strokeWidth={strokeWidth ?? Math.max(2, h * 0.15)}
        listening={false}
      />
    )
  } else {
    const sides = regularPolygonSides(shape.kind)
    if (sides) {
      const radius = Math.min(w, h) / 2
      outline = (
        <RegularPolygon
          x={w / 2}
          y={h / 2}
          sides={sides}
          radius={radius}
          fill={shape.fill || undefined}
          stroke={shape.stroke || undefined}
          strokeWidth={strokeWidth}
          listening={false}
        />
      )
    } else {
      // rect / roundRect (and the honest fallback for anything else,
      // including an unresolved/empty kind — e.g. a real .pptx's
      // placeholder, which never carries an explicit prstGeom of its own).
      // fill/stroke both undefined (the common placeholder case) paints
      // nothing here, same as before — only the text below is new.
      outline = (
        <Rect
          x={0}
          y={0}
          width={w}
          height={h}
          cornerRadius={shape.kind === 'roundRect' ? roundRectCornerRadiusPx(w, h) : 0}
          fill={shape.fill || undefined}
          stroke={shape.stroke || undefined}
          strokeWidth={strokeWidth}
          listening={false}
        />
      )
    }
  }

  return (
    <>
      {outline}
      {shapeText(shape, w, h, scale)}
    </>
  )
}

/** Same shape as ShapePositionOverride (types.ts) — kept as its own type
 *  here since this file's internal edit-patch plumbing predates
 *  SlideSpec.shapeOverrides and the two serve slightly different callers
 *  (this one is always a delta from ONE commit; a stored override can
 *  accumulate fields across several). Exported so a host wiring
 *  onShapeOverride can name the patch type without reaching into
 *  DeckCanvasView's internals. */
export interface ShapeEditPatch {
  x?: number
  y?: number
  cx?: number
  cy?: number
}

function KonvaShape({
  shape,
  index,
  scale,
  editable,
  selected,
  onSelect,
  onCommit,
  registerNode,
}: {
  shape: WireShape
  index: number
  scale: number
  editable: boolean
  selected: boolean
  onSelect: (index: number) => void
  onCommit: (index: number, patch: ShapeEditPatch) => void
  registerNode: (index: number, node: Konva.Group | null) => void
}) {
  const r = shapeRectPx(shape, scale)

  return (
    <Group
      ref={(node) => registerNode(index, node)}
      x={r.x}
      y={r.y}
      width={r.w}
      height={r.h}
      // Hard clip to the shape's own box, in its local (pre-scale)
      // coordinate space — scales along with the Group during a live
      // resize, so it always matches whatever KonvaShapeInner is drawing.
      // The structural guarantee half of the text-overflow fix (see
      // canvasGeometry.ts): NOTHING can ever paint outside a shape's own
      // bounds, regardless of any font/measurement mismatch.
      clipX={0}
      clipY={0}
      clipWidth={r.w}
      clipHeight={r.h}
      draggable={editable}
      onClick={(e) => {
        if (!editable) return
        e.cancelBubble = true
        onSelect(index)
      }}
      onTap={(e) => {
        if (!editable) return
        e.cancelBubble = true
        onSelect(index)
      }}
      onDragEnd={(e) => {
        const node = e.target
        onCommit(index, dragResultEmu(node.x(), node.y(), scale))
      }}
      onTransformEnd={(e) => {
        const node = e.target as Konva.Group
        const newWidthPx = node.width() * node.scaleX()
        const newHeightPx = node.height() * node.scaleY()
        // Konva's Transformer works by scaling, not by resizing width/
        // height directly — commit the scaled size as the Group's real
        // width/height, then reset scale to 1 so it doesn't compound on the
        // NEXT resize (the standard Konva pattern for a resizable node).
        node.scaleX(1)
        node.scaleY(1)
        const size = resizeResultEmu(newWidthPx, newHeightPx, scale)
        node.width(Math.max(MIN_SHAPE_PX, newWidthPx))
        node.height(Math.max(MIN_SHAPE_PX, newHeightPx))
        onCommit(index, { ...dragResultEmu(node.x(), node.y(), scale), ...size })
      }}
    >
      <KonvaShapeInner shape={shape} w={r.w} h={r.h} scale={scale} />
      {editable && (
        // Every visible primitive above is rendered with listening={false}
        // (it's decorative-only, same as the read-only S7.3 renderer) — so
        // WITHOUT this, the wrapping Group would have no hit region at all
        // to click/drag (Konva's default hit detection comes from listening
        // children; a Group draws nothing of its own) and could never be
        // selected in the first place. This invisible full-bounds rect is
        // that hit region — present whenever editable, not just once
        // already selected (a shape has to be clickable to BECOME
        // selected).
        <Rect x={0} y={0} width={r.w} height={r.h} fill="transparent" />
      )}
      {editable && selected && (
        // A visible selection outline — the Transformer's own handles show
        // WHERE to grab, but nothing else marks WHICH shape is currently
        // selected once you're not actively dragging a handle.
        <Rect x={0} y={0} width={r.w} height={r.h} stroke="#2f6fed" strokeWidth={1.5} dash={[4, 3]} listening={false} />
      )}
    </Group>
  )
}

function KonvaSlide({
  slide,
  cx,
  cy,
  widthPx,
  editable,
  onCommit,
}: {
  slide: WireSlide
  cx: number
  cy: number
  widthPx: number
  editable: boolean
  onCommit: (index: number, patch: ShapeEditPatch) => void
}) {
  const vp = useMemo(() => makeViewport(cx, cy, widthPx), [cx, cy, widthPx])
  const [selected, setSelected] = useState<number | null>(null)
  const nodesRef = useRef<Map<number, Konva.Group>>(new Map())
  const transformerRef = useRef<Konva.Transformer>(null)

  // Reset selection whenever editing turns off or the slide's own shape
  // list changes shape (a different slide, or content edited elsewhere) —
  // a stale selected index pointing at nothing (or the wrong shape) would
  // otherwise leave the Transformer attached to whatever now sits at that
  // index.
  useEffect(() => {
    setSelected(null)
  }, [editable, slide])

  useEffect(() => {
    const tr = transformerRef.current
    if (!tr) return
    const node = selected !== null ? nodesRef.current.get(selected) : null
    tr.nodes(node ? [node] : [])
    tr.getLayer()?.batchDraw()
  }, [selected, slide])

  const registerNode = (index: number, node: Konva.Group | null) => {
    if (node) nodesRef.current.set(index, node)
    else nodesRef.current.delete(index)
  }

  return (
    <Stage width={vp.widthPx} height={vp.heightPx}>
      <Layer>
        <Rect
          x={0}
          y={0}
          width={vp.widthPx}
          height={vp.heightPx}
          fill={slide.background || '#ffffff'}
          listening={editable}
          onClick={() => editable && setSelected(null)}
          onTap={() => editable && setSelected(null)}
        />
        {slide.shapes.map((shape, i) => (
          <KonvaShape
            key={i}
            shape={shape}
            index={i}
            scale={vp.scale}
            editable={editable}
            selected={selected === i}
            onSelect={setSelected}
            onCommit={onCommit}
            registerNode={registerNode}
          />
        ))}
        {editable && (
          <Transformer
            ref={transformerRef}
            rotateEnabled={false}
            flipEnabled={false}
            boundBoxFunc={(oldBox, newBox) => (newBox.width < MIN_SHAPE_PX || newBox.height < MIN_SHAPE_PX ? oldBox : newBox)}
          />
        )}
      </Layer>
    </Stage>
  )
}

export interface DeckCanvasViewProps {
  /** Compile a DeckSpec through compileDeckToWire — the normal, template-
   *  driven path. Provide exactly one of spec/deck; if both are given,
   *  `deck` wins (it's the more direct source, and skips a needless
   *  compile). */
  spec?: DeckSpec
  /** A pre-compiled/parsed wire Deck to render directly, bypassing
   *  DeckSpec/compileDeckToWire entirely (S9 partial start) — how a real
   *  uploaded .pptx gets onto this SAME canvas: parse it server-side via
   *  the gateway's POST /v1/pptx/parse (pptxpatch.ParsePPTX) and hand the
   *  response's wire Deck straight to this prop. Its content generally
   *  won't map onto any SlideKind template at all, which is exactly why
   *  this bypasses DeckSpec rather than trying to force a reconciliation.
   *  In this mode: onShapeOverride never fires (there is no SlideSpec to
   *  persist an override into — no write-back path exists yet for a
   *  parsed real-world deck); onDeckChange (session-local) behaves
   *  identically to the spec-driven path. */
  deck?: WireDeck
  /** Rendered stage width in px; height follows the deck's slide aspect ratio. */
  width?: number
  /** Controlled slide index (mirrors DeckView's at/onAtChange). */
  at?: number
  onAtChange?: (index: number) => void
  /** Enable drag-to-reposition + resize on the CURRENTLY VISIBLE slide's
   *  shapes (S8, basic slice — see this file's top doc comment for exactly
   *  what is and isn't covered). Default false: identical to the read-only
   *  S7.3 renderer. */
  editable?: boolean
  /** Fires with the full updated wire Deck after every committed drag/
   *  resize (on release, not per-frame). Always session/mount-local (this
   *  component's own working copy, lost on unmount or a spec change) — a
   *  host MAY hold onto it for an immediate "export what I see, including
   *  my layout tweaks" flow (compile it straight into a POST /v1/pptx/build
   *  call) without waiting on a save round trip. For DURABLE persistence,
   *  see onShapeOverride instead — that's what survives a reload. */
  onDeckChange?: (deck: WireDeck) => void
  /** Fires once per committed drag/resize with (slideIndex, shape key,
   *  patch) — everything a host needs to fold into
   *  SlideSpec.shapeOverrides[key] and persist through its normal DeckSpec
   *  save path (see types.ts's ShapePositionOverride). Only fires for a
   *  shape that HAS a key (every shape compileDeckToWire itself produces
   *  does; a WireDeck a host built by hand without keys just won't surface
   *  overrides here — no silent guessing at identity). slideIndex is the
   *  index into the CURRENT `spec.slides` (matches DeckSpec's own
   *  indexing), not `at`/`atProp`, though today they're always the same
   *  since editing only ever touches the visible slide. */
  onShapeOverride?: (slideIndex: number, key: string, patch: ShapeEditPatch) => void
}

/** The compiled wire Deck a DeckCanvasView is currently showing — exported
 *  so a host that also needs the wire JSON (e.g. to POST it to
 *  /v1/pptx/build) can compile once and hand it to both. Note this is the
 *  FRESH compile from spec, not DeckCanvasView's own (possibly
 *  drag/resize-edited) working copy — see onDeckChange to observe edits. */
export function useCompiledDeck(spec: DeckSpec): WireDeck {
  return useMemo(() => compileDeckToWire(spec), [spec])
}

const EMPTY_WIRE_DECK: WireDeck = { slides: [] }

/**
 * @deprecated Legacy Konva preview/editor only. It is not a production
 * Office-layout authority. Use @injoffice/pptx-authored plus
 * @injoffice/pptx-render for native authored-deck rendering.
 */
export function DeckCanvasView({ spec, deck, width = 960, at: atProp, onAtChange, editable = false, onDeckChange, onShapeOverride }: DeckCanvasViewProps) {
  // deck (a pre-parsed/pre-compiled WireDeck, S9) always wins when given —
  // compileDeckToWire only runs at all when there's a spec AND no deck, so
  // a deck-driven host never pays for a needless compile. Neither given is
  // the degenerate case (renders nothing, same as an empty deck).
  const compiledFromSpec = useMemo(() => (spec ? compileDeckToWire(spec) : EMPTY_WIRE_DECK), [spec])
  const compiled = deck ?? compiledFromSpec
  // The working copy starts as the fresh compile/deck and re-syncs to it
  // whenever `compiled` itself changes (spec content edited elsewhere, or
  // the host hands in a different deck). Once a host wires onShapeOverride
  // through to a SAVE (persisting into SlideSpec.shapeOverrides), THAT
  // resync will already reflect the saved edits on the next compile — but
  // until the host actually saves, this working copy is still the only
  // place an in-progress drag/resize lives. Editing (when enabled) mutates
  // this copy directly; it is what's actually rendered, in both modes, so
  // read-only rendering is unaffected (nothing ever diverges from
  // `compiled` when editable is false, since nothing calls setWorking).
  const [working, setWorking] = useState<WireDeck>(compiled)
  useEffect(() => {
    setWorking(compiled)
  }, [compiled])

  const [atState, setAtState] = useState(0)
  const at = atProp ?? atState
  const total = working.slides.length
  const clamped = Math.min(Math.max(at, 0), Math.max(total - 1, 0))
  const slide = working.slides[clamped]

  const setAt = (next: number) => {
    if (onAtChange) onAtChange(next)
    if (atProp === undefined) setAtState(next)
  }

  const onCommit = (shapeIndex: number, patch: ShapeEditPatch) => {
    const key = slide.shapes[shapeIndex]?.key
    // onShapeOverride is a SlideSpec.shapeOverrides concept — only
    // meaningful (and only fired) when this view is actually spec-driven.
    // A deck-driven (S9, parsed-upload) render has no SlideSpec to persist
    // an override into.
    if (key && spec) onShapeOverride?.(clamped, key, patch)
    setWorking((prev) => {
      const next: WireDeck = {
        ...prev,
        slides: prev.slides.map((s, si) =>
          si !== clamped ? s : { ...s, shapes: s.shapes.map((sh, shi) => (shi !== shapeIndex ? sh : { ...sh, ...patch })) },
        ),
      }
      onDeckChange?.(next)
      return next
    })
  }

  if (!slide) return null
  const cx = working.cx || 12192000
  const cy = working.cy || 6858000

  return (
    <div className="ios-deck-canvas" style={{ display: 'flex', flexDirection: 'column', gap: 10 }}>
      <KonvaSlide slide={slide} cx={cx} cy={cy} widthPx={width} editable={editable} onCommit={onCommit} />
      {total > 1 && (
        <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
          <button type="button" className="ioc-add" disabled={clamped === 0} onClick={() => setAt(clamped - 1)}>
            ← Prev
          </button>
          <span>
            {clamped + 1} / {total}
          </span>
          <button type="button" className="ioc-add" disabled={clamped === total - 1} onClick={() => setAt(clamped + 1)}>
            Next →
          </button>
        </div>
      )}
    </div>
  )
}
