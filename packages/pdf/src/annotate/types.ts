export type MarkupType = 'highlight' | 'underline' | 'strikeout'

export interface NoteReplyTarget {
  objNum: number
  rect: [number, number, number, number]
  contents: string
}

interface DrawBase {
  page: number
  color: [number, number, number]
}

interface StrokedDraw extends DrawBase {
  width: number
}

export type DrawingSpec =
  | (StrokedDraw & { kind: 'ink'; paths: number[][] })
  | (StrokedDraw & { kind: 'rect'; rect: [number, number, number, number] })
  | (StrokedDraw & { kind: 'ellipse'; rect: [number, number, number, number] })
  | (StrokedDraw & { kind: 'line'; from: [number, number]; to: [number, number] })
  | (StrokedDraw & { kind: 'arrow'; from: [number, number]; to: [number, number] })
  | (DrawBase & {
      kind: 'note'
      at: [number, number]
      contents: string
      author?: string
      createdMs?: number
      localId?: string
      replyToSaved?: NoteReplyTarget
      replyToLocalId?: string
    })

export interface NoteEditSpec {
  page: number
  objNum: number
  rect: [number, number, number, number]
  oldContents: string
  contents: string
}

export interface FormValueSpec {
  name: string
  kind: 'text' | 'checkbox' | 'radio' | 'choice'
  value?: string
  checked?: boolean
}

export interface StampSpec {
  page: number
  image: string
  rect: [number, number, number, number]
  opacity?: number
}

export interface SignatureStampSpec {
  page: number
  image: string
  rect: [number, number, number, number]
  formFieldName?: string
}

export interface MarkupSpec {
  page: number
  type: MarkupType
  color: [number, number, number]
  quads: number[][]
}

export interface AnnotDeleteSpec {
  page: number
  objNum: number
  subtype: MarkupType | 'note'
  rect: [number, number, number, number]
  contents?: string
}
