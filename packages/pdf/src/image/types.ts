export type ImageLayer = 'belowText' | 'aboveText'

export type ImageEditSpec =
  | {
      kind: 'insertImage'
      page: number
      image: string
      rect: [number, number, number, number]
      layer: ImageLayer
      rotate?: number
    }
  | {
      kind: 'transformImage'
      page: number
      oldRect: [number, number, number, number]
      rect: [number, number, number, number]
      layer?: ImageLayer
      quarterTurns?: number
    }
  | {
      kind: 'replaceImage'
      page: number
      oldRect: [number, number, number, number]
      rect: [number, number, number, number]
      image: string
      layer?: ImageLayer
      quarterTurns?: number
    }
  | {
      kind: 'deleteImage'
      page: number
      oldRect: [number, number, number, number]
    }

export interface PageImageRef {
  page: number
  rect: [number, number, number, number]
  aboveText: boolean
}

export interface ImageEditFailure {
  editIndex: number
  page: number
  reason: string
}

export interface ImageEditsResult {
  bytes: Uint8Array
  skipped: ImageEditFailure[]
}
