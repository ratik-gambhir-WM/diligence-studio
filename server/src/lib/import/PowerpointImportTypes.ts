import type { ExtractedRelationship, XmlNode } from '../shared/PowerpointTypes'

export type TransformMatrix = {
  a: number
  b: number
  c: number
  d: number
  e: number
  f: number
}

export type ExtractedNonVisualRecord = {
  description?: string
  hidden?: boolean
  id?: number
  isTextBox?: boolean
  name?: string
  placeholder?: {
    idx?: string
    type?: string
  }
}

export type ExtractedTransformRecord = {
  flipH?: boolean
  flipV?: boolean
  heightInches: number
  heightPx: number
  rotation: number
  widthInches: number
  widthPx: number
  xInches: number
  xPx: number
  yInches: number
  yPx: number
}

export type ExtractedElementRecord = {
  kind: 'connector' | 'graphicFrame' | 'shape'
  nonVisual: ExtractedNonVisualRecord
  path: string
  presetGeometry?: {
    preset: string
    xmlAst?: XmlNode
  }
  relationshipIds?: string[]
  shapeProperties?: XmlNode
  style?: XmlNode
  tag: string
  text?: ExtractedTextBodyRecord
  transform: ExtractedTransformRecord
  xmlAst: XmlNode
  zIndex: number
}

export type ExtractedTextRunRecord = {
  text?: string
  properties?: Record<string, string>
}

export type ExtractedTextBodyRecord = {
  bodyProperties?: Record<string, string>
  paragraphs?: Array<{
    endParagraphRunProperties?: Record<string, string>
    properties?: Record<string, string>
    runs?: ExtractedTextRunRecord[]
  }>
  plainText?: string
}

export type ExtractedSupportPartRecord = {
  base64?: string
  contentTypeHint: string
  path: string
  rawXml?: string
  relationshipType: string
  size: number
}

export type ExtractedSlideRecord = {
  backgroundColor?: string
  rawRelationshipsXml?: string
  rawXml: string
  relationshipIds: string[]
  relationships: ExtractedRelationship[]
  relationshipsPath: string
  shapeTree?: {
    elements?: ExtractedElementRecord[]
  }
  slideNumber: number
  slidePath: string
  slideSize: SlideSize & {
    heightInches: number
    widthInches: number
  }
  sourcePptx: string
  summary: {
    textElementCount: number
    totalDrawableElements: number
  }
  supportParts: Record<string, ExtractedSupportPartRecord>
}

export type SlideSize = {
  cx: number
  cy: number
  widthPx: number
  heightPx: number
}

export type ThemeTypography = {
  bodyFont: string
  headingFont: string
  colors: Record<string, string>
}

export type PlaceholderSourceIndex = {
  layout: XmlNode[]
  master: XmlNode[]
  masterTextStyles: {
    body?: XmlNode
    other?: XmlNode
    title?: XmlNode
  }
}

export type PowerPointCanvasTextRun = {
  bold?: boolean
  breakLine?: boolean
  color: string
  fontFace: string
  fontSize: number
  italic?: boolean
  text: string
  underline?: boolean
}

type PowerPointCanvasElementBase = {
  id: string
  opacity?: number
  rotate?: number
}

export type PowerPointCanvasShapeElement = PowerPointCanvasElementBase & {
  align?: 'center' | 'left' | 'right'
  bold?: boolean
  borderRadius?: number
  fill: string
  fillOpacity?: number
  flipH?: boolean
  flipV?: boolean
  fontFace?: string
  fontSize?: number
  h: number
  padding?: number
  runs?: PowerPointCanvasTextRun[]
  shape: string
  stroke: string
  strokeOpacity?: number
  strokeWidth: number
  text?: string
  textColor?: string
  type: 'shape'
  valign?: 'bottom' | 'middle' | 'top'
  w: number
  x: number
  y: number
}

export type PowerPointCanvasTextElement = Omit<PowerPointCanvasShapeElement, 'shape' | 'type'> & {
  italic?: boolean
  type: 'text'
}

export type PowerPointCanvasLineElement = PowerPointCanvasElementBase & {
  beginArrow?: 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
  dash?: 'dash' | 'dot' | 'solid'
  elbowDirection?: 'horizontal-first' | 'vertical-first'
  endArrow?: 'arrow' | 'diamond' | 'oval' | 'stealth' | 'triangle'
  lineType: 'elbow' | 'straight'
  stroke: string
  strokeOpacity?: number
  strokeWidth: number
  type: 'line'
  x1: number
  x2: number
  y1: number
  y2: number
}

export type PowerPointCanvasImageElement = PowerPointCanvasElementBase & {
  altText?: string
  borderRadius?: number
  crop?: {
    bottom?: number
    left?: number
    right?: number
    top?: number
  }
  fit: 'contain' | 'cover' | 'stretch'
  flipH?: boolean
  flipV?: boolean
  h: number
  src: string
  type: 'image'
  w: number
  x: number
  y: number
}

export type PowerPointCanvasElement =
  | PowerPointCanvasImageElement
  | PowerPointCanvasLineElement
  | PowerPointCanvasShapeElement
  | PowerPointCanvasTextElement

export type PowerPointCanvasJson = {
  presentation: {
    title: string
    preserveElementOrder: boolean
    showBranding: boolean
    slides: Array<{
      id: string
      name: string
      width: number
      height: number
      backgroundColor: string
      elements: PowerPointCanvasElement[]
    }>
  }
}

export type ImportPowerPointBytesOptions = {
  /** Logical source name used for the presentation title. */
  sourceName?: string
  /** One-based slide number. Omit to import the complete deck. */
  slide?: number
}

export type ImportedPowerPoint = {
  importedSlideCount: number
  jsonSpec: PowerPointCanvasJson
  sourceSlideCount: number
  warnings: string[]
}
