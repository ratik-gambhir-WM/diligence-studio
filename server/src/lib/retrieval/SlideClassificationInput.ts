import { createHash } from 'node:crypto'

import type {
  PowerPointCanvasElement,
  PowerPointCanvasJson,
} from '../import/PowerpointImportTypes'
import type { TemplateKind } from '../../repositories/TemplateRepository'
import {
  SLIDE_CLASSIFICATION_PROMPT_VERSION,
  SLIDE_CLASSIFICATION_SCHEMA_VERSION,
} from './SlideRetrievalMetadata'

export const SLIDE_CLASSIFICATION_INPUT_BUILDER_VERSION = 1

export type SlideClassificationInput = {
  digest: string
  preview?: {
    bytes: Buffer
    contentType: 'image/png'
  }
}

export type SlideClassificationInputLimits = {
  maxDigestBytes: number
  maxElements: number
  maxPreviewBytes: number
  maxPreviewDimension: number
  maxTextChars: number
}

export const DEFAULT_SLIDE_CLASSIFICATION_INPUT_LIMITS: SlideClassificationInputLimits = {
  maxDigestBytes: 32_000,
  maxElements: 250,
  maxPreviewBytes: 5 * 1024 * 1024,
  maxPreviewDimension: 4_096,
  maxTextChars: 12_000,
}

export type BuildSlideClassificationInputOptions = {
  kind: TemplateKind
  limits?: Partial<SlideClassificationInputLimits>
  preview?: { bytes: Buffer; contentType: 'image/png'; height: number; width: number }
  templateJson: PowerPointCanvasJson
  title?: string
}

/** Build the bounded, redacted text/image input sent to the slide classifier. */
export function buildSlideClassificationInput(
  options: BuildSlideClassificationInputOptions,
): SlideClassificationInput {
  const limits = { ...DEFAULT_SLIDE_CLASSIFICATION_INPUT_LIMITS, ...options.limits }
  const slide = options.templateJson.presentation.slides[0]
  if (!slide || options.templateJson.presentation.slides.length !== 1) {
    throw new Error('Classification input requires exactly one slide.')
  }

  const counts = new Map<PowerPointCanvasElement['type'], number>()
  for (const element of slide.elements) counts.set(element.type, (counts.get(element.type) ?? 0) + 1)
  let textCharacters = 0
  const elementLines: string[] = []
  const visibleElements = slide.elements.slice(0, limits.maxElements)
  for (const [index, element] of visibleElements.entries()) {
    elementLines.push(describeElement(element, index + 1, slide.width, slide.height, (text) => {
      const remaining = Math.max(0, limits.maxTextChars - textCharacters)
      const bounded = text.slice(0, remaining)
      textCharacters += bounded.length
      return bounded
    }))
  }

  const verticalBuckets = new Set<number>()
  const horizontalBuckets = new Set<number>()
  for (const element of visibleElements) {
    if (element.type === 'line') continue
    verticalBuckets.add(Math.round((element.x / slide.width) * 20))
    horizontalBuckets.add(Math.round((element.y / slide.height) * 20))
  }
  const connectedLines = visibleElements.filter((element) => (
    element.type === 'line' && (element.beginArrow !== undefined || element.endArrow !== undefined)
  )).length

  const lines = [
    `title: ${normalizeText(options.title ?? options.templateJson.presentation.title)}`,
    `kind: ${options.kind}`,
    `slide: ${normalizeText(slide.name)}`,
    `dimensions: ${round(slide.width)} x ${round(slide.height)}`,
    `background: ${normalizeText(slide.backgroundColor)}`,
    `element counts: ${[...counts.entries()].sort(([left], [right]) => left.localeCompare(right)).map(([type, count]) => `${type}=${count}`).join(' | ') || 'none'}`,
    `layout hints: x-buckets=${verticalBuckets.size} | y-buckets=${horizontalBuckets.size} | connected-lines=${connectedLines}`,
    ...elementLines.filter(Boolean),
  ]
  if (slide.elements.length > visibleElements.length) {
    lines.push(`truncated elements: ${slide.elements.length - visibleElements.length}`)
  }
  const digest = truncateUtf8(lines.join('\n'), limits.maxDigestBytes)
  if (!digest.trim()) throw new Error('Classification digest cannot be empty.')

  const preview = options.preview
  const includePreview = preview
    && preview.contentType === 'image/png'
    && preview.bytes.length > 0
    && preview.bytes.length <= limits.maxPreviewBytes
    && preview.width > 0
    && preview.height > 0
    && preview.width <= limits.maxPreviewDimension
    && preview.height <= limits.maxPreviewDimension
  return includePreview
    ? { digest, preview: { bytes: Buffer.from(preview.bytes), contentType: 'image/png' } }
    : { digest }
}

/** Fingerprint every input and version that can change classification output. */
export function buildSlideClassificationInputFingerprint(
  templateJson: PowerPointCanvasJson,
  previewChecksum: string | null,
): string {
  const canonical = stableStringify(templateJson)
  return createHash('sha256').update([
    String(SLIDE_CLASSIFICATION_SCHEMA_VERSION),
    SLIDE_CLASSIFICATION_PROMPT_VERSION,
    String(SLIDE_CLASSIFICATION_INPUT_BUILDER_VERSION),
    previewChecksum ?? '',
    canonical,
  ].join('\n')).digest('hex')
}

/** Return a stable SHA-256 checksum for preview provenance without retaining its bytes. */
export function checksumBytes(bytes: Uint8Array) {
  return createHash('sha256').update(bytes).digest('hex')
}

function describeElement(
  element: PowerPointCanvasElement,
  index: number,
  slideWidth: number,
  slideHeight: number,
  takeText: (text: string) => string,
) {
  if (element.type === 'line') {
    return `element ${index}: line; route=${element.lineType}; arrows=${element.beginArrow ?? 'none'}->${element.endArrow ?? 'none'}; box=${normalizedLineBox(element, slideWidth, slideHeight)}`
  }
  const box = normalizedBox(element.x, element.y, element.w, element.h, slideWidth, slideHeight)
  if (element.type === 'image') {
    return `element ${index}: image; alt=${normalizeText(element.altText ?? '') || 'none'}; fit=${element.fit}; crop=${stableStringify(element.crop ?? {})}; box=${box}`
  }
  const runs = element.runs?.map((run) => run.text).join('') ?? ''
  const visibleText = normalizeText(element.text ?? runs)
  const text = takeText(visibleText)
  const shape = element.type === 'shape' ? `; shape=${normalizeText(element.shape)}` : ''
  return `element ${index}: ${element.type}${shape}; text=${text || 'none'}; box=${box}`
}

function normalizedLineBox(
  line: Extract<PowerPointCanvasElement, { type: 'line' }>,
  slideWidth: number,
  slideHeight: number,
) {
  return [line.x1 / slideWidth, line.y1 / slideHeight, line.x2 / slideWidth, line.y2 / slideHeight]
    .map((value) => round(value)).join(',')
}

function normalizedBox(x: number, y: number, width: number, height: number, slideWidth: number, slideHeight: number) {
  return [x / slideWidth, y / slideHeight, width / slideWidth, height / slideHeight]
    .map((value) => round(value)).join(',')
}

function round(value: number) {
  return Number.isFinite(value) ? Number(value.toFixed(4)) : 0
}

function normalizeText(value: string) {
  return value.trim().replace(/\s+/gu, ' ')
}

function truncateUtf8(value: string, maximumBytes: number) {
  const bytes = Buffer.from(value, 'utf8')
  if (bytes.length <= maximumBytes) return value
  return bytes.subarray(0, maximumBytes).toString('utf8').replace(/\uFFFD$/u, '')
}

function stableStringify(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map(stableStringify).join(',')}]`
  if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>
    return `{${Object.keys(record).sort().map((key) => `${JSON.stringify(key)}:${stableStringify(record[key])}`).join(',')}}`
  }
  return JSON.stringify(value)
}
