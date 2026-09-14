import { createHash } from 'node:crypto'

import type { TemplateKind } from '../../repositories/TemplateRepository'
import type { SlideRetrievalMetadata } from './SlideRetrievalMetadata'

export const SLIDE_EMBEDDING_DOCUMENT_VERSION = 1

const BOOLEAN_LABELS = [
  ['has_timeline', 'timeline'],
  ['has_table', 'table'],
  ['has_chart', 'chart'],
  ['has_process_flow', 'process flow'],
  ['has_kpis', 'kpis'],
  ['has_recommendations', 'recommendations'],
] as const satisfies readonly (readonly [keyof SlideRetrievalMetadata, string])[]

/** Project normalized metadata into a deterministic, labeled embedding document. */
export function buildSlideEmbeddingDocument(
  title: string,
  kind: TemplateKind,
  metadata: SlideRetrievalMetadata,
  maximumBytes: number,
): string {
  const fields: Array<readonly [string, string | readonly string[]]> = [
    ['title', title],
    ['kind', kind],
    ['slide type', metadata.slide_type],
    ['purpose', metadata.slide_purpose],
    ['description', metadata.description],
    ['topics', metadata.topics],
    ['business domains', metadata.business_domains],
    ['technologies', metadata.technologies],
    ['entities', metadata.entities],
    ['use cases', metadata.use_cases],
    ['audience', metadata.audience],
    ['layout', metadata.layout_type],
    ['visual elements', metadata.visual_elements],
    ['content density', metadata.content_density],
    ['information types', metadata.information_types],
    ['structural features', metadata.structural_features],
    ['retrieval keywords', metadata.retrieval_keywords],
    ['positive facets', BOOLEAN_LABELS.filter(([field]) => metadata[field] === true).map(([, label]) => label)],
  ]
  const document = fields.flatMap(([label, value]) => {
    const normalized = typeof value === 'string'
      ? normalize(value)
      : deduplicate(value.map(normalize)).join(' | ')
    return normalized ? [`${label}: ${normalized}`] : []
  }).join('\n')
  if (!document) throw new Error('Embedding document cannot be empty.')
  if (Buffer.byteLength(document, 'utf8') > maximumBytes) {
    throw new Error('Embedding document exceeds the configured byte limit.')
  }
  return document
}

/** Fingerprint the exact document, model, dimensions, and builder version used by an embedding. */
export function buildSlideEmbeddingFingerprint(
  document: string,
  model: string,
  dimensions: number | undefined,
) {
  return createHash('sha256').update([
    String(SLIDE_EMBEDDING_DOCUMENT_VERSION),
    model,
    dimensions === undefined ? '' : String(dimensions),
    document,
  ].join('\n')).digest('hex')
}

function normalize(value: string) {
  return value.trim().replace(/\s+/gu, ' ')
}

function deduplicate(values: readonly string[]) {
  const seen = new Set<string>()
  return values.filter((value) => {
    const key = value.toLocaleLowerCase('en-US')
    if (!value || seen.has(key)) return false
    seen.add(key)
    return true
  })
}
