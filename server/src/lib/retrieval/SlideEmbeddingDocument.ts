import { createHash } from 'node:crypto'

import type { TemplateKind } from '../../repositories/TemplateRepository'
import type { SlideRetrievalMetadata } from './SlideRetrievalMetadata'

export const SLIDE_EMBEDDING_DOCUMENT_VERSION = 2

export type SlideEmbeddingDocuments = {
  capability: string
  subject: string
}

/** Project subject meaning and reusable template capability into independent documents. */
export function buildSlideEmbeddingDocuments(
  title: string,
  kind: TemplateKind,
  metadata: SlideRetrievalMetadata,
  maximumBytes: number,
): SlideEmbeddingDocuments {
  const domainTopics = metadata.subject.domains.map((domain) => (
    `${domain.id} (${domain.relevance}): ${domain.topics.join(' | ') || 'general'}`
  ))
  const subject = buildDocument([
    ['title', title],
    ['summary', metadata.subject.summary],
    ['domain topics', domainTopics],
    ['other topics', metadata.subject.other_topics],
    ['claims', metadata.subject.claims],
    ['technologies', metadata.subject.technologies],
    ['entities', metadata.subject.entities],
    ['synonyms', metadata.subject.synonyms],
    ['retrieval keywords', metadata.retrieval_keywords],
  ], maximumBytes)
  const slots = metadata.template_fit.content_slots.map((slot) => (
    `${slot.role}: ${slot.capacity}`
  ))
  const capability = buildDocument([
    ['kind', kind],
    ['communication intents', metadata.communication.intents],
    ['information types', metadata.communication.information_types],
    ['audience', metadata.communication.audience],
    ['archetype', metadata.template_fit.archetype],
    ['content slots', slots],
    ['layout', metadata.visual.layout_type],
    ['visual elements', metadata.visual.visual_elements],
    ['content density', metadata.visual.content_density],
    ['structural features', metadata.visual.structural_features],
  ], maximumBytes)
  return { capability, subject }
}

/** Fingerprint the exact document role, contents, model, dimensions, and builder version. */
export function buildSlideEmbeddingFingerprint(
  role: keyof SlideEmbeddingDocuments,
  document: string,
  model: string,
  dimensions: number | undefined,
) {
  return createHash('sha256').update([
    String(SLIDE_EMBEDDING_DOCUMENT_VERSION),
    role,
    model,
    dimensions === undefined ? '' : String(dimensions),
    document,
  ].join('\n')).digest('hex')
}

function buildDocument(
  fields: ReadonlyArray<readonly [string, string | readonly string[]]>,
  maximumBytes: number,
) {
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
