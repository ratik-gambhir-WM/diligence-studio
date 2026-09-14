import { describe, expect, it } from 'vitest'

import {
  buildSlideEmbeddingDocument,
  buildSlideEmbeddingFingerprint,
} from '../src/lib/retrieval/SlideEmbeddingDocument'
import {
  cosineSimilarity,
  deserializeEmbedding,
  serializeEmbedding,
} from '../src/lib/retrieval/SlideVector'
import { normalizeSlideRetrievalMetadata } from '../src/lib/retrieval/SlideRetrievalMetadata'
import { VALID_METADATA } from './slideTestFixtures'

describe('slide embedding contracts', () => {
  it('builds deterministic labeled documents with only positive facets', () => {
    const metadata = normalizeSlideRetrievalMetadata(VALID_METADATA)
    const document = buildSlideEmbeddingDocument('Product Architecture', 'diagram', metadata, 10_000)
    expect(document).toContain('title: Product Architecture')
    expect(document).toContain('technologies: Salesforce | Snowflake')
    expect(document).toContain('positive facets: process flow')
    expect(document).not.toContain('has_table')
    expect(buildSlideEmbeddingFingerprint(document, 'embedding-model', 3))
      .toBe(buildSlideEmbeddingFingerprint(document, 'embedding-model', 3))
    expect(() => buildSlideEmbeddingDocument('Title', 'diagram', metadata, 10)).toThrow()
  })

  it('round trips little-endian Float32 vectors and validates cosine similarity', () => {
    const bytes = serializeEmbedding([1, 0.5, -0.25])
    expect(bytes.readFloatLE(0)).toBe(1)
    expect(deserializeEmbedding(bytes, 3)).toEqual([1, 0.5, -0.25])
    expect(cosineSimilarity([1, 0], [1, 0])).toBe(1)
    expect(() => serializeEmbedding([0, 0])).toThrow('non-zero norm')
    expect(() => deserializeEmbedding(bytes, 2)).toThrow('dimensions')
  })
})
