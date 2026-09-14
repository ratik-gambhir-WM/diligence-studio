import { describe, expect, it } from 'vitest'

import { normalizeSlideRetrievalMetadata } from '../src/lib/retrieval/SlideRetrievalMetadata'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import type { RetrievalClassificationRecord } from '../src/repositories/TemplateRepository'
import {
  SlideRetrievalError,
  SlideRetrievalService,
  reciprocalRankFusion,
  type SlideQueryInput,
} from '../src/services/SlideRetrievalService'
import { createTemplate, VALID_METADATA } from './slideTestFixtures'

const EMPTY_FILTERS: SlideQueryInput['filters'] = {
  business_domains: [],
  content_density: [],
  has_chart: null,
  has_kpis: null,
  has_process_flow: null,
  has_recommendations: null,
  has_table: null,
  has_timeline: null,
  kinds: [],
  slide_types: [],
  technologies: [],
}

describe('slide retrieval service', () => {
  it('supports app-scoped lexical, semantic, hybrid, filters, IDs, similarity, and facets', async () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repository.ensureApp('app-one')
    repository.ensureApp('app-two')
    addReady(repository, 'app-one', 'architecture', {
      ...VALID_METADATA,
      slide_type: 'architecture',
      technologies: ['Snowflake'],
    }, [1, 0])
    addReady(repository, 'app-one', 'timeline', {
      ...VALID_METADATA,
      slide_type: 'timeline',
      technologies: ['SAP'],
      has_process_flow: false,
      has_timeline: true,
    }, [0.8, 0.2])
    addReady(repository, 'app-two', 'private', {
      ...VALID_METADATA,
      description: 'Private Snowflake architecture',
      technologies: ['Snowflake'],
    }, [1, 0])

    const service = new SlideRetrievalService({
      embedder: { embed: async () => [1, 0] },
      embeddingDimensions: 2,
      embeddingModel: 'embedding-model',
      repository,
    })
    const query = (overrides: Partial<SlideQueryInput>) => service.query('app-one', {
      facet_name: null,
      filters: EMPTY_FILTERS,
      limit: 20,
      mode: 'filter',
      query: null,
      select: ['identity'],
      similar_to_template_id: null,
      template_ids: [],
      ...overrides,
    })

    await expect(query({ mode: 'text', query: 'Snowflake' })).resolves.toMatchObject({
      mode_used: 'text', results: [{ templateId: 'architecture' }],
    })
    await expect(query({ mode: 'semantic', query: 'systems' })).resolves.toMatchObject({
      results: [{ templateId: 'architecture' }, { templateId: 'timeline' }],
    })
    const hybrid = await query({ mode: 'hybrid', query: 'Snowflake' })
    expect(hybrid).toMatchObject({ mode_used: 'hybrid' })
    if (hybrid.mode_used === 'facets') throw new Error('Expected slide results.')
    expect(hybrid.results[0]?.templateId).toBe('architecture')
    await expect(query({
      filters: { ...EMPTY_FILTERS, has_timeline: true },
      mode: 'filter',
    })).resolves.toMatchObject({ results: [{ templateId: 'timeline' }] })
    await expect(query({ mode: 'by_id', template_ids: ['timeline'] })).resolves.toMatchObject({
      results: [{ templateId: 'timeline', sections: { identity: expect.any(Object) } }],
    })
    await expect(query({ mode: 'similar', similar_to_template_id: 'architecture' })).resolves.toMatchObject({
      results: [{ templateId: 'timeline' }],
    })
    await expect(query({ facet_name: 'technologies', mode: 'facets' })).resolves.toEqual({
      facet: 'technologies',
      mode_used: 'facets',
      values: [{ count: 1, value: 'SAP' }, { count: 1, value: 'Snowflake' }],
    })
    expect(JSON.stringify(await query({ mode: 'filter', select: ['all'] }))).not.toContain('private')
    repository.close()
  })

  it('validates mode fields and visibly falls back when hybrid embeddings are unavailable', async () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repository.ensureApp('app-one')
    addReady(repository, 'app-one', 'architecture', VALID_METADATA, [1, 0])
    const service = new SlideRetrievalService({ repository })
    const base = {
      facet_name: null,
      filters: EMPTY_FILTERS,
      limit: 5,
      query: 'architecture',
      select: ['identity'],
      similar_to_template_id: null,
      template_ids: [],
    }
    await expect(service.query('app-one', { ...base, mode: 'semantic' }))
      .rejects.toMatchObject({ code: 'semantic_unavailable' })
    await expect(service.query('app-one', { ...base, mode: 'hybrid' }))
      .resolves.toMatchObject({ warnings: ['semantic_unavailable'] })
    await expect(service.query('app-one', { ...base, mode: 'text', template_ids: ['irrelevant'] }))
      .rejects.toBeInstanceOf(SlideRetrievalError)
    await expect(service.query('app-one', { ...base, mode: 'text', select: ['all', 'identity'] }))
      .rejects.toBeInstanceOf(SlideRetrievalError)
    repository.close()
  })

  it('uses deterministic reciprocal-rank fusion', () => {
    const record = (templateId: string, createdAt = '2026-01-01'): RetrievalClassificationRecord => ({
      classification: {
        attemptCount: 1,
        classifiedAt: createdAt,
        classifiedFingerprint: templateId,
        embeddingAttemptCount: 1,
        embeddingDimensions: null,
        embeddingDocument: null,
        embeddingFingerprint: null,
        embeddingLastErrorCode: null,
        embeddingModel: null,
        embeddingNextAttemptAt: null,
        embeddingStatus: 'not_ready',
        inputFingerprint: templateId,
        lastErrorCode: null,
        metadata: null,
        model: null,
        nextAttemptAt: null,
        promptVersion: 'prompt',
        schemaVersion: 1,
        status: 'ready',
        templateId,
        vector: null,
      },
      createdAt,
      kind: 'diagram',
      previewAvailable: false,
      templateId,
      title: templateId,
    })
    const a = record('a')
    const b = record('b')
    const c = record('c')
    expect(reciprocalRankFusion([a, b], [b, c]).map((item) => item.templateId)).toEqual(['b', 'a', 'c'])
  })
})

function addReady(
  repository: SqliteTemplateRepository,
  appId: string,
  templateId: string,
  rawMetadata: typeof VALID_METADATA,
  vector: readonly number[],
) {
  repository.insertWithPendingClassification({
    assets: [],
    metadata: {
      checksum: null,
      createdAt: `2026-01-0${templateId.length % 8 + 1}`,
      description: '',
      kind: 'diagram',
      source: 'import',
      templateId,
    },
    template: { templateId, templateJson: createTemplate(templateId) },
  }, { inputFingerprint: templateId, promptVersion: 'prompt', schemaVersion: 1 }, appId)
  const job = repository.claimNextSlideClassification('2026-01-01', '2026-01-02')
  if (!job || job.template.templateId !== templateId) throw new Error('Expected the inserted classification job.')
  const metadata = normalizeSlideRetrievalMetadata(rawMetadata)
  repository.completeSlideClassification({
    classifiedAt: '2026-01-01',
    classifiedFingerprint: templateId,
    embeddingDocument: `document ${templateId}`,
    embeddingFingerprint: `embedding ${templateId}`,
    metadata,
    model: 'classification-model',
    templateId,
  })
  const embedding = repository.claimNextSlideEmbedding('2026-01-01', '2026-01-02')
  if (!embedding || embedding.templateId !== templateId) throw new Error('Expected the inserted embedding job.')
  repository.completeSlideEmbedding(templateId, embedding.fingerprint, vector, 'embedding-model', 2)
}
