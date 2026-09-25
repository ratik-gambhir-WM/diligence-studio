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
  archetypes: [],
  content_density: [],
  domains: [],
  intents: [],
  kinds: [],
  slot_roles: [],
  topics: [],
}

describe('slide retrieval service', () => {
  it('supports app-scoped lexical, dual semantic, filters, IDs, similarity, and facets', async () => {
    const repository = new SqliteTemplateRepository()
    repository.ensureApp('app-one')
    repository.ensureApp('app-two')
    addReady(repository, 'app-one', 'cyber-finding', VALID_METADATA, [1, 0], [0, 1])
    addReady(repository, 'app-one', 'data-summary', {
      ...VALID_METADATA,
      subject: {
        ...VALID_METADATA.subject,
        domains: [{ id: 'data-ai' as const, relevance: 'primary' as const, topics: ['data-platform'] }],
        summary: 'Data platform modernization roadmap',
        technologies: ['Databricks'],
      },
      communication: { ...VALID_METADATA.communication, intents: ['current-state' as const] },
      template_fit: { ...VALID_METADATA.template_fit, archetype: 'Current State Summary' },
    }, [0.8, 0.2], [0.2, 0.8])
    addReady(repository, 'app-two', 'private', VALID_METADATA, [1, 0], [0, 1])

    const service = new SlideRetrievalService({
      embedder: { embed: async (text) => text.includes('communication intents') ? [0, 1] : [1, 0] },
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
      select: ['subject'],
      semantic_target: 'both',
      similar_to_template_id: null,
      template_ids: [],
      ...overrides,
    })

    const lexical = await query({ mode: 'text', query: 'security testing' })
    if (lexical.mode_used === 'facets') throw new Error('Expected slide results.')
    expect(lexical.results[0]?.templateId).toBe('cyber-finding')
    await expect(query({ mode: 'semantic', query: 'cyber risk', semantic_target: 'subject' }))
      .resolves.toMatchObject({ results: [{ templateId: 'cyber-finding' }, { templateId: 'data-summary' }] })
    await expect(query({
      filters: { ...EMPTY_FILTERS, domains: ['data-ai'] },
      mode: 'filter',
    })).resolves.toMatchObject({ results: [{ templateId: 'data-summary' }] })
    await expect(query({ mode: 'by_id', template_ids: ['data-summary'] })).resolves.toMatchObject({
      results: [{ templateId: 'data-summary', sections: { subject: expect.any(Object) } }],
    })
    await expect(query({ mode: 'similar', similar_to_template_id: 'cyber-finding' }))
      .resolves.toMatchObject({ results: [{ templateId: 'data-summary' }] })
    await expect(query({ facet_name: 'domains', mode: 'facets' })).resolves.toEqual({
      facet: 'domains',
      mode_used: 'facets',
      values: [
        { count: 1, value: 'cybersecurity' },
        { count: 1, value: 'data-ai' },
        { count: 1, value: 'software-architecture' },
      ],
    })
    expect(JSON.stringify(await query({ mode: 'filter', select: ['all'] }))).not.toContain('private')
    repository.close()
  })

  it('finds a template for markdown using exact metadata and both vectors', async () => {
    const repository = new SqliteTemplateRepository()
    repository.ensureApp('app-one')
    addReady(repository, 'app-one', 'cyber-finding', VALID_METADATA, [1, 0], [0, 1])
    addReady(repository, 'app-one', 'generic-data', {
      ...VALID_METADATA,
      subject: {
        ...VALID_METADATA.subject,
        domains: [{ id: 'data-ai' as const, relevance: 'primary' as const, topics: ['analytics'] }],
      },
    }, [1, 0], [0, 1])
    const service = new SlideRetrievalService({
      embedder: { embed: async (text) => text.includes('communication intents') ? [0, 1] : [1, 0] },
      embeddingDimensions: 2,
      embeddingModel: 'embedding-model',
      repository,
    })

    const finding = await service.findForFinding('app-one', {
      markdown: '# Finding\nSecurity testing has a material control gap.\n## Evidence\nCoverage is incomplete.\n## Recommendation\nExpand testing.',
      limit: 2,
    })
    expect(finding.results[0]).toMatchObject({
      matchedDomains: ['cybersecurity'],
      matchedIntents: ['finding', 'evidence', 'recommendation'],
      matchedTopics: ['security-testing'],
      templateId: 'cyber-finding',
    })
    expect(finding.results.find((result) => result.templateId === 'generic-data')?.matchedDomains).toEqual([])
    repository.close()
  })

  it('validates mode fields and visibly falls back when embeddings are unavailable', async () => {
    const repository = new SqliteTemplateRepository()
    repository.ensureApp('app-one')
    addReady(repository, 'app-one', 'cyber-finding', VALID_METADATA, [1, 0], [0, 1])
    const service = new SlideRetrievalService({ repository })
    const base = {
      facet_name: null,
      filters: EMPTY_FILTERS,
      limit: 5,
      query: 'security',
      select: ['subject'],
      semantic_target: 'both',
      similar_to_template_id: null,
      template_ids: [],
    }
    await expect(service.query('app-one', { ...base, mode: 'semantic' }))
      .rejects.toMatchObject({ code: 'semantic_unavailable' })
    await expect(service.query('app-one', { ...base, mode: 'hybrid' }))
      .resolves.toMatchObject({ warnings: ['semantic_unavailable'] })
    await expect(service.query('app-one', { ...base, mode: 'text', template_ids: ['irrelevant'] }))
      .rejects.toBeInstanceOf(SlideRetrievalError)
    repository.close()
  })

  it('uses deterministic reciprocal-rank fusion', () => {
    const record = (templateId: string): RetrievalClassificationRecord => ({
      classification: {
        attemptCount: 1,
        capabilityEmbeddingDocument: null,
        capabilityEmbeddingFingerprint: null,
        capabilityVector: null,
        classifiedAt: '2026-01-01',
        classifiedFingerprint: templateId,
        embeddingAttemptCount: 1,
        embeddingDimensions: null,
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
        schemaVersion: 2,
        status: 'ready',
        subjectEmbeddingDocument: null,
        subjectEmbeddingFingerprint: null,
        subjectVector: null,
        templateId,
      },
      createdAt: '2026-01-01',
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
  rawMetadata: unknown,
  subjectVector: readonly number[],
  capabilityVector: readonly number[],
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
  }, { inputFingerprint: templateId, promptVersion: 'prompt', schemaVersion: 2 }, appId)
  const job = repository.claimNextSlideClassification('2026-01-01', '2026-01-02')
  if (!job || job.template.templateId !== templateId) throw new Error('Expected the inserted classification job.')
  const metadata = normalizeSlideRetrievalMetadata(rawMetadata)
  repository.completeSlideClassification({
    capabilityEmbeddingDocument: `capability ${templateId}`,
    capabilityEmbeddingFingerprint: `capability fingerprint ${templateId}`,
    classifiedAt: '2026-01-01',
    classifiedFingerprint: templateId,
    metadata,
    model: 'classification-model',
    subjectEmbeddingDocument: `subject ${templateId}`,
    subjectEmbeddingFingerprint: `subject fingerprint ${templateId}`,
    templateId,
  })
  const embedding = repository.claimNextSlideEmbedding('2026-01-01', '2026-01-02')
  if (!embedding || embedding.templateId !== templateId) throw new Error('Expected the inserted embedding job.')
  repository.completeSlideEmbedding(
    templateId,
    embedding.subjectFingerprint,
    embedding.capabilityFingerprint,
    subjectVector,
    capabilityVector,
    'embedding-model',
    2,
  )
}
