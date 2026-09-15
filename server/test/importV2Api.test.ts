import request from 'supertest'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { createApp } from '../src/app'
import {
  SlideProviderError,
  type SlideClassifier,
  type SlideEmbedder,
} from '../src/integrations/SlideProvider'
import { normalizeSlideRetrievalMetadata } from '../src/lib/retrieval/SlideRetrievalMetadata'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { ImportTemplateService } from '../src/services/ImportTemplateService'
import { createSlideAgentTools } from '../src/services/SlideAgentTools'
import { SlideClassificationService } from '../src/services/SlideClassificationService'
import { SlideRetrievalService } from '../src/services/SlideRetrievalService'
import { createTemplate, VALID_METADATA } from './slideTestFixtures'

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

let repository: SqliteTemplateRepository

beforeEach(() => {
  repository = new SqliteTemplateRepository()
})

afterEach(() => repository.close())

describe('POST /api/v2/import', () => {
  it('waits for classification and embedding, then atomically publishes ready retrieval data', async () => {
    const providerCalls: string[] = []
    const processingService = createProcessingService({
      classify: async () => {
        providerCalls.push('classify')
        expect(repository.findById('template-v2')).toBeDefined()
        expect(repository.findSlideClassification('template-v2')).toMatchObject({
          metadata: null,
          status: 'processing',
        })
        return normalizeSlideRetrievalMetadata(VALID_METADATA)
      },
      embed: async () => {
        providerCalls.push('embed')
        expect(repository.findSlideClassification('template-v2')).toMatchObject({
          embeddingStatus: 'not_ready',
          metadata: null,
        })
        return [1, 0]
      },
    })
    const service = new ImportTemplateService(
      { convert: async () => ({ templateJson: createImportTemplate(), warnings: [] }) },
      repository,
      () => 'template-v2',
      undefined,
      undefined,
      processingService,
    )
    const response = await request(createTestApp(service))
      .post('/api/v2/import?kind=commentary')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('synthetic PowerPoint'))
      .expect(201)

    expect(response.body).toEqual({
      previewAvailable: false,
      retrieval: { status: 'ready' },
      templateId: 'template-v2',
      templateJson: createImportTemplate(),
      warnings: ['A preview image could not be generated for this template.'],
    })
    expect(providerCalls).toEqual(['classify', 'embed', 'embed'])
    expect(repository.findSlideClassification('template-v2')).toMatchObject({
      embeddingDimensions: 2,
      embeddingModel: 'embedding-model',
      embeddingStatus: 'ready',
      capabilityVector: [1, 0],
      metadata: { subject: { domains: expect.arrayContaining([{ id: 'cybersecurity', relevance: 'primary', topics: ['security-testing'] }]) } },
      status: 'ready',
      subjectVector: [1, 0],
    })
    expect(repository.searchSlideClassifications(
      'DiligenceStudio_WestMonroe',
      '"architecture"*',
      10,
    )).toEqual(['template-v2'])
    const agentTools = createSlideAgentTools(
      'DiligenceStudio_WestMonroe',
      new SlideRetrievalService({
        embedder: { embed: async () => [1, 0] },
        embeddingDimensions: 2,
        embeddingModel: 'embedding-model',
        repository,
      }),
      service,
    )
    await expect(agentTools.query_slides({
      facet_name: null,
      filters: {
        archetypes: [], content_density: [], domains: [], intents: [], kinds: [], slot_roles: [], topics: [],
      },
      limit: 10,
      mode: 'semantic',
      query: 'architecture',
      select: ['subject'],
      semantic_target: 'both',
      similar_to_template_id: null,
      template_ids: [],
    })).resolves.toMatchObject({ results: [{ templateId: 'template-v2' }] })
    await expect(agentTools.find_slides_for_finding({
      markdown: '# Finding\nSecurity testing coverage is incomplete.\n## Evidence\nA control gap exists.\n## Recommendation\nExpand testing.',
      limit: 5,
    })).resolves.toMatchObject({
      results: [{ matchedDomains: ['cybersecurity'], templateId: 'template-v2' }],
    })
    expect(agentTools.get_slide({ template_id: 'template-v2' })).toMatchObject({
      templateId: 'template-v2',
    })
  })

  it('keeps v1 imports classification-free and exposes no other v2 method', async () => {
    const ids = ['v1-template', 'v2-template']
    const service = new ImportTemplateService(
      { convert: async () => ({ templateJson: createImportTemplate(), warnings: [] }) },
      repository,
      () => ids.shift() ?? 'unexpected',
    )
    const app = createTestApp(service)
    await request(app).post('/api/v1/import').set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('v1')).expect(201)
    expect(repository.findSlideClassification('v1-template')).toBeUndefined()
    await request(app).get('/api/v2/import').expect(405)
    await request(app).post('/api/v2/query').expect(404)
  })

  it('rejects multi-slide decks without storing a template or processing record', async () => {
    const first = createTemplate().presentation.slides[0]
    if (!first) throw new Error('Expected fixture slide.')
    const service = new ImportTemplateService(
      {
        convert: async () => ({
          templateJson: {
            ...createTemplate(),
            presentation: {
              ...createTemplate().presentation,
              slides: [first, { ...first, id: 'slide-2' }],
            },
          },
          warnings: [],
        }),
      },
      repository,
      () => 'not-stored',
      undefined,
      undefined,
      createProcessingService(),
    )
    const response = await request(createTestApp(service))
      .post('/api/v2/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('two slides'))
      .expect(422)
    expect(response.body.error.code).toBe('template_must_have_one_slide')
    expect(repository.findById('not-stored')).toBeUndefined()
    expect(repository.findSlideClassification('not-stored')).toBeUndefined()
  })

  it('fails before conversion when synchronous retrieval processing is not configured', async () => {
    let converted = false
    const service = new ImportTemplateService(
      { convert: async () => {
        converted = true
        return { templateJson: createImportTemplate(), warnings: [] }
      } },
      repository,
      () => 'not-stored',
    )
    const response = await request(createTestApp(service))
      .post('/api/v2/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('synthetic PowerPoint'))
      .expect(503)
    expect(response.body.error.code).toBe('slide_classification_unavailable')
    expect(converted).toBe(false)
    expect(repository.findById('not-stored')).toBeUndefined()
  })

  it('does not publish partial metadata when embedding fails', async () => {
    const service = new ImportTemplateService(
      { convert: async () => ({ templateJson: createImportTemplate(), warnings: [] }) },
      repository,
      () => 'failed-embedding',
      undefined,
      undefined,
      createProcessingService({
        embed: async () => {
          throw new SlideProviderError('embedding_unavailable', true)
        },
      }),
    )
    const response = await request(createTestApp(service))
      .post('/api/v2/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('synthetic PowerPoint'))
      .expect(503)
    expect(response.body.error.code).toBe('embedding_unavailable')
    expect(repository.findById('failed-embedding')).toBeDefined()
    expect(repository.findSlideClassification('failed-embedding')).toMatchObject({
      lastErrorCode: 'embedding_unavailable',
      metadata: null,
      status: 'failed',
      capabilityVector: null,
      subjectVector: null,
    })
    expect(repository.searchSlideClassifications(
      'DiligenceStudio_WestMonroe',
      '"architecture"*',
      10,
    )).toEqual([])
  })
})

function createProcessingService(overrides: {
  classify?: SlideClassifier['classify']
  embed?: SlideEmbedder['embed']
} = {}) {
  return new SlideClassificationService({
    classificationModel: 'classification-model',
    classifier: {
      classify: overrides.classify ?? (async () => normalizeSlideRetrievalMetadata(VALID_METADATA)),
    },
    embedder: { embed: overrides.embed ?? (async () => [1, 0]) },
    embeddingDimensions: 2,
    embeddingMaxTextBytes: 10_000,
    embeddingModel: 'embedding-model',
    inputLimits: {},
    repository,
  })
}

function createTestApp(importService: ImportTemplateService) {
  return createApp({
    exportService: new ExportPowerPointService(repository),
    importService,
    maxExportJsonBytes: 1024,
    maxUploadBytes: 1024,
  })
}

function createImportTemplate() {
  const template = createTemplate()
  return {
    ...template,
    presentation: {
      ...template.presentation,
      slides: template.presentation.slides.map((slide) => ({
        ...slide,
        elements: slide.elements.filter((element) => element.type !== 'image'),
      })),
    },
  }
}
