// @vitest-environment node

import request from 'supertest'
import { afterEach, describe, expect, it } from 'vitest'

import { createApp } from '../src/app'
import { normalizeSlideRetrievalMetadata } from '../src/lib/retrieval/SlideRetrievalMetadata'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { ImportTemplateService } from '../src/services/ImportTemplateService'
import { createTemplate, VALID_METADATA } from './slideTestFixtures'

const repositories: SqliteTemplateRepository[] = []

afterEach(() => repositories.splice(0).forEach((repository) => repository.close()))

describe('GET /api/v2/templates/:templateId', () => {
  it('returns only the template ID unless retrieval fields are explicitly requested', async () => {
    const repository = createRepository()
    addReadyEmbedding(repository, 'app-one', 'security-slide')
    const app = createTestApp(repository)

    await request(app)
      .get('/api/v2/templates/security-slide')
      .set('X-App-Id', 'app-one')
      .expect('Cache-Control', 'private, no-store')
      .expect(200, { templateId: 'security-slide' })

    await request(app)
      .get('/api/v2/templates/security-slide?metadata=false&embeddings=false')
      .set('X-App-Id', 'app-one')
      .expect(200, { templateId: 'security-slide' })
  })

  it('returns normalized metadata and raw subject and capability vectors when requested', async () => {
    const repository = createRepository()
    addReadyEmbedding(repository, 'app-one', 'security-slide')

    const response = await request(createTestApp(repository))
      .get('/api/v2/templates/security-slide?metadata=true&embeddings=true')
      .set('X-App-Id', 'app-one')
      .expect(200)

    expect(response.body).toEqual({
      embeddings: {
        capabilityVector: [0.25, 0.75],
        dimensions: 2,
        model: 'embedding-model',
        status: 'ready',
        subjectVector: [1, 0],
      },
      metadata: normalizeSlideRetrievalMetadata(VALID_METADATA),
      templateId: 'security-slide',
    })
  })

  it('does not expose another app\'s vectors and returns null values for unclassified templates', async () => {
    const repository = createRepository()
    addReadyEmbedding(repository, 'app-one', 'security-slide')
    repository.ensureApp('app-two')
    repository.insert({ templateId: 'unclassified', templateJson: createTemplate() }, [], undefined, undefined, 'app-two')
    const app = createTestApp(repository)

    await request(app)
      .get('/api/v2/templates/security-slide?embeddings=true')
      .set('X-App-Id', 'app-two')
      .expect(404)

    await request(app)
      .get('/api/v2/templates/unclassified?metadata=true&embeddings=true')
      .set('X-App-Id', 'app-two')
      .expect(200, {
        embeddings: {
          capabilityVector: null,
          dimensions: null,
          model: null,
          status: 'not_ready',
          subjectVector: null,
        },
        metadata: null,
        templateId: 'unclassified',
      })
  })

  it('rejects malformed inspection flags and confirms the former mock endpoint is gone', async () => {
    const repository = createRepository()
    repository.insert({ templateId: 'security-slide', templateJson: createTemplate() }, [])
    const app = createTestApp(repository)

    for (const query of ['metadata=yes', 'embeddings=true&embeddings=false', 'unknown=true']) {
      const response = await request(app).get(`/api/v2/templates/security-slide?${query}`).expect(400)
      expect(response.body.error.code).toBe('invalid_template_v2_query')
    }

    const removed = await request(app).get('/api/v2/mock/embeddings').expect(404)
    expect(removed.body.error.code).toBe('route_not_found')
  })
})

function createRepository() {
  const repository = new SqliteTemplateRepository()
  repositories.push(repository)
  return repository
}

function createTestApp(repository: SqliteTemplateRepository) {
  return createApp({
    exportService: new ExportPowerPointService(repository),
    importService: new ImportTemplateService({ convert: async () => {
      throw new Error('The import service is not used by this test.')
    } }, repository),
    maxExportJsonBytes: 1024,
    maxUploadBytes: 1024,
  })
}

function addReadyEmbedding(repository: SqliteTemplateRepository, appId: string, templateId: string) {
  repository.ensureApp(appId)
  repository.insertWithPendingClassification(
    {
      assets: [],
      metadata: {
        checksum: null,
        createdAt: '2026-01-01T00:00:00.000Z',
        description: 'Test slide',
        kind: 'diagram',
        source: 'import',
        templateId,
      },
      template: { templateId, templateJson: createTemplate('Security slide') },
    },
    { inputFingerprint: `${templateId}-input`, promptVersion: 'prompt', schemaVersion: 2 },
    appId,
  )
  const classification = repository.claimNextSlideClassification('2026-01-01', '2026-01-02')
  if (!classification) throw new Error('Expected a classification job.')
  repository.completeSlideClassification({
    capabilityEmbeddingDocument: 'capability document',
    capabilityEmbeddingFingerprint: `${templateId}-capability`,
    classifiedAt: '2026-01-01T00:00:00.000Z',
    classifiedFingerprint: classification.inputFingerprint,
    metadata: normalizeSlideRetrievalMetadata(VALID_METADATA),
    model: 'classification-model',
    subjectEmbeddingDocument: 'subject document',
    subjectEmbeddingFingerprint: `${templateId}-subject`,
    templateId,
  })
  const embedding = repository.claimNextSlideEmbedding('2026-01-01', '2026-01-02')
  if (!embedding) throw new Error('Expected an embedding job.')
  repository.completeSlideEmbedding(
    templateId,
    embedding.subjectFingerprint,
    embedding.capabilityFingerprint,
    [1, 0],
    [0.25, 0.75],
    'embedding-model',
    2,
  )
}
