import { describe, expect, it } from 'vitest'

import {
  buildSlideClassificationInput,
  buildSlideClassificationInputFingerprint,
} from '../src/lib/retrieval/SlideClassificationInput'
import {
  normalizeSlideRetrievalMetadata,
  SlideRetrievalMetadataSchema,
} from '../src/lib/retrieval/SlideRetrievalMetadata'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { SlideClassificationService } from '../src/services/SlideClassificationService'
import { SlideClassificationWorker } from '../src/services/SlideClassificationWorker'
import { createTemplate, VALID_METADATA } from './slideTestFixtures'

describe('slide classification contracts', () => {
  it('strictly validates and deterministically normalizes metadata', () => {
    expect(SlideRetrievalMetadataSchema.safeParse({ ...VALID_METADATA, extra: true }).success).toBe(false)
    expect(normalizeSlideRetrievalMetadata(VALID_METADATA)).toMatchObject({
      slide_type: 'architecture-overview',
      layout_type: 'diagram-led',
      slide_purpose: 'Explain the platform architecture',
      topics: ['Architecture', 'Integrations'],
    })
    expect(() => normalizeSlideRetrievalMetadata({ ...VALID_METADATA, has_chart: undefined })).toThrow()
    expect(() => normalizeSlideRetrievalMetadata({ ...VALID_METADATA, topics: [''] })).toThrow()
  })

  it('builds a bounded digest that preserves facts but excludes image bytes', () => {
    const input = buildSlideClassificationInput({
      kind: 'diagram',
      limits: { maxDigestBytes: 4_000, maxPreviewBytes: 10 },
      preview: { bytes: Buffer.alloc(11), contentType: 'image/png', height: 10, width: 10 },
      templateJson: createTemplate(),
    })
    expect(input.digest).toContain('Customer data flow')
    expect(input.digest).toContain('connected-lines=1')
    expect(input.digest).toContain('alt=Sensitive diagram')
    expect(input.digest).not.toContain('DO_NOT_INCLUDE')
    expect(input.preview).toBeUndefined()
  })

  it('creates stable fingerprints that change with preview or slide content', () => {
    const template = createTemplate()
    expect(buildSlideClassificationInputFingerprint(template, 'a'))
      .toBe(buildSlideClassificationInputFingerprint(template, 'a'))
    expect(buildSlideClassificationInputFingerprint(template, 'a'))
      .not.toBe(buildSlideClassificationInputFingerprint(template, 'b'))
    expect(buildSlideClassificationInputFingerprint(template, 'a'))
      .not.toBe(buildSlideClassificationInputFingerprint(createTemplate('Other'), 'a'))
  })

  it('migrates without backfilling and atomically creates and claims pending work', () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repository.ensureApp('app-one')
    repository.insert({ templateId: 'v1', templateJson: createTemplate('V1') }, [], undefined, undefined, 'app-one')
    expect(repository.findSlideClassification('v1')).toBeUndefined()
    repository.insertWithPendingClassification({
      assets: [],
      template: { templateId: 'v2', templateJson: createTemplate('V2') },
    }, { inputFingerprint: 'fingerprint', promptVersion: 'prompt', schemaVersion: 1 }, 'app-one')
    expect(repository.findSlideClassification('v2')).toMatchObject({
      attemptCount: 0,
      embeddingStatus: 'not_ready',
      status: 'pending',
    })
    const job = repository.claimNextSlideClassification(
      new Date(0).toISOString(),
      new Date(60_000).toISOString(),
    )
    expect(job).toMatchObject({ attemptCount: 1, inputFingerprint: 'fingerprint' })
    expect(repository.claimNextSlideClassification(new Date(0).toISOString(), new Date(60_000).toISOString()))
      .toBeUndefined()
    repository.close()
  })

  it('processes classification and embedding as independent durable worker stages', async () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repository.ensureApp('app-one')
    repository.insertWithPendingClassification({
      assets: [],
      template: { templateId: 'worker-slide', templateJson: createTemplate() },
    }, { inputFingerprint: 'input', promptVersion: 'prompt', schemaVersion: 1 }, 'app-one')
    let classifications = 0
    let embeddings = 0
    const service = new SlideClassificationService({
      classificationModel: 'classification-model',
      classifier: {
        classify: async () => {
          classifications += 1
          return normalizeSlideRetrievalMetadata(VALID_METADATA)
        },
      },
      embedder: {
        embed: async () => {
          embeddings += 1
          return [1, 0]
        },
      },
      embeddingDimensions: 2,
      embeddingMaxTextBytes: 10_000,
      embeddingModel: 'embedding-model',
      inputLimits: {},
      repository,
    })
    const worker = new SlideClassificationWorker({
      classificationMaxAttempts: 2,
      classificationTimeoutMs: 100,
      concurrency: 1,
      embeddingMaxAttempts: 2,
      embeddingTimeoutMs: 100,
      repository,
      retryBaseMs: 1,
      service,
    })
    worker.start()
    await waitFor(() => repository.findSlideClassification('worker-slide')?.embeddingStatus === 'ready')
    await worker.stop()
    expect({ classifications, embeddings }).toEqual({ classifications: 1, embeddings: 1 })
    expect(repository.findSlideClassification('worker-slide')).toMatchObject({
      embeddingDimensions: 2,
      embeddingModel: 'embedding-model',
      embeddingStatus: 'ready',
      status: 'ready',
      vector: [1, 0],
    })
    repository.close()
  })

  it('rejects stale completion after a fingerprint refresh without erasing usable state', () => {
    const repository = new SqliteTemplateRepository(':memory:')
    repository.ensureApp('app-one')
    repository.insertWithPendingClassification({
      assets: [],
      template: { templateId: 'refresh-slide', templateJson: createTemplate() },
    }, { inputFingerprint: 'old-input', promptVersion: 'old-prompt', schemaVersion: 1 }, 'app-one')
    const claimed = repository.claimNextSlideClassification('2026-01-01', '2026-01-02')
    if (!claimed) throw new Error('Expected a classification job.')
    expect(repository.refreshPendingSlideClassification('refresh-slide', {
      inputFingerprint: 'new-input', promptVersion: 'new-prompt', schemaVersion: 2,
    })).toBe(true)
    expect(() => repository.completeSlideClassification({
      classifiedAt: '2026-01-01',
      classifiedFingerprint: claimed.inputFingerprint,
      embeddingDocument: 'document',
      embeddingFingerprint: 'embedding',
      metadata: normalizeSlideRetrievalMetadata(VALID_METADATA),
      model: 'classification-model',
      templateId: 'refresh-slide',
    })).toThrow('lease is no longer active')
    expect(repository.findSlideClassification('refresh-slide')).toMatchObject({
      inputFingerprint: 'new-input',
      metadata: null,
      status: 'pending',
    })
    repository.close()
  })
})

async function waitFor(predicate: () => boolean) {
  const deadline = Date.now() + 1_000
  while (!predicate()) {
    if (Date.now() > deadline) throw new Error('Timed out waiting for worker state.')
    await new Promise((resolve) => setTimeout(resolve, 5))
  }
}
