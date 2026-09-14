import request from 'supertest'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { createApp } from '../src/app'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { ImportTemplateService } from '../src/services/ImportTemplateService'
import { createTemplate } from './slideTestFixtures'

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

let repository: SqliteTemplateRepository

beforeEach(() => {
  repository = new SqliteTemplateRepository(':memory:')
})

afterEach(() => repository.close())

describe('POST /api/v2/import', () => {
  it('atomically stores one slide and durable pending work without invoking a provider', async () => {
    let notified = 0
    const service = new ImportTemplateService(
      { convertFile: async () => ({ templateJson: createImportTemplate(), warnings: [] }) },
      repository,
      () => 'template-v2',
      undefined,
      undefined,
      () => { notified += 1 },
    )
    const response = await request(createTestApp(service))
      .post('/api/v2/import?kind=commentary')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('synthetic PowerPoint'))
      .expect(201)

    expect(response.body).toEqual({
      previewAvailable: false,
      retrieval: { status: 'pending' },
      templateId: 'template-v2',
      templateJson: createImportTemplate(),
      warnings: ['A preview image could not be generated for this template.'],
    })
    expect(repository.findById('template-v2')).toBeDefined()
    expect(repository.findSlideClassification('template-v2')).toMatchObject({ status: 'pending' })
    expect(notified).toBe(1)
  })

  it('keeps v1 imports classification-free and exposes no other v2 method', async () => {
    const ids = ['v1-template', 'v2-template']
    const service = new ImportTemplateService(
      { convertFile: async () => ({ templateJson: createImportTemplate(), warnings: [] }) },
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

  it('rejects multi-slide decks without storing a template or pending job', async () => {
    const first = createTemplate().presentation.slides[0]
    if (!first) throw new Error('Expected fixture slide.')
    const service = new ImportTemplateService(
      {
        convertFile: async () => ({
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
})

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
