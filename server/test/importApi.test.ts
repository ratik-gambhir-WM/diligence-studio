// @vitest-environment node

import PptxGenJS from 'pptxgenjs'
import request from 'supertest'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { createApp } from '../src/app'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { ImportTemplateService } from '../src/services/ImportTemplateService'
import { LibraryPowerPointConverter } from '../src/services/PowerPointConverter'
import type { PowerPointCanvasJson } from '../src/lib/import/PowerpointImportTypes'

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'
const ONE_PIXEL_PNG =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII='

let templates: SqliteTemplateRepository

beforeEach(() => {
  templates = new SqliteTemplateRepository(':memory:')
})

afterEach(() => {
  templates.close()
})

describe('import API', () => {
  it('registers apps and keeps imported templates scoped to their app ID', async () => {
    const importService = new ImportTemplateService(
      {
        convertFile: async () => ({ templateJson: createLegacyTemplate(), warnings: [] }),
      },
      templates,
      () => 'tenant-template',
      () => 'tenant-asset',
    )
    const app = createTestApp(importService, 1024)

    await request(app)
      .post('/api/v1/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .set('X-App-Id', 'app-one')
      .send(Buffer.from('staged PowerPoint bytes'))
      .expect(201)

    expect(templates.findApp('app-one')).toMatchObject({
      appId: 'app-one',
      displayName: 'app-one',
      metadata: {},
    })
    expect((await request(app)
      .get('/api/v1/templates')
      .set('X-App-Id', 'app-one')
      .expect(200)).body.templates).toEqual([
      expect.objectContaining({ templateId: 'tenant-template' }),
    ])
    expect((await request(app)
      .get('/api/v1/templates')
      .set('X-App-Id', 'app-two')
      .expect(200)).body.templates).toEqual([])
    await request(app)
      .get('/api/v1/templates/tenant-template')
      .set('X-App-Id', 'app-two')
      .expect(404)
  })

  it('rejects malformed or conflicting app IDs', async () => {
    const app = createTestApp(
      new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      1024,
    )

    expect((await request(app)
      .get('/api/v1/templates')
      .set('X-App-Id', '../other-app')
      .expect(400)).body.error.code).toBe('invalid_app_id')
    expect((await request(app)
      .get('/api/v1/templates?appId=app-two')
      .set('X-App-Id', 'app-one')
      .expect(400)).body.error.code).toBe('conflicting_app_id')
  })

  it('imports every PowerPoint slide as a separate SQLite template', async () => {
    const templateIds = ['batch-template-1', 'batch-template-2', 'batch-template-3']
    const renderedSlideTexts: string[] = []
    const previewBytes = Buffer.from(ONE_PIXEL_PNG.split(',')[1] ?? '', 'base64')
    const importService = new ImportTemplateService(
      new LibraryPowerPointConverter(),
      templates,
      () => templateIds.shift() ?? 'unexpected-template-id',
      undefined,
      {
        generate: async ({ templateJson }) => {
          const firstElement = templateJson.presentation.slides[0]?.elements[0]
          if (firstElement?.type === 'text' && firstElement.text !== undefined) {
            renderedSlideTexts.push(firstElement.text)
          }
          return {
            bytes: previewBytes,
            contentType: 'image/png',
            height: 1,
            width: 1,
          }
        },
      },
    )
    const app = createTestApp(importService, 1024 * 1024)

    const imported = await request(app)
      .post('/api/v1/batchImport?kind=commentary')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(await createPowerPoint(['First batch slide', 'Second batch slide', 'Third batch slide']))
      .expect(201)

    expect(imported.headers).toMatchObject({
      'x-imported-template-count': '3',
      'x-powerpoint-warning-count': '0',
      'x-request-id': expect.any(String),
    })
    expect(imported.body.warnings).toEqual([])
    expect(imported.body.templates).toHaveLength(3)
    expect(imported.body.templates.map((item: { templateId: string }) => item.templateId)).toEqual([
      'batch-template-1',
      'batch-template-2',
      'batch-template-3',
    ])
    expect(imported.body.templates.map((item: { templateJson: PowerPointCanvasJson }) => (
      item.templateJson.presentation.slides.length
    ))).toEqual([1, 1, 1])
    expect(imported.body.templates.map((item: { templateJson: PowerPointCanvasJson }) => (
      item.templateJson.presentation.slides[0]?.elements[0]
    ))).toEqual([
      expect.objectContaining({ text: 'First batch slide', type: 'text' }),
      expect.objectContaining({ text: 'Second batch slide', type: 'text' }),
      expect.objectContaining({ text: 'Third batch slide', type: 'text' }),
    ])
    expect(renderedSlideTexts).toEqual([
      'First batch slide',
      'Second batch slide',
      'Third batch slide',
    ])

    for (const templateId of ['batch-template-1', 'batch-template-2', 'batch-template-3']) {
      expect(templates.findById(templateId)?.templateJson.presentation.slides).toHaveLength(1)
      expect(templates.findPreview(templateId)?.bytes).toEqual(previewBytes)
    }
    const listed = await request(app).get('/api/v1/templates?kind=commentary').expect(200)
    expect(listed.body.templates).toHaveLength(3)
    expect(listed.body.templates).toEqual(expect.arrayContaining([
      expect.objectContaining({
        templateId: 'batch-template-1',
        title: 'First batch slide',
        slideCount: 1,
      }),
      expect.objectContaining({
        templateId: 'batch-template-2',
        title: 'Second batch slide',
        slideCount: 1,
      }),
      expect.objectContaining({
        templateId: 'batch-template-3',
        title: 'Third batch slide',
        slideCount: 1,
      }),
    ]))
  })

  it('rolls back the complete batch when one SQLite insert fails', async () => {
    const importService = new ImportTemplateService(
      {
        convertFile: async () => ({ templateJson: createBatchTemplate(), warnings: [] }),
      },
      templates,
      () => 'duplicate-template-id',
    )
    const app = createTestApp(importService, 1024)

    const response = await request(app)
      .post('/api/v1/batchImport')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('staged PowerPoint bytes'))
      .expect(500)

    expect(response.body.error.code).toBe('internal_error')
    expect(templates.findById('duplicate-template-id')).toBeUndefined()
    expect(templates.list()).toEqual([])
  })

  it('keeps the original import endpoint restricted to a single slide', async () => {
    const app = createTestApp(
      new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      1024 * 1024,
    )

    const response = await request(app)
      .post('/api/v1/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(await createPowerPoint(['First slide', 'Second slide']))
      .expect(422)

    expect(response.body.error.code).toBe('template_must_have_one_slide')
    expect(templates.list()).toEqual([])
  })

  it('stores compact image references and retrieves canvas JSON with hydrated image data', async () => {
    const importService = new ImportTemplateService(
      new LibraryPowerPointConverter(),
      templates,
      () => 'template-123',
      () => 'asset-456',
    )
    const app = createTestApp(importService, 1024 * 1024)
    const source = await createPowerPoint()

    const imported = await request(app)
      .post('/api/v1/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(source)
      .expect(201)

    expect(imported.headers).toMatchObject({
      location: '/api/v1/templates/template-123?appId=DiligenceStudio_WestMonroe',
      'x-powerpoint-warning-count': '1',
      'x-request-id': expect.any(String),
      'x-template-id': 'template-123',
      'x-template-preview-status': 'unavailable',
    })
    expect(Object.keys(imported.body)).toEqual(['presentation'])
    expect(imported.body).toMatchObject({
      presentation: {
        preserveElementOrder: true,
        showBranding: false,
        slides: [
          {
            id: 'slide-1',
            elements: [
              expect.objectContaining({
                type: 'text',
                text: 'Stored by the import API',
                fill: 'transparent',
                stroke: 'transparent',
              }),
              expect.objectContaining({
                type: 'image',
                src: expect.stringMatching(/^data:image\/png;base64,/u),
              }),
            ],
          },
        ],
      },
    })
    expect(templates.findById('template-123')?.templateJson.presentation.slides[0]?.elements[1])
      .toMatchObject({ src: '/import/template-123/assets/asset-456' })

    const listed = await request(app).get('/api/v1/templates').expect(200)
    expect(listed.body).toEqual({
      templates: [{
        templateId: 'template-123',
        title: expect.any(String),
        description: 'Imported PowerPoint template',
        kind: 'diagram',
        previewUrl: null,
        slideCount: 1,
        elementCount: 2,
      }],
    })

    const retrieved = await request(app).get('/api/v1/templates/template-123').expect(200)
    expect(Object.keys(retrieved.body)).toEqual(['presentation'])
    expect(retrieved.body).toMatchObject({
      presentation: {
        slides: [
          {
            elements: [
              expect.objectContaining({ type: 'text' }),
              expect.objectContaining({
                type: 'image',
                src: expect.stringMatching(/^data:image\/png;base64,/u),
              }),
            ],
          },
        ],
      },
    })
    expect(JSON.stringify(retrieved.body)).toContain('base64')

    const legacyAlias = await request(app).get('/api/v1/import/template-123').expect(200)
    expect(legacyAlias.body).toEqual(retrieved.body)

    const image = await request(app)
      .get('/api/v1/import/template-123/assets/asset-456')
      .buffer(true)
      .parse((incoming, callback) => {
        const chunks: Buffer[] = []
        incoming.on('data', (chunk: Buffer) => chunks.push(chunk))
        incoming.on('end', () => callback(null, Buffer.concat(chunks)))
      })
      .expect(200)
    expect(image.headers).toMatchObject({
      'cache-control': 'no-store',
      'content-type': 'image/png',
      'x-content-type-options': 'nosniff',
    })
    expect(Buffer.isBuffer(image.body)).toBe(true)
    expect(image.body.subarray(1, 4).toString()).toBe('PNG')

    const exported = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send(retrieved.body)
      .buffer(true)
      .parse((incoming, callback) => {
        const chunks: Buffer[] = []
        incoming.on('data', (chunk: Buffer) => chunks.push(chunk))
        incoming.on('end', () => callback(null, Buffer.concat(chunks)))
      })
      .expect(200)
    expect(exported.headers['content-type']).toBe(POWERPOINT_CONTENT_TYPE)
    expect(exported.body.subarray(0, 2).toString()).toBe('PK')
  })

  it('rejects unsupported media types before conversion', async () => {
    const importService = new ImportTemplateService(new LibraryPowerPointConverter(), templates)
    const app = createTestApp(importService, 1024)

    for (const endpoint of ['/api/v1/import', '/api/v1/batchImport']) {
      const response = await request(app)
        .post(endpoint)
        .set('Content-Type', 'application/xml')
        .send('<presentation />')
        .expect(415)

      expect(response.body.error).toMatchObject({
        code: 'unsupported_media_type',
        message: expect.any(String),
        requestId: expect.any(String),
      })
    }
  })

  it('stores and serves a preview while filtering the catalog by kind', async () => {
    const previewBytes = Buffer.from(ONE_PIXEL_PNG.split(',')[1] ?? '', 'base64')
    const importService = new ImportTemplateService(
      new LibraryPowerPointConverter(),
      templates,
      () => 'commentary-template',
      () => 'unused-asset',
      {
        generate: async () => ({
          bytes: previewBytes,
          contentType: 'image/png',
          height: 1,
          width: 1,
        }),
      },
    )
    const app = createTestApp(importService, 1024 * 1024)

    const imported = await request(app)
      .post('/api/v1/import?kind=commentary')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(await createPowerPoint())
      .expect(201)

    expect(imported.headers).toMatchObject({
      link: '</api/v1/templates/commentary-template/preview?appId=DiligenceStudio_WestMonroe>; rel="preview"',
      'x-powerpoint-warning-count': '0',
      'x-template-preview-status': 'ready',
    })
    expect((await request(app).get('/api/v1/templates?kind=diagram').expect(200)).body.templates).toEqual([])
    expect((await request(app).get('/api/v1/templates?kind=commentary').expect(200)).body).toEqual({
      templates: [expect.objectContaining({
        kind: 'commentary',
        previewUrl: '/templates/commentary-template/preview?appId=DiligenceStudio_WestMonroe',
        templateId: 'commentary-template',
      })],
    })

    const preview = await request(app)
      .get('/api/v1/templates/commentary-template/preview')
      .buffer(true)
      .parse((incoming, callback) => {
        const chunks: Buffer[] = []
        incoming.on('data', (chunk: Buffer) => chunks.push(chunk))
        incoming.on('end', () => callback(null, Buffer.concat(chunks)))
      })
      .expect(200)
    expect(preview.headers).toMatchObject({
      'cache-control': 'private, no-store',
      'content-length': String(previewBytes.length),
      'content-type': 'image/png',
      'x-content-type-options': 'nosniff',
    })
    expect(preview.body).toEqual(previewBytes)

    await request(app).delete('/api/v1/templates/commentary-template').expect(204)
    expect(templates.findPreview('commentary-template')).toBeUndefined()
  })

  it('returns PNG previews in fixed pages of 10', async () => {
    const previewBytes = Buffer.from(ONE_PIXEL_PNG.split(',')[1] ?? '', 'base64')
    for (let index = 1; index <= 12; index += 1) {
      const templateId = `preview-template-${index}`
      templates.insert(
        { templateId, templateJson: createLegacyTemplate() },
        [],
        {
          checksum: null,
          createdAt: `2026-01-01T00:${String(index).padStart(2, '0')}:00.000Z`,
          description: 'Preview pagination test',
          kind: 'diagram',
          source: 'import',
          templateId,
        },
        {
          bytes: previewBytes,
          contentType: 'image/png',
          height: 1,
          templateId,
          width: 1,
        },
      )
    }
    templates.insert(
      { templateId: 'without-preview', templateJson: createLegacyTemplate() },
      [],
    )
    const app = createTestApp(
      new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      1024,
    )

    const firstPage = await request(app).get('/api/v1/templates/previews').expect(200)
    expect(firstPage.headers['cache-control']).toBe('private, no-store')
    expect(firstPage.body.pagination).toEqual({
      hasNextPage: true,
      hasPreviousPage: false,
      page: 1,
      pageSize: 10,
      totalItems: 12,
      totalPages: 2,
    })
    expect(firstPage.body.previews).toHaveLength(10)
    expect(firstPage.body.previews.map((item: { templateId: string }) => item.templateId)).toEqual([
      'preview-template-12',
      'preview-template-11',
      'preview-template-10',
      'preview-template-9',
      'preview-template-8',
      'preview-template-7',
      'preview-template-6',
      'preview-template-5',
      'preview-template-4',
      'preview-template-3',
    ])
    expect(firstPage.body.previews[0]).toEqual({
      contentType: 'image/png',
      dataUrl: ONE_PIXEL_PNG,
      height: 1,
      previewUrl: '/templates/preview-template-12/preview?appId=DiligenceStudio_WestMonroe',
      templateId: 'preview-template-12',
      width: 1,
    })

    const secondPage = await request(app).get('/api/v1/templates/previews?page=2').expect(200)
    expect(secondPage.body.pagination).toEqual({
      hasNextPage: false,
      hasPreviousPage: true,
      page: 2,
      pageSize: 10,
      totalItems: 12,
      totalPages: 2,
    })
    expect(secondPage.body.previews.map((item: { templateId: string }) => item.templateId)).toEqual([
      'preview-template-2',
      'preview-template-1',
    ])
  })

  it('rejects invalid preview pages and non-GET preview-list requests', async () => {
    const app = createTestApp(
      new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      1024,
    )

    expect((await request(app).get('/api/v1/templates/previews').expect(200)).body).toEqual({
      pagination: {
        hasNextPage: false,
        hasPreviousPage: false,
        page: 1,
        pageSize: 10,
        totalItems: 0,
        totalPages: 0,
      },
      previews: [],
    })
    for (const page of ['0', '-1', '1.5', 'not-a-number', '9007199254740992']) {
      const response = await request(app).get(`/api/v1/templates/previews?page=${page}`).expect(400)
      expect(response.body.error).toMatchObject({
        code: 'invalid_preview_page',
        message: 'Preview page must be a positive integer.',
        requestId: expect.any(String),
      })
    }
    await request(app).get('/api/v1/templates/previews?page=1&page=2').expect(400)
    await request(app).post('/api/v1/templates/previews').expect(405)
  })

  it('rejects invalid template kinds and distinguishes missing previews', async () => {
    templates.insert({ templateId: 'without-preview', templateJson: createLegacyTemplate() }, [])
    const app = createTestApp(
      new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      1024,
    )

    expect((await request(app).get('/api/v1/templates?kind=other').expect(400)).body.error.code)
      .toBe('invalid_template_kind')
    expect((await request(app).get('/api/v1/templates/without-preview/preview').expect(404)).body.error.code)
      .toBe('template_preview_not_found')
    expect((await request(app).get('/api/v1/templates/missing/preview').expect(404)).body.error.code)
      .toBe('template_not_found')
    await request(app).post('/api/v1/templates/without-preview/preview').expect(405)
  })

  it('returns a sanitized validation error for a malformed package', async () => {
    const importService = new ImportTemplateService(new LibraryPowerPointConverter(), templates)
    const app = createTestApp(importService, 1024)

    const response = await request(app)
      .post('/api/v1/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('not a zip'))
      .expect(422)

    expect(response.body.error).toMatchObject({
      code: 'invalid_powerpoint',
      message: 'The uploaded file is not a supported PowerPoint OOXML presentation.',
    })
    expect(JSON.stringify(response.body)).not.toContain('not a zip')
  })

  it('enforces the configured upload limit', async () => {
    const importService = new ImportTemplateService(new LibraryPowerPointConverter(), templates)
    const app = createTestApp(importService, 4)

    for (const endpoint of ['/api/v1/import', '/api/v1/batchImport']) {
      const response = await request(app)
        .post(endpoint)
        .set('Content-Type', POWERPOINT_CONTENT_TYPE)
        .send(Buffer.from('oversized'))
        .expect(413)

      expect(response.body.error.code).toBe('payload_too_large')
    }
  })

  it('cancels preview work and does not persist after a request timeout', async () => {
    let previewWasAborted = false
    const importService = new ImportTemplateService(
      {
        convertFile: async () => ({ templateJson: createLegacyTemplate(), warnings: [] }),
      },
      templates,
      () => 'timed-out-template',
      () => 'unused-asset',
      {
        generate: async (_request, signal) => new Promise((resolve) => {
          const finish = () => {
            previewWasAborted = true
            resolve(undefined)
          }
          if (signal?.aborted) finish()
          else signal?.addEventListener('abort', finish, { once: true })
        }),
      },
    )
    const app = createTestApp(importService, 1024, 10)

    await request(app)
      .post('/api/v1/import')
      .set('Content-Type', POWERPOINT_CONTENT_TYPE)
      .send(Buffer.from('staged PowerPoint bytes'))
      .expect(408)

    expect(previewWasAborted).toBe(true)
    expect(templates.findById('timed-out-template')).toBeUndefined()
  })

  it('returns a stable not-found error for an unknown template ID', async () => {
    const importService = new ImportTemplateService(new LibraryPowerPointConverter(), templates)
    const app = createTestApp(importService, 1024)

    const response = await request(app).get('/api/v1/import/missing').expect(404)
    expect(response.body.error.code).toBe('template_not_found')
  })

  it('does not expose an image outside its owning template', async () => {
    const importService = new ImportTemplateService(new LibraryPowerPointConverter(), templates)
    const app = createTestApp(importService, 1024)

    const response = await request(app)
      .get('/api/v1/import/missing/assets/missing')
      .expect(404)
    expect(response.body.error.code).toBe('template_asset_not_found')
  })

  it('returns a stable error when stored JSON references a missing joined asset', async () => {
    const missingAssetTemplate = createLegacyTemplate()
    const image = missingAssetTemplate.presentation.slides[0]?.elements[0]
    if (!image || image.type !== 'image') {
      throw new Error('Expected the test template to contain an image.')
    }
    image.src = '/import/orphan-template/assets/missing-asset'
    templates.insert({
      templateId: 'orphan-template',
      templateJson: missingAssetTemplate,
    }, [])
    const importService = new ImportTemplateService(new LibraryPowerPointConverter(), templates)
    const app = createTestApp(importService, 1024)

    const response = await request(app).get('/api/v1/import/orphan-template').expect(422)
    expect(response.body.error.code).toBe('template_asset_not_found')
  })

  it('migrates legacy embedded images when a stored template is first retrieved', async () => {
    templates.insert({
      templateId: 'legacy-template',
      templateJson: createLegacyTemplate(),
    }, [])
    const importService = new ImportTemplateService(
      new LibraryPowerPointConverter(),
      templates,
      () => 'unused-template-id',
      () => 'migrated-asset',
    )
    const app = createTestApp(importService, 1024)

    const response = await request(app).get('/api/v1/import/legacy-template').expect(200)
    expect(response.body.presentation.slides[0].elements[0]).toMatchObject({
      type: 'image',
      src: expect.stringMatching(/^data:image\/png;base64,/u),
    })
    expect(JSON.stringify(response.body)).toContain('base64')

    const stored = templates.findById('legacy-template')
    expect(JSON.stringify(stored)).not.toContain('base64')
    expect(stored?.templateJson.presentation.slides[0]?.elements[0]).toMatchObject({
      src: '/import/legacy-template/assets/migrated-asset',
    })
    await request(app)
      .get('/api/v1/import/legacy-template/assets/migrated-asset')
      .expect('Content-Type', 'image/png')
      .expect(200)
  })

  it('repairs legacy non-positive canvas dimensions and persists the migration', async () => {
    const templateJson = createLegacyTemplate()
    const slide = templateJson.presentation.slides[0]
    const image = slide?.elements[0]
    if (!slide || !image || image.type !== 'image') {
      throw new Error('Expected a slide with an image in the legacy fixture.')
    }
    slide.width = 0
    slide.height = -1
    image.w = 0
    image.h = -2
    templates.insert({ templateId: 'zero-sized-template', templateJson }, [])
    const app = createTestApp(
      new ImportTemplateService(
        new LibraryPowerPointConverter(),
        templates,
        () => 'unused-template-id',
        () => 'repaired-asset',
      ),
      1024,
    )

    const response = await request(app).get('/api/v1/templates/zero-sized-template').expect(200)
    expect(response.body.presentation.slides[0]).toMatchObject({ width: 1, height: 1 })
    expect(response.body.presentation.slides[0].elements[0]).toMatchObject({ w: 1, h: 1 })
    expect(templates.findById('zero-sized-template')?.templateJson.presentation.slides[0])
      .toMatchObject({ width: 1, height: 1 })
  })

  it('lists newest templates first and deletes templates with their assets', async () => {
    const first = createLegacyTemplate()
    first.presentation.title = 'First template'
    const firstSlide = first.presentation.slides[0]
    if (!firstSlide) {
      throw new Error('Expected a slide fixture.')
    }
    firstSlide.elements = []
    const second = createLegacyTemplate()
    second.presentation.title = 'Second template'
    const image = second.presentation.slides[0]?.elements[0]
    if (!image || image.type !== 'image') {
      throw new Error('Expected an image fixture.')
    }
    image.src = '/import/template-second/assets/asset-second'
    templates.insert({ templateId: 'template-first', templateJson: first }, [])
    templates.insert(
      { templateId: 'template-second', templateJson: second },
      [{
        assetId: 'asset-second',
        templateId: 'template-second',
        contentType: 'image/png',
        bytes: Buffer.from([1, 2, 3]),
      }],
    )
    const app = createTestApp(
      new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      1024,
    )

    const listed = await request(app).get('/api/v1/templates').expect(200)
    expect(listed.body).toEqual({
      templates: [
        {
          description: 'Imported PowerPoint template',
          templateId: 'template-second',
          kind: 'diagram',
          previewUrl: null,
          title: 'Second template',
          slideCount: 1,
          elementCount: 1,
        },
        {
          description: 'Imported PowerPoint template',
          templateId: 'template-first',
          kind: 'diagram',
          previewUrl: null,
          title: 'First template',
          slideCount: 1,
          elementCount: 0,
        },
      ],
    })

    const deleted = await request(app).delete('/api/v1/templates/template-second').expect(204)
    expect(deleted.headers['x-request-id']).toEqual(expect.any(String))
    expect(deleted.text).toBe('')
    expect(templates.findById('template-second')).toBeUndefined()
    expect(templates.findAsset('template-second', 'asset-second')).toBeUndefined()

    const repeated = await request(app).delete('/api/v1/templates/template-second').expect(404)
    expect(repeated.body.error.code).toBe('template_not_found')
  })
})

function createTestApp(
  importService: ImportTemplateService,
  maxUploadBytes: number,
  requestTimeoutMs?: number,
) {
  return createApp({
    exportService: new ExportPowerPointService(templates),
    importService,
    maxExportJsonBytes: 1024,
    maxUploadBytes,
    requestTimeoutMs,
  })
}

async function createPowerPoint(slideTexts: readonly string[] = ['Stored by the import API']) {
  const presentation = new PptxGenJS()
  presentation.layout = 'LAYOUT_WIDE'
  for (const [index, slideText] of slideTexts.entries()) {
    const slide = presentation.addSlide()
    slide.addText(slideText, {
      x: 0.5,
      y: 0.5,
      w: 4,
      h: 0.5,
    })
    if (index === 0 && slideTexts.length === 1) {
      slide.addImage({
        data: ONE_PIXEL_PNG,
        x: 0.5,
        y: 1.5,
        w: 1,
        h: 1,
      })
    }
  }
  const output = await presentation.write({ outputType: 'nodebuffer' })
  if (!Buffer.isBuffer(output)) {
    throw new Error('Expected PptxGenJS to return a Node.js buffer.')
  }
  return output
}

function createBatchTemplate(): PowerPointCanvasJson {
  const first = createLegacyTemplate().presentation.slides[0]
  if (!first) {
    throw new Error('Expected a slide fixture.')
  }
  return {
    presentation: {
      title: 'Batch template',
      preserveElementOrder: true,
      showBranding: false,
      slides: [
        { ...first, id: 'slide-1', name: 'First slide', elements: [] },
        { ...first, id: 'slide-2', name: 'Second slide', elements: [] },
      ],
    },
  }
}

function createLegacyTemplate(): PowerPointCanvasJson {
  return {
    presentation: {
      title: 'Legacy template',
      preserveElementOrder: true,
      showBranding: false,
      slides: [
        {
          id: 'slide-1',
          name: 'Legacy slide',
          width: 1280,
          height: 720,
          backgroundColor: 'FFFFFF',
          elements: [
            {
              id: 'legacy-image',
              type: 'image',
              x: 0,
              y: 0,
              w: 100,
              h: 100,
              src: ONE_PIXEL_PNG,
              fit: 'contain',
            },
          ],
        },
      ],
    },
  }
}
