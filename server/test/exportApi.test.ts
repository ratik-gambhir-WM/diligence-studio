// @vitest-environment node

import JSZip from 'jszip'
import PptxGenJS from 'pptxgenjs'
import request from 'supertest'
import { afterEach, beforeEach, describe, expect, it } from 'vitest'

import { createApp } from '../src/app'
import type { ApiRequestLog } from '../src/apiLogging'
import type { PowerPointCanvasJson } from '../src/lib/import/PowerpointImportTypes'
import { SqliteTemplateRepository } from '../src/repositories/SqliteTemplateRepository'
import { ExportPowerPointService } from '../src/services/ExportPowerPointService'
import { ImportTemplateService } from '../src/services/ImportTemplateService'
import { LibraryPowerPointConverter } from '../src/services/PowerPointConverter'
import type { JsonObject } from '../src/lib/shared/PowerpointTypes'

const POWERPOINT_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.presentationml.presentation'

let templates: SqliteTemplateRepository

beforeEach(() => {
  templates = new SqliteTemplateRepository(':memory:')
})

afterEach(() => {
  templates.close()
})

describe('export API', () => {
  it('normalizes compact canvas JSON and returns PowerPoint bytes', async () => {
    const app = createTestApp(1024 * 1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send(createCompactPresentation())
      .buffer(true)
      .parse((incoming, callback) => {
        const chunks: Buffer[] = []
        incoming.on('data', (chunk: Buffer) => chunks.push(chunk))
        incoming.on('end', () => callback(null, Buffer.concat(chunks)))
      })
      .expect(200)

    expect(response.headers).toMatchObject({
      'cache-control': 'no-store',
      'content-disposition': 'attachment; filename="export-api-deck.pptx"',
      'content-type': POWERPOINT_CONTENT_TYPE,
      'x-powerpoint-warning-count': '0',
    })
    expect(Buffer.isBuffer(response.body)).toBe(true)
    expect(response.body.subarray(0, 2).toString()).toBe('PK')

    const zip = await JSZip.loadAsync(response.body)
    const presentationXml = await zip.file('ppt/presentation.xml')?.async('text')
    const slideXml = await zip.file('ppt/slides/slide1.xml')?.async('text')
    const slideMasterXml = await zip.file('ppt/slideMasters/slideMaster1.xml')?.async('text')
    expect(presentationXml).toContain('<p:sldId')
    expect(slideXml).toContain('Exported by the API')
    expect(slideXml).toContain('prst="flowChartMagneticDisk"')
    expect(slideXml).not.toContain('prst="flowchartmagneticdisk"')
    expect(slideMasterXml).toContain('<p:sldLayoutId id="2147483649" r:id="rId1"/>')
  })

  it('rejects non-JSON request bodies', async () => {
    const app = createTestApp(1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'text/plain')
      .send('{}')
      .expect(415)

    expect(response.body.error.code).toBe('unsupported_media_type')
  })

  it('rejects compressed bodies before checking their media type', async () => {
    const app = createTestApp(1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'text/plain')
      .set('Content-Encoding', 'gzip')
      .send('{}')
      .expect(415)

    expect(response.body.error.code).toBe('unsupported_content_encoding')
  })

  it('rejects malformed JSON with a stable error', async () => {
    const app = createTestApp(1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send('{"presentation":')
      .expect(400)

    expect(response.body.error).toMatchObject({
      code: 'invalid_json',
      requestId: expect.any(String),
    })
  })

  it('rejects an empty JSON body before presentation normalization', async () => {
    const app = createTestApp(1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send()
      .expect(400)

    expect(response.body.error.code).toBe('invalid_json')
  })

  it('rejects presentation JSON that cannot be normalized', async () => {
    const app = createTestApp(1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send({ unrelated: true })
      .expect(422)

    expect(response.body.error.code).toBe('invalid_presentation_json')
  })

  it('parses valid primitive JSON before rejecting its presentation shape', async () => {
    const app = createTestApp(1024)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send('true')
      .expect(422)

    expect(response.body.error.code).toBe('invalid_presentation_json')
  })

  it('does not allow image paths or remote URLs to reach the server-side renderer', async () => {
    const app = createTestApp(4096)
    const presentation = createCompactPresentation('/etc/passwd')

    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send(presentation)
      .expect(422)

    expect(response.body.error.code).toBe('unsupported_image_source')
    expect(JSON.stringify(response.body)).not.toContain('/etc/passwd')
  })

  it('does not resolve another app\'s stored template assets during export', async () => {
    templates.ensureApp('app-one')
    templates.ensureApp('app-two')
    templates.insert(
      { templateId: 'scoped-template', templateJson: createStoredTemplate() },
      [{
        assetId: 'scoped-asset',
        bytes: Buffer.from([1, 2, 3]),
        contentType: 'image/png',
        templateId: 'scoped-template',
      }],
      undefined,
      undefined,
      'app-one',
    )

    const response = await request(createTestApp(4096))
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .set('X-App-Id', 'app-two')
      .send(createCompactPresentation('/import/scoped-template/assets/scoped-asset'))
      .expect(422)

    expect(response.body.error.code).toBe('template_asset_not_found')
  })

  it('enforces the configured JSON body limit', async () => {
    const app = createTestApp(64)
    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send(createCompactPresentation())
      .expect(413)

    expect(response.body.error.code).toBe('payload_too_large')
  })

  it('returns an empty 408 response when endpoint work exceeds the request timeout', async () => {
    const logs: ApiRequestLog[] = []
    const app = createApp({
      exportService: {
        export: async () => {
          await new Promise((resolve) => setTimeout(resolve, 50))
          return { bytes: new Uint8Array(), fileName: 'late.pptx', warnings: [] }
        },
        insert: async () => ({ bytes: new Uint8Array(), fileName: 'late.pptx', warnings: [] }),
      },
      importService: new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      logger: (entry) => logs.push(entry),
      maxExportJsonBytes: 1024 * 1024,
      maxUploadBytes: 1024,
      requestTimeoutMs: 10,
    })

    const response = await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send(createCompactPresentation())
      .expect(408)

    expect(response.headers['x-request-id']).toEqual(expect.any(String))
    expect(response.text).toBe('')
    expect(logs).toEqual([
      expect.objectContaining({
        errorCode: 'request_timeout',
        level: 'warn',
        outcome: 'failure',
        requestId: response.headers['x-request-id'],
        statusCode: 408,
      }),
    ])
  })

  it('returns the stable route-not-found envelope with a request ID', async () => {
    const response = await request(createTestApp(1024)).get('/missing').expect(404)

    expect(response.body.error).toMatchObject({
      code: 'route_not_found',
      requestId: expect.any(String),
    })
  })

  it('does not expose API routes without the versioned prefix', async () => {
    const app = createTestApp(1024)

    await request(app).post('/export').expect(404)
    await request(app).post('/api/v2/export').expect(404)
  })

  it('returns an empty 405 for a known path with the wrong HTTP method', async () => {
    const response = await request(createTestApp(1024)).get('/api/v1/export').expect(405)

    expect(response.headers['x-request-id']).toEqual(expect.any(String))
    expect(response.text).toBe('')
  })

  it('logs successful and failed requests with their response request IDs', async () => {
    const logs: ApiRequestLog[] = []
    const app = createTestApp(1024, (entry) => logs.push(entry))

    const succeeded = await request(app)
      .get('/api/v1/templates?appId=trace-test')
      .expect(200)
    const failed = await request(app)
      .post('/api/v1/export?appId=trace-test')
      .set('Content-Type', 'application/json')
      .send('{')
      .expect(400)

    expect(logs).toHaveLength(2)
    expect(logs[0]).toMatchObject({
      event: 'api_request_completed',
      level: 'info',
      method: 'GET',
      outcome: 'success',
      path: '/api/v1/templates',
      requestId: succeeded.headers['x-request-id'],
      statusCode: 200,
    })
    expect(logs[1]).toMatchObject({
      errorCode: 'invalid_json',
      errorName: 'ApiError',
      event: 'api_request_completed',
      level: 'warn',
      method: 'POST',
      outcome: 'failure',
      path: '/api/v1/export',
      requestId: failed.headers['x-request-id'],
      statusCode: 400,
    })
    expect(logs.every((entry) => Number.isFinite(entry.durationMs))).toBe(true)
    expect(logs.every((entry) => !JSON.stringify(entry).includes('trace-test'))).toBe(true)
  })

  it('logs unexpected failures without exposing the thrown error message', async () => {
    const logs: ApiRequestLog[] = []
    const app = createApp({
      exportService: {
        export: async () => {
          throw new TypeError('sensitive provider response')
        },
        insert: async () => ({ bytes: new Uint8Array(), fileName: 'unused.pptx', warnings: [] }),
      },
      importService: new ImportTemplateService(new LibraryPowerPointConverter(), templates),
      logger: (entry) => logs.push(entry),
      maxExportJsonBytes: 1024 * 1024,
      maxUploadBytes: 1024,
    })

    await request(app)
      .post('/api/v1/export')
      .set('Content-Type', 'application/json')
      .send(createCompactPresentation())
      .expect(500)

    expect(logs).toHaveLength(1)
    expect(logs[0]).toMatchObject({
      errorCode: 'internal_error',
      errorName: 'TypeError',
      level: 'error',
      outcome: 'failure',
      statusCode: 500,
    })
    expect(JSON.stringify(logs[0])).not.toContain('sensitive provider response')
  })

  it('inserts the generated slide into a bounded multipart target deck', async () => {
    const target = await createTargetPowerPoint()
    const response = await request(createTestApp(1024 * 1024))
      .post('/api/v1/export/insert')
      .field('presentation', JSON.stringify(createCompactPresentation()))
      .field('insertAfterSlide', '0')
      .attach('target', target, {
        contentType: POWERPOINT_CONTENT_TYPE,
        filename: 'target.pptx',
      })
      .buffer(true)
      .parse((incoming, callback) => {
        const chunks: Buffer[] = []
        incoming.on('data', (chunk: Buffer) => chunks.push(chunk))
        incoming.on('end', () => callback(null, Buffer.concat(chunks)))
      })
      .expect(200)

    expect(response.headers).toMatchObject({
      'content-disposition': 'attachment; filename="target-with-slide.pptx"',
      'content-type': POWERPOINT_CONTENT_TYPE,
    })
    const zip = await JSZip.loadAsync(response.body)
    expect(Object.keys(zip.files).filter((name) => /^ppt\/slides\/slide\d+\.xml$/u.test(name)))
      .toHaveLength(2)
    expect(await zip.file('ppt/slides/slide2.xml')?.async('text')).toContain('Exported by the API')
  })

  it('rejects incomplete and invalid insert multipart requests', async () => {
    const app = createTestApp(1024 * 1024)
    expect((await request(app)
      .post('/api/v1/export/insert')
      .field('presentation', JSON.stringify(createCompactPresentation()))
      .field('insertAfterSlide', '-1')
      .expect(400)).body.error.code).toBe('invalid_insert_position')

    expect((await request(app)
      .post('/api/v1/export/insert')
      .field('presentation', JSON.stringify(createCompactPresentation()))
      .field('insertAfterSlide', '0')
      .attach('target', Buffer.from('not a deck'), {
        contentType: POWERPOINT_CONTENT_TYPE,
        filename: 'target.pptx',
      })
      .expect(422)).body.error.code).toBe('invalid_powerpoint')
  })
})

function createTestApp(maxExportJsonBytes: number, logger?: (entry: ApiRequestLog) => void) {
  return createApp({
    exportService: new ExportPowerPointService(templates),
    importService: new ImportTemplateService(new LibraryPowerPointConverter(), templates),
    logger,
    maxExportJsonBytes,
    maxUploadBytes: Math.max(maxExportJsonBytes, 1024 * 1024),
  })
}

function createCompactPresentation(imageSource?: string) {
  const elements: JsonObject[] = [
    {
      id: 'title',
      type: 'text',
      x: 80,
      y: 80,
      w: 500,
      h: 80,
      text: 'Exported by the API',
      fontSize: 30,
      textColor: '070154',
    },
    {
      id: 'database',
      type: 'shape',
      shape: 'flowChartMagneticDisk',
      x: 80,
      y: 180,
      w: 120,
      h: 80,
      fill: 'E8EEF8',
      stroke: '070154',
      strokeWidth: 1,
    },
  ]
  if (imageSource) {
    elements.push({
      id: 'unsafe-image',
      type: 'image',
      x: 0,
      y: 0,
      w: 100,
      h: 100,
      src: imageSource,
    })
  }

  return {
    presentation: {
      title: 'Export API Deck',
      preserveElementOrder: true,
      showBranding: false,
      slides: [
        {
          id: 'slide-1',
          name: 'Exported slide',
          width: 1280,
          height: 720,
          backgroundColor: 'FFFFFF',
          elements,
        },
      ],
    },
  }
}

function createStoredTemplate(): PowerPointCanvasJson {
  return {
    presentation: {
      preserveElementOrder: true,
      showBranding: false,
      slides: [{
        backgroundColor: 'FFFFFF',
        elements: [],
        height: 720,
        id: 'slide-1',
        name: 'Stored template',
        width: 1280,
      }],
      title: 'Stored template',
    },
  }
}

async function createTargetPowerPoint() {
  const presentation = new PptxGenJS()
  presentation.layout = 'LAYOUT_WIDE'
  presentation.addSlide().addText('Existing target slide', {
    h: 1,
    w: 4,
    x: 1,
    y: 1,
  })
  const output = await presentation.write({ outputType: 'nodebuffer' })
  if (!Buffer.isBuffer(output)) {
    throw new Error('Expected a PowerPoint buffer.')
  }
  return output
}
