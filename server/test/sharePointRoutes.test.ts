// @vitest-environment node

import express from 'express'
import request from 'supertest'
import { describe, expect, it, vi } from 'vitest'

import { createSharePointRouter } from '../src/routes/sharePointRoutes'

describe('SharePoint routes', () => {
  it('resolves a resource with the authenticated Microsoft token', async () => {
    const service = {
      downloadFile: vi.fn(),
      resolveResource: vi.fn(async () => ({
        files: [],
        kind: 'folder' as const,
        name: 'Architecture',
        resourceUrl: 'https://relentlessblue.sharepoint.com/:f:/s/TestSite/folder?e=abc',
        webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/Architecture',
      })),
    }
    const app = createTestApp(service)

    const response = await request(app)
      .post('/api/v1/sharepoint/resolve')
      .set('Content-Type', 'application/json')
      .send({ url: 'https://relentlessblue.sharepoint.com/:f:/s/TestSite/folder?e=abc' })

    expect(response.status).toBe(200)
    expect(response.body.kind).toBe('folder')
    expect(service.resolveResource).toHaveBeenCalledWith(
      'https://relentlessblue.sharepoint.com/:f:/s/TestSite/folder?e=abc',
      'access-token',
    )
  })

  it('streams SharePoint file bytes with download headers', async () => {
    const service = {
      downloadFile: vi.fn(async () => ({
        bytes: Buffer.from([1, 2, 3]),
        contentType: 'application/pdf',
        fileName: 'diagram.pdf',
      })),
      resolveResource: vi.fn(),
    }
    const app = createTestApp(service)

    const response = await request(app)
      .get('/api/v1/sharepoint/files/drive-1/file-1/content')

    expect(response.status).toBe(200)
    expect(response.headers['content-type']).toContain('application/pdf')
    expect(response.headers['content-disposition']).toBe('attachment; filename="diagram.pdf"')
    expect([...response.body]).toEqual([1, 2, 3])
    expect(service.downloadFile).toHaveBeenCalledWith('drive-1', 'file-1', 'access-token')
  })
})

function createTestApp(service: Parameters<typeof createSharePointRouter>[0]) {
  const app = express()
  app.use(
    '/api/v1/sharepoint',
    createSharePointRouter(service, async () => 'access-token'),
  )
  return app
}
