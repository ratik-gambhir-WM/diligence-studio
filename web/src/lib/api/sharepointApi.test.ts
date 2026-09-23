import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  downloadSharePointFile,
  resolveSharePointResource,
  SharePointApiError,
} from './sharepointApi'

afterEach(() => {
  vi.unstubAllGlobals()
})

describe('SharePoint API client', () => {
  it('resolves a SharePoint folder and validates its file records', async () => {
    const fetchMock = vi.fn(async () => jsonResponse({
      files: [{
        driveId: 'drive-1',
        fileId: 'file-1',
        lastModifiedDateTime: '2026-09-23T00:00:00Z',
        mimeType: 'application/pdf',
        name: 'diagram.pdf',
        path: 'Architecture',
        size: 123,
        webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/diagram.pdf',
      }],
      kind: 'folder',
      name: 'Architecture',
      resourceUrl: 'https://relentlessblue.sharepoint.com/:f:/s/TestSite/folder?e=abc',
      webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/Architecture',
    }))
    vi.stubGlobal('fetch', fetchMock)

    await expect(resolveSharePointResource('https://relentlessblue.sharepoint.com/:f:/s/TestSite/folder?e=abc'))
      .resolves.toMatchObject({ kind: 'folder', name: 'Architecture' })
    expect(fetchMock).toHaveBeenCalledWith('/api/v1/sharepoint/resolve', {
      body: JSON.stringify({ url: 'https://relentlessblue.sharepoint.com/:f:/s/TestSite/folder?e=abc' }),
      credentials: 'include',
      headers: { 'Content-Type': 'application/json' },
      method: 'POST',
      signal: undefined,
    })
  })

  it('downloads a file and uses the server-provided filename', async () => {
    const fetchMock = vi.fn(async () => new Response(new Uint8Array([1, 2, 3]), {
      headers: {
        'Content-Disposition': 'attachment; filename="diagram.pdf"',
        'Content-Type': 'application/pdf',
      },
    }))
    vi.stubGlobal('fetch', fetchMock)

    const file = await downloadSharePointFile({
      driveId: 'drive/one',
      fileId: 'file/one',
      lastModifiedDateTime: '',
      mimeType: 'application/pdf',
      name: 'fallback.pdf',
      path: '',
      size: 3,
      webUrl: 'https://relentlessblue.sharepoint.com/sites/TestSite/diagram.pdf',
    })

    expect(file.name).toBe('diagram.pdf')
    expect(file.type).toBe('application/pdf')
    expect(await file.arrayBuffer()).toEqual(new Uint8Array([1, 2, 3]).buffer)
    expect(fetchMock).toHaveBeenCalledWith(
      '/api/v1/sharepoint/files/drive%2Fone/file%2Fone/content',
      { credentials: 'include', signal: undefined },
    )
  })

  it('surfaces the sanitized server error for malformed responses', async () => {
    vi.stubGlobal('fetch', vi.fn(async () => jsonResponse({
      error: { code: 'sharepoint_access_denied', message: 'Access denied.' },
    }, { status: 403 })))

    await expect(resolveSharePointResource('https://relentlessblue.sharepoint.com/file'))
      .rejects.toBeInstanceOf(SharePointApiError)
    await expect(resolveSharePointResource('https://relentlessblue.sharepoint.com/file'))
      .rejects.toMatchObject({ code: 'sharepoint_access_denied', message: 'Access denied.' })
  })
})

function jsonResponse(body: unknown, init: ResponseInit = {}) {
  const headers = new Headers(init.headers)
  headers.set('Content-Type', 'application/json')
  return new Response(JSON.stringify(body), { ...init, headers })
}
